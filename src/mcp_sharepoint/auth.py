"""Authentication helpers for Microsoft Graph interactions."""

from __future__ import annotations

import json
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

import msal


class AuthenticationError(RuntimeError):
    """Represents an authentication failure."""


class AuthenticationFlowNotFound(AuthenticationError):
    """Raised when a referenced device flow can no longer be found."""


@dataclass
class AccountSummary:
    """Serializable representation of an MSAL account."""

    username: str
    name: Optional[str]
    home_account_id: str
    environment: Optional[str]
    tenant_profiles: Optional[Dict[str, Any]]
    is_active: bool = False


class GraphAuthManager:
    """Coordinate authentication for Microsoft Graph with multi-account support."""

    def __init__(
        self,
        *,
        client_id: str,
        tenant_id: str,
        scopes: List[str],
        cache_path: str,
        logger,
    ) -> None:
        if not client_id:
            raise ValueError("client_id is required to authenticate against Microsoft Graph")

        authority = f"https://login.microsoftonline.com/{tenant_id or 'common'}"
        self._logger = logger
        self._scopes = scopes
        self._cache_path = Path(cache_path)
        self._cache_lock = threading.Lock()
        self._pending_device_flows: Dict[str, Dict[str, Any]] = {}
        self._active_account_id: Optional[str] = None

        self._token_cache = msal.SerializableTokenCache()
        if self._cache_path.exists():
            try:
                cache_data = self._cache_path.read_text(encoding="utf-8")
                if cache_data:
                    self._token_cache.deserialize(cache_data)
            except OSError as exc:  # pragma: no cover - defensive
                self._logger.warning("Unable to read token cache: %s", exc)

        self._app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=self._token_cache,
        )

    # ------------------------------------------------------------------
    # Cache persistence helpers
    # ------------------------------------------------------------------
    def _save_cache(self) -> None:
        if self._token_cache.has_state_changed:
            with self._cache_lock:
                try:
                    self._cache_path.parent.mkdir(parents=True, exist_ok=True)
                    self._cache_path.write_text(self._token_cache.serialize(), encoding="utf-8")
                except OSError as exc:  # pragma: no cover - defensive
                    self._logger.warning("Unable to persist token cache: %s", exc)

    # ------------------------------------------------------------------
    # Account helpers
    # ------------------------------------------------------------------
    def _serialize_account(self, account: Optional[Dict[str, Any]]) -> Optional[AccountSummary]:
        if not account:
            return None
        return AccountSummary(
            username=account.get("username", ""),
            name=account.get("name"),
            home_account_id=account.get("home_account_id", ""),
            environment=account.get("environment"),
            tenant_profiles=account.get("tenant_profiles"),
            is_active=account.get("home_account_id") == self._active_account_id,
        )

    def list_accounts(self) -> List[AccountSummary]:
        """Return a list of cached accounts."""

        accounts = self._app.get_accounts()
        summaries = [self._serialize_account(account) for account in accounts]

        # Auto-select a single cached account if none is active
        if not self._active_account_id and len(accounts) == 1:
            self._active_account_id = accounts[0].get("home_account_id")
            if summaries and summaries[0]:
                summaries[0].is_active = True
        else:
            for summary in summaries:
                if summary:
                    summary.is_active = summary.home_account_id == self._active_account_id

        return [summary for summary in summaries if summary]

    def get_active_account(self) -> Optional[Dict[str, Any]]:
        if not self._active_account_id:
            return None
        for account in self._app.get_accounts():
            if account.get("home_account_id") == self._active_account_id:
                return account
        self._active_account_id = None
        return None

    def set_active_account(
        self,
        *,
        home_account_id: Optional[str] = None,
        username: Optional[str] = None,
    ) -> AccountSummary:
        """Set the account used for subsequent Graph operations."""

        if not home_account_id and not username:
            raise AuthenticationError("Either home_account_id or username must be provided")

        for account in self._app.get_accounts():
            if home_account_id and account.get("home_account_id") == home_account_id:
                self._active_account_id = home_account_id
                summary = self._serialize_account(account)
                if summary:
                    summary.is_active = True
                return summary
            if username and account.get("username", "").lower() == username.lower():
                self._active_account_id = account.get("home_account_id")
                summary = self._serialize_account(account)
                if summary:
                    summary.is_active = True
                return summary

        raise AuthenticationError("No cached account matches the provided identifier")

    def get_active_account_summary(self) -> Optional[AccountSummary]:
        return self._serialize_account(self.get_active_account())

    # ------------------------------------------------------------------
    # Token acquisition
    # ------------------------------------------------------------------
    def acquire_token_silent(self) -> str:
        """Acquire an access token for Microsoft Graph using the cached account."""

        account = self.get_active_account()
        if not account:
            accounts = self._app.get_accounts()
            if len(accounts) == 1:
                account = accounts[0]
                self._active_account_id = account.get("home_account_id")
            else:
                raise AuthenticationError(
                    "No active Microsoft account selected. Use the authentication tools to sign in "
                    "and select an account."
                )

        result = self._app.acquire_token_silent(self._scopes, account=account)
        if not result or "access_token" not in result:
            error_description = None
            if result:
                error_description = result.get("error_description") or json.dumps(result)
            raise AuthenticationError(
                "Unable to acquire a token silently. Complete the device login flow first." + (
                    f" Details: {error_description}" if error_description else ""
                )
            )

        self._save_cache()
        return result["access_token"]

    # ------------------------------------------------------------------
    # Device code flow
    # ------------------------------------------------------------------
    def start_device_login(self) -> Dict[str, Any]:
        """Initiate a device code login flow and return instructions for the user."""

        flow = self._app.initiate_device_flow(scopes=self._scopes)
        if "user_code" not in flow:
            raise AuthenticationError("Failed to start the device login flow")

        flow_id = str(uuid.uuid4())
        self._pending_device_flows[flow_id] = flow
        self._logger.info("Initiated device login flow %s for Microsoft Graph", flow_id)
        return {
            "flow_id": flow_id,
            "user_code": flow["user_code"],
            "verification_uri": flow.get("verification_uri") or flow.get("verification_uri_complete"),
            "expires_in": flow.get("expires_in"),
            "interval": flow.get("interval"),
            "message": flow.get("message"),
        }

    def complete_device_login(self, flow_id: str, timeout: Optional[int] = None) -> Dict[str, Any]:
        """Complete the device login flow previously initiated."""

        flow = self._pending_device_flows.pop(flow_id, None)
        if not flow:
            raise AuthenticationFlowNotFound(
                "The requested device flow does not exist or has already been completed."
            )

        result = self._app.acquire_token_by_device_flow(flow, timeout=timeout)
        if not result or "access_token" not in result:
            message = result.get("error_description") if isinstance(result, dict) else None
            raise AuthenticationError(
                "Device login did not succeed." + (f" Details: {message}" if message else "")
            )

        self._save_cache()

        preferred_username = None
        account_info = result.get("account")
        if account_info:
            preferred_username = account_info.get("username")
        if not preferred_username:
            claims = result.get("id_token_claims") or {}
            preferred_username = claims.get("preferred_username") or claims.get("email")

        selected_account = None
        if preferred_username:
            matching_accounts = self._app.get_accounts(username=preferred_username)
            if matching_accounts:
                selected_account = matching_accounts[0]
                self._active_account_id = selected_account.get("home_account_id")
        if not selected_account:
            accounts = self._app.get_accounts()
            if accounts:
                selected_account = accounts[-1]
                self._active_account_id = selected_account.get("home_account_id")

        summary = self._serialize_account(selected_account)
        if summary:
            summary.is_active = True

        self._logger.info("Device login completed for account %s", summary.username if summary else "unknown")

        return {
            "success": True,
            "account": summary.__dict__ if summary else None,
            "expires_in": result.get("expires_in"),
            "scope": result.get("scope"),
            "token_type": result.get("token_type"),
        }

    # ------------------------------------------------------------------
    # Misc helpers
    # ------------------------------------------------------------------
    def get_context(self) -> Dict[str, Any]:
        """Return a diagnostic snapshot of the authentication context."""

        active = self.get_active_account_summary()
        return {
            "active_account": active.__dict__ if active else None,
            "available_accounts": [summary.__dict__ for summary in self.list_accounts()],
            "scopes": self._scopes,
            "cache_path": str(self._cache_path),
        }

