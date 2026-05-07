# Microsoft 365 (SharePoint & OneDrive) MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A modern Model Context Protocol (MCP) server that lets ChatGPT (and other MCP compatible clients) interact with SharePoint Online and OneDrive content using Microsoft Graph. The server is designed for environments where multiple Microsoft 365 accounts may exist on the same workstation: each session can authenticate with the signed-in Windows/Microsoft account that the user chooses at runtime.

> **Why this fork?** The original SharePoint-only implementation relied on application credentials tied to a specific tenant. This version embraces user-centric authentication (device code flow + MSAL cache) so the connector works with whichever account is selected on the machine—ideal for teams that share GPT licenses but want to access their own SharePoint/OneDrive resources.

## Feature Highlights

- 🔐 **Device-code authentication** – Pick the Microsoft account that is signed into the local machine instead of the ChatGPT account.
- 👥 **Multi-account aware** – List cached accounts, switch between them, and inspect the active context at any time.
- 🗂️ **Drive discovery** – Enumerate OneDrive personal drives and SharePoint document libraries.
- 📄 **Full file lifecycle** – Create folders, upload/update files (text or binary), download contents, and delete resources.
- 🔍 **Deep research** – Perform scoped searches inside a drive or global Microsoft Graph Search across SharePoint + OneDrive.
- 📚 **Site exploration** – Search for SharePoint sites and list their document libraries before activating one.

## MCP Tools Overview

| Tool | Purpose |
| --- | --- |
| `Start_Device_Login` / `Complete_Device_Login` | Initiate and complete the device-code flow for the desired Microsoft account. |
| `List_Available_Accounts` / `Set_Active_Account` | Inspect cached accounts and pick which one is active. |
| `Get_Auth_Context` / `Get_Graph_Context` | Debug helpers showing the current authentication + drive context. |
| `List_My_Drives` / `List_Site_Drives` / `Search_SharePoint_Sites` | Discover drives and sites available to the active account. |
| `Set_Active_Drive` | Choose the OneDrive or SharePoint library used for subsequent operations. |
| `List_Drive_Items`, `Get_Drive_Item_Metadata`, `Get_Drive_Item_Content` | Explore the selected drive and read file contents. |
| `Create_Drive_Folder`, `Upload_Drive_File`, `Update_Drive_File`, `Delete_Drive_Item` | Create, modify, and remove items. |
| `Search_Drive_Items` | Search inside the currently-selected drive. |
| `Deep_Search_Microsoft365` | Perform a Microsoft Graph Search across SharePoint and OneDrive resources for deep research scenarios. |

All tools automatically reuse the selected Microsoft account and drive context.

## Prerequisites

1. **Azure AD App Registration** – Register a *public* client application (no secret) in Azure AD / Entra ID.
   - Redirect URI can be `https://login.microsoftonline.com/common/oauth2/nativeclient`.
   - Required delegated permissions: `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`, and optionally others you need.
2. **Client ID** – Note the *Application (client) ID* of the registration.
3. **Python 3.10+** – A recent Python interpreter for running the MCP server.

## Configuration

The server is configured with environment variables (you can store them in a `.env` file during development):

| Variable | Description |
| --- | --- |
| `MCP_GRAPH_CLIENT_ID` | **Required.** Client ID of the Azure AD public application (fallback: `SHP_ID_APP`). |
| `MCP_GRAPH_TENANT_ID` | Optional. Tenant ID to restrict sign-ins. Defaults to `common` for multi-tenant. |
| `MCP_GRAPH_SCOPES` | Optional comma-separated scopes. Defaults to `Files.ReadWrite.All,Sites.ReadWrite.All,User.Read,offline_access`. |
| `MCP_GRAPH_CACHE_PATH` | Optional custom path for the MSAL token cache. Default: `~/.cache/mcp_sharepoint/token_cache.bin`. |
| `MCP_GRAPH_DEFAULT_DRIVE_ID` | Optional drive ID to auto-select on startup. |
| `MCP_GRAPH_LOG_LEVEL` / `MCP_GRAPH_LOG_FILE` | Optional logging settings. |

## Quickstart

### Installation

```bash
pip install -e .
```

### Running the server

```bash
python -m mcp_sharepoint
```

Integrate the binary into your MCP-aware client (ChatGPT desktop, Claude Desktop, MCP Inspector, etc.) by referencing the script `mcp-sharepoint` (created via the console script entry point) or by executing `python -m mcp_sharepoint` directly.

### First-time authentication workflow

1. Call `Start_Device_Login` from your MCP client.
2. Follow the `verification_uri` and `user_code` provided to authenticate with the Microsoft account currently logged into the workstation.
3. Call `Complete_Device_Login` with the returned `flow_id` to finish the sign-in.
4. (Optional) Use `List_Available_Accounts` and `Set_Active_Account` to switch between cached profiles.
5. Discover drives with `List_My_Drives` or `List_Site_Drives`, then select one using `Set_Active_Drive`.
6. You are ready to browse, search, and modify files using the rest of the tools.

## Development

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e .
```

Use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector) to debug traffic:

```bash
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint
```

## PatchCore Industrial Vision Tool

This repository also ships a self-contained PatchCore-based anomaly
detector (`patchcore_vision` package) for industrial inspection on
TPV / point-of-sale hardware. It supports both CPU and GPU and is
tuned to stay below 2000 ms per inspection on commodity CPUs.

### Install

```bash
pip install -e '.[vision]'
```

### Train and inspect from the CLI

```bash
# Train from a folder of "good" reference images
patchcore-vision train ./dataset/good -o model.pt --backbone resnet18 --device auto

# Inspect a single part on CPU
patchcore-vision inspect model.pt ./samples/part_42.jpg --device cpu

# Inspect a folder full of crops (e.g. ~100 1024x1024 squares from a 4K frame)
patchcore-vision inspect-batch model.pt ./samples/frame_001/ \
    --device cpu --batch-size 16 --list-defective

# Confirm the 2000 ms latency budget on the target machine for a full frame
patchcore-vision benchmark model.pt --device cpu \
    --num-crops 100 --crop-size 1024 --batch-size 16 --iterations 5
```

### Launch the easy-to-use GUI

```bash
patchcore-vision serve --host 0.0.0.0 --port 7860
```

The Gradio interface exposes a **Train** tab (drag-and-drop reference
images, choose backbone / device, save model) and an **Inspect** tab
(load model, drop image, see verdict + heat-map + per-frame latency).

### Python API

```python
from patchcore_vision import PatchCore, PatchCoreConfig

detector = PatchCore(PatchCoreConfig(backbone="resnet18", device="cpu"))
detector.fit(["good/01.png", "good/02.png", "good/03.png"])
detector.calibrate(["good/01.png", "good/02.png", "good/03.png"])

# Single image
result = detector.inspect("part.png")
print(result.score, result.is_defective, f"{result.elapsed_ms:.0f} ms")

# Full 4K frame: ~100 crops of 1024x1024 inspected in one pipeline pass
results = detector.inspect_many(crop_paths, batch_size=16, return_maps=False)
defects = [p for p, r in zip(crop_paths, results) if r.is_defective]

detector.save("model.pt")
```

### 4K cameras & ~100 crops per frame

The integration target — a 4K camera that emits ~100 1024×1024 crops
per inspection — is supported as a first-class scenario:

* The 1024×1024 crops are downsampled to the network input size
  (`image_size`) before the backbone forward, so the per-crop cost
  stays constant regardless of crop resolution.
* `inspect_many()` (and the `inspect-batch` CLI / "Inspect (batch)"
  GUI tab) forwards mini-batches of crops in a single call and
  computes the nearest-neighbour search against the memory bank as
  a single batched matmul, instead of looping in Python.
* The CLI / GUI defaults (and `PatchCoreConfig.tpv_fast()` from the
  Python API) ship the **TPV preset**: ResNet-18, `image_size=192`,
  `coreset_size=768`, `batch_size=8`. On a 4-thread industrial CPU
  this measured **~1.6 s** for a full 100×1024 frame, leaving
  comfortable headroom under the 2000 ms budget. On 8-thread
  hardware the same preset typically lands under 1 s.
* Always run `patchcore-vision benchmark model.pt --num-crops 100
  --crop-size 1024` on the *actual* TPV to certify the budget before
  deployment.

### Tuning levers

| Lever | Effect on latency | Effect on quality |
| --- | --- | --- |
| `image_size` 224 → 192 → 160 | ~30% / ~50% reduction | mild loss on fine defects |
| `coreset_size` 1024 → 768 → 512 | ~25% / ~50% reduction of NN matmul | mild loss on rare backgrounds |
| `batch_size` 16 → 8 | better L2 cache reuse on most CPUs | none |
| `backbone="resnet50"` | ~3× slower on CPU | better discrimination |
| `device="cuda"` for training | training ~10× faster | none (load with `device="cpu"` to deploy) |

### Latency tips for TPV CPUs

* Train on GPU (`device="cuda"` or `"auto"`) and deploy on CPU
  (`device="cpu"`) — the saved `.pt` works on both.
* If the 100×1024 budget is tight on weak hardware, drop to
  `PatchCoreConfig(image_size=160, coreset_size=512)`.
* For maximum spatial fidelity on a strong CPU (8+ threads),
  raise to `image_size=224, coreset_size=1024` — the upper end of
  what comfortably fits the budget.

## License

This project is released under the [MIT License](LICENSE).

