"""PatchCore implementation tuned for industrial vision on CPU or GPU.

The implementation follows Roth et al. 2022 ("Towards Total Recall in
Industrial Anomaly Detection") with pragmatic simplifications so it can
train on a handful of "good" reference images and run inference in well
under two seconds on a typical industrial CPU.

Key design choices:

* A torchvision backbone (ResNet-18 by default) is used in eval mode.
  Layer 2 and layer 3 features are aggregated via a 3x3 average pool,
  upsampled to a common grid and concatenated to form patch features.
* A greedy k-center "coreset" selects a small memory bank from the
  training patches, keeping the nearest-neighbour search cheap at
  inference time.
* The anomaly map is the per-patch distance to the memory bank,
  bilinearly upsampled to the input resolution.

The :class:`PatchCore` object can be saved/loaded as a single ``.pt``
file containing both the configuration and the memory bank.
"""

from __future__ import annotations

import io
import json
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
import torch
import torch.nn.functional as F
from PIL import Image
from torchvision import transforms
from torchvision.models import (
    ResNet18_Weights,
    ResNet50_Weights,
    resnet18,
    resnet50,
)

_IMAGENET_MEAN = (0.485, 0.456, 0.406)
_IMAGENET_STD = (0.229, 0.224, 0.225)

_BACKBONES = {
    "resnet18": (resnet18, ResNet18_Weights.IMAGENET1K_V1),
    "resnet50": (resnet50, ResNet50_Weights.IMAGENET1K_V2),
}


@dataclass
class PatchCoreConfig:
    """Hyper-parameters for the PatchCore detector."""

    backbone: str = "resnet18"
    image_size: int = 224
    pool_size: int = 3
    coreset_size: int = 1024
    coreset_seed: int = 0
    device: str = "auto"  # "auto" | "cpu" | "cuda"
    smoothing_sigma: float = 4.0

    def resolve_device(self) -> torch.device:
        if self.device == "auto":
            return torch.device("cuda" if torch.cuda.is_available() else "cpu")
        return torch.device(self.device)


@dataclass
class InspectionResult:
    """Outcome of a single inspection."""

    score: float
    threshold: float
    is_defective: bool
    anomaly_map: np.ndarray  # (H, W) float32
    elapsed_ms: float
    extras: dict = field(default_factory=dict)


class _FeatureNet(torch.nn.Module):
    """Wraps a torchvision ResNet to expose layer2 and layer3 features."""

    def __init__(self, backbone: str) -> None:
        super().__init__()
        if backbone not in _BACKBONES:
            raise ValueError(
                f"Unsupported backbone '{backbone}'. Choose from {list(_BACKBONES)}"
            )
        ctor, weights = _BACKBONES[backbone]
        net = ctor(weights=weights)
        self.stem = torch.nn.Sequential(net.conv1, net.bn1, net.relu, net.maxpool)
        self.layer1 = net.layer1
        self.layer2 = net.layer2
        self.layer3 = net.layer3
        self.eval()
        for param in self.parameters():
            param.requires_grad_(False)

    def forward(self, x: torch.Tensor) -> tuple[torch.Tensor, torch.Tensor]:
        x = self.stem(x)
        x = self.layer1(x)
        f2 = self.layer2(x)
        f3 = self.layer3(f2)
        return f2, f3


def _aggregate(f2: torch.Tensor, f3: torch.Tensor, pool: int) -> torch.Tensor:
    pad = pool // 2
    f2 = F.avg_pool2d(f2, kernel_size=pool, stride=1, padding=pad)
    f3 = F.avg_pool2d(f3, kernel_size=pool, stride=1, padding=pad)
    f3 = F.interpolate(f3, size=f2.shape[-2:], mode="bilinear", align_corners=False)
    return torch.cat([f2, f3], dim=1)


def _greedy_coreset(features: torch.Tensor, target: int, seed: int = 0) -> torch.Tensor:
    """Greedy k-center subsampling on the rows of ``features``.

    Runs entirely on the device of ``features``. Memory cost is O(N).
    """

    n = features.shape[0]
    if target >= n:
        return torch.arange(n, device=features.device)

    generator = torch.Generator(device="cpu").manual_seed(seed)
    start = int(torch.randint(0, n, (1,), generator=generator).item())

    selected = torch.empty(target, dtype=torch.long, device=features.device)
    selected[0] = start
    diff = features - features[start : start + 1]
    min_dist = (diff * diff).sum(dim=1)

    for i in range(1, target):
        nxt = int(torch.argmax(min_dist).item())
        selected[i] = nxt
        diff = features - features[nxt : nxt + 1]
        new_dist = (diff * diff).sum(dim=1)
        torch.minimum(min_dist, new_dist, out=min_dist)
    return selected


def _gaussian_kernel(sigma: float, device: torch.device) -> torch.Tensor:
    radius = max(1, int(round(3.0 * sigma)))
    coords = torch.arange(-radius, radius + 1, dtype=torch.float32, device=device)
    kernel_1d = torch.exp(-(coords ** 2) / (2.0 * sigma * sigma))
    kernel_1d = kernel_1d / kernel_1d.sum()
    kernel_2d = kernel_1d[:, None] * kernel_1d[None, :]
    return kernel_2d


def _gaussian_blur(x: torch.Tensor, sigma: float) -> torch.Tensor:
    if sigma <= 0:
        return x
    kernel = _gaussian_kernel(sigma, x.device)
    kernel = kernel.expand(1, 1, *kernel.shape)
    pad = kernel.shape[-1] // 2
    x = F.pad(x, (pad, pad, pad, pad), mode="reflect")
    return F.conv2d(x, kernel)


class PatchCore:
    """Train and run a PatchCore anomaly detector.

    Typical usage::

        detector = PatchCore(PatchCoreConfig(backbone="resnet18"))
        detector.fit(["good/01.png", "good/02.png", ...])
        detector.calibrate(["good/01.png", ...])  # set threshold
        result = detector.inspect("part.png")
        print(result.score, result.is_defective)
        detector.save("model.pt")
    """

    def __init__(self, config: PatchCoreConfig | None = None) -> None:
        self.config = config or PatchCoreConfig()
        self._device = self.config.resolve_device()
        self._net = _FeatureNet(self.config.backbone).to(self._device)
        self._transform = transforms.Compose(
            [
                transforms.Resize(
                    (self.config.image_size, self.config.image_size),
                    interpolation=transforms.InterpolationMode.BILINEAR,
                ),
                transforms.ToTensor(),
                transforms.Normalize(_IMAGENET_MEAN, _IMAGENET_STD),
            ]
        )
        self._memory_bank: torch.Tensor | None = None
        self._feature_grid: tuple[int, int] | None = None
        self.threshold: float = float("inf")
        self.train_stats: dict = {}

    # ------------------------------------------------------------------ utils
    @property
    def device(self) -> torch.device:
        return self._device

    def to(self, device: str | torch.device) -> "PatchCore":
        self._device = torch.device(device)
        self._net.to(self._device)
        if self._memory_bank is not None:
            self._memory_bank = self._memory_bank.to(self._device)
        return self

    def _load_image(self, source) -> torch.Tensor:
        if isinstance(source, (str, Path)):
            image = Image.open(source).convert("RGB")
        elif isinstance(source, Image.Image):
            image = source.convert("RGB")
        elif isinstance(source, np.ndarray):
            image = Image.fromarray(source).convert("RGB")
        else:
            raise TypeError(f"Unsupported image source: {type(source)!r}")
        return self._transform(image)

    def _batch(self, sources: Iterable, batch_size: int) -> Iterable[torch.Tensor]:
        buffer: list[torch.Tensor] = []
        for src in sources:
            buffer.append(self._load_image(src))
            if len(buffer) >= batch_size:
                yield torch.stack(buffer).to(self._device)
                buffer = []
        if buffer:
            yield torch.stack(buffer).to(self._device)

    def _extract_patches(self, batch: torch.Tensor) -> torch.Tensor:
        f2, f3 = self._net(batch)
        feats = _aggregate(f2, f3, self.config.pool_size)
        b, c, h, w = feats.shape
        self._feature_grid = (h, w)
        return feats.permute(0, 2, 3, 1).reshape(b * h * w, c)

    # --------------------------------------------------------------- training
    def fit(
        self,
        sources: Sequence,
        batch_size: int = 8,
        progress_cb=None,
    ) -> dict:
        """Fit the memory bank from a sequence of *good* images."""

        if not sources:
            raise ValueError("fit() requires at least one reference image")

        start = time.perf_counter()
        all_features: list[torch.Tensor] = []
        seen = 0
        with torch.inference_mode():
            for batch in self._batch(sources, batch_size):
                feats = self._extract_patches(batch)
                all_features.append(feats.cpu())
                seen += batch.shape[0]
                if progress_cb is not None:
                    progress_cb(seen, len(sources))
        bank = torch.cat(all_features, dim=0)

        target = min(self.config.coreset_size, bank.shape[0])
        # Coreset selection on CPU keeps memory predictable on large fits.
        idx = _greedy_coreset(bank, target=target, seed=self.config.coreset_seed)
        self._memory_bank = bank[idx].contiguous().to(self._device)

        elapsed = time.perf_counter() - start
        self.train_stats = {
            "num_images": seen,
            "patches_total": int(bank.shape[0]),
            "memory_bank_size": int(self._memory_bank.shape[0]),
            "feature_dim": int(self._memory_bank.shape[1]),
            "feature_grid": self._feature_grid,
            "elapsed_seconds": float(elapsed),
            "device": str(self._device),
        }
        return self.train_stats

    def calibrate(
        self,
        sources: Sequence,
        quantile: float = 0.995,
        margin: float = 1.10,
    ) -> float:
        """Set ``self.threshold`` from the score distribution on good images.

        The threshold defaults to ``quantile`` of the scores multiplied by
        ``margin`` so a small operating cushion is left above the worst
        observed *good* sample.
        """

        if self._memory_bank is None:
            raise RuntimeError("PatchCore must be fitted before calibration.")
        if not sources:
            raise ValueError("calibrate() requires at least one reference image")

        scores = [self._inspect_tensor(self._load_image(s).unsqueeze(0)).score
                  for s in sources]
        scores_arr = np.asarray(scores, dtype=np.float32)
        cutoff = float(np.quantile(scores_arr, quantile)) if len(scores_arr) > 1 else float(scores_arr[0])
        self.threshold = float(cutoff * margin)
        return self.threshold

    # ------------------------------------------------------------------ infer
    def _inspect_tensor(self, batch: torch.Tensor) -> InspectionResult:
        if self._memory_bank is None:
            raise RuntimeError("PatchCore must be fitted before inspection.")
        if batch.shape[0] != 1:
            raise ValueError("inspect() processes a single image at a time")

        start = time.perf_counter()
        with torch.inference_mode():
            patches = self._extract_patches(batch)  # (HW, C)
            # Squared L2 distance via expanded form (memory-efficient).
            bank = self._memory_bank
            bank_sq = (bank * bank).sum(dim=1)
            patch_sq = (patches * patches).sum(dim=1, keepdim=True)
            dists = patch_sq + bank_sq.unsqueeze(0) - 2.0 * patches @ bank.T
            dists.clamp_(min=0.0)
            min_dist, _ = dists.min(dim=1)
            min_dist = torch.sqrt(min_dist + 1e-12)

            h, w = self._feature_grid  # type: ignore[misc]
            anomaly_map = min_dist.view(1, 1, h, w)
            anomaly_map = F.interpolate(
                anomaly_map,
                size=(self.config.image_size, self.config.image_size),
                mode="bilinear",
                align_corners=False,
            )
            anomaly_map = _gaussian_blur(anomaly_map, self.config.smoothing_sigma)
            anomaly_map = anomaly_map.squeeze().detach().cpu().numpy().astype(np.float32)
            score = float(anomaly_map.max())

        elapsed_ms = (time.perf_counter() - start) * 1000.0
        return InspectionResult(
            score=score,
            threshold=self.threshold,
            is_defective=score >= self.threshold,
            anomaly_map=anomaly_map,
            elapsed_ms=elapsed_ms,
        )

    def inspect(self, source) -> InspectionResult:
        """Score a single image. Accepts a path, PIL image or ndarray."""
        tensor = self._load_image(source).unsqueeze(0).to(self._device)
        return self._inspect_tensor(tensor)

    # ----------------------------------------------------------- persistence
    def save(self, path: str | Path) -> Path:
        if self._memory_bank is None:
            raise RuntimeError("Nothing to save: model has not been fitted.")
        target = Path(path)
        payload = {
            "version": 1,
            "config": asdict(self.config),
            "threshold": self.threshold,
            "feature_grid": self._feature_grid,
            "memory_bank": self._memory_bank.detach().cpu(),
            "train_stats": self.train_stats,
        }
        torch.save(payload, target)
        return target

    @classmethod
    def load(cls, path: str | Path, device: str | None = None) -> "PatchCore":
        payload = torch.load(path, map_location="cpu", weights_only=False)
        cfg_dict = dict(payload["config"])
        if device is not None:
            cfg_dict["device"] = device
        config = PatchCoreConfig(**cfg_dict)
        instance = cls(config)
        instance._memory_bank = payload["memory_bank"].to(instance.device)
        instance._feature_grid = tuple(payload["feature_grid"])  # type: ignore[arg-type]
        instance.threshold = float(payload.get("threshold", float("inf")))
        instance.train_stats = dict(payload.get("train_stats", {}))
        return instance

    # ---------------------------------------------------------- introspection
    def summary(self) -> str:
        memory = (
            self._memory_bank.shape if self._memory_bank is not None else "uninitialised"
        )
        return json.dumps(
            {
                "config": asdict(self.config),
                "device": str(self._device),
                "threshold": self.threshold,
                "memory_bank": str(memory),
                "feature_grid": self._feature_grid,
                "train_stats": self.train_stats,
            },
            indent=2,
        )


def benchmark_inference(
    detector: PatchCore,
    iterations: int = 5,
    image_size: int | None = None,
) -> dict:
    """Run a synthetic inference benchmark for the configured device."""

    if detector._memory_bank is None:
        raise RuntimeError("Detector must be fitted before benchmarking.")
    size = image_size or detector.config.image_size
    rng = np.random.default_rng(0)
    fake = (rng.random((size, size, 3)) * 255).astype(np.uint8)
    detector.inspect(fake)  # warm-up
    timings = []
    for _ in range(iterations):
        result = detector.inspect(fake)
        timings.append(result.elapsed_ms)
    return {
        "iterations": iterations,
        "device": str(detector.device),
        "image_size": size,
        "mean_ms": float(np.mean(timings)),
        "p50_ms": float(np.median(timings)),
        "p95_ms": float(np.quantile(timings, 0.95)),
        "max_ms": float(np.max(timings)),
    }
