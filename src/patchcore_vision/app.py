"""Gradio user interface for the PatchCore industrial vision tool.

The app exposes two workflows:

* **Train** – upload a folder of *good* reference images, choose backbone
  and device, and produce a portable ``.pt`` model file.
* **Inspect** – load a previously trained model and score a part image,
  visualising the heat-map and reporting latency in milliseconds so an
  operator can confirm the <2000 ms budget on a TPV machine.
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path

import numpy as np
import torch
from PIL import Image

try:
    import gradio as gr
except ImportError as exc:  # pragma: no cover - import-time guard
    raise SystemExit(
        "gradio is not installed. Install with: pip install '.[vision]'"
    ) from exc

from .model import PatchCore, PatchCoreConfig, benchmark_batch, benchmark_inference


_DEFAULT_MODEL_DIR = Path(os.environ.get("PATCHCORE_MODEL_DIR", "./patchcore_models"))


def _device_choices() -> list[str]:
    choices = ["auto", "cpu"]
    if torch.cuda.is_available():
        choices.append("cuda")
    return choices


def _heatmap_overlay(image: Image.Image, anomaly_map: np.ndarray, alpha: float = 0.55) -> Image.Image:
    """Blend a JET-style heat-map over the input image."""

    image = image.convert("RGB").resize((anomaly_map.shape[1], anomaly_map.shape[0]))
    array = np.asarray(image, dtype=np.float32) / 255.0

    a = anomaly_map.astype(np.float32)
    a_min, a_max = float(a.min()), float(a.max())
    if a_max - a_min < 1e-8:
        normalised = np.zeros_like(a)
    else:
        normalised = (a - a_min) / (a_max - a_min)

    # Simple JET approximation without pulling matplotlib at runtime.
    r = np.clip(1.5 - np.abs(4.0 * normalised - 3.0), 0.0, 1.0)
    g = np.clip(1.5 - np.abs(4.0 * normalised - 2.0), 0.0, 1.0)
    b = np.clip(1.5 - np.abs(4.0 * normalised - 1.0), 0.0, 1.0)
    heat = np.stack([r, g, b], axis=-1)

    blended = (1.0 - alpha) * array + alpha * heat
    blended = np.clip(blended * 255.0, 0, 255).astype(np.uint8)
    return Image.fromarray(blended)


def _train_handler(
    files: list,
    backbone: str,
    device: str,
    image_size: int,
    coreset_size: int,
    calibrate_quantile: float,
    margin: float,
    output_name: str,
    progress: "gr.Progress | None" = gr.Progress(track_tqdm=False),
):
    if not files:
        return "❌ Provide at least one reference image.", None, None

    paths = [Path(getattr(f, "name", f)) for f in files]
    config = PatchCoreConfig(
        backbone=backbone,
        image_size=int(image_size),
        coreset_size=int(coreset_size),
        device=device,
    )
    detector = PatchCore(config)

    def _cb(done, total):
        if progress is not None:
            progress(done / max(total, 1), desc=f"Extracting features {done}/{total}")

    train_start = time.perf_counter()
    stats = detector.fit(paths, batch_size=4, progress_cb=_cb)
    threshold = detector.calibrate(
        paths,
        quantile=float(calibrate_quantile),
        margin=float(margin),
    )
    elapsed = time.perf_counter() - train_start

    bench_single = benchmark_inference(detector, iterations=5)
    bench_batch = benchmark_batch(
        detector, num_crops=100, crop_size=1024, batch_size=8, iterations=2
    )

    _DEFAULT_MODEL_DIR.mkdir(parents=True, exist_ok=True)
    target = _DEFAULT_MODEL_DIR / (output_name or "patchcore_model.pt")
    detector.save(target)

    report = {
        "model_path": str(target),
        "threshold": threshold,
        "training_seconds": round(elapsed, 2),
        "training_stats": stats,
        "single_image_ms_p95": round(bench_single["p95_ms"], 1),
        "batch_100x1024_total_ms_max": round(bench_batch["total_ms_max"], 1),
        "batch_100x1024_per_crop_ms": round(bench_batch["per_crop_ms_mean"], 2),
        "device_used": str(detector.device),
        "meets_2000ms_single": bench_single["p95_ms"] < 2000.0,
        "meets_2000ms_batch_100x1024": bench_batch["meets_2000ms_budget"],
    }
    summary_lines = [
        f"Model saved to {target}",
        f"Threshold: {threshold:.4f}",
        f"Memory bank: {stats['memory_bank_size']} patches "
        f"x {stats['feature_dim']} dims",
        f"Single image latency (p95): {bench_single['p95_ms']:.1f} ms on {detector.device}",
        f"100x1024 batch total: {bench_batch['total_ms_mean']:.0f} ms "
        f"(max {bench_batch['total_ms_max']:.0f} ms) -> "
        f"{'OK <2000ms' if bench_batch['meets_2000ms_budget'] else 'OVER BUDGET'}",
    ]
    return "\n".join(summary_lines), report, str(target)


def _inspect_batch_handler(
    model_file,
    files: list,
    device: str,
    batch_size: int,
):
    if model_file is None:
        return "❌ Load a trained model first.", None, None
    if not files:
        return "❌ Drop one or more crops to inspect.", None, None

    paths = [Path(getattr(f, "name", f)) for f in files]
    detector = PatchCore.load(Path(getattr(model_file, "name", model_file)), device=device)

    t0 = time.perf_counter()
    results = detector.inspect_many(paths, batch_size=int(batch_size), return_maps=False)
    total_ms = (time.perf_counter() - t0) * 1000.0

    rows = [
        [
            p.name,
            f"{r.score:.4f}",
            f"{r.threshold:.4f}",
            "DEFECT" if r.is_defective else "OK",
        ]
        for p, r in zip(paths, results)
    ]
    summary = (
        f"Inspected {len(results)} crops in {total_ms:.0f} ms on {detector.device} "
        f"({total_ms / max(len(results), 1):.1f} ms/crop). "
        f"Defective: {sum(r.is_defective for r in results)}. "
        f"Within 2000 ms budget: {total_ms < 2000.0}"
    )
    payload = {
        "device": str(detector.device),
        "num_crops": len(results),
        "total_ms": total_ms,
        "per_crop_ms": total_ms / max(len(results), 1),
        "meets_2000ms_budget": total_ms < 2000.0,
        "defective_count": int(sum(r.is_defective for r in results)),
    }
    return summary, payload, rows


def _inspect_handler(
    model_file,
    image: Image.Image | None,
    device: str,
    threshold_override: float | None,
):
    if model_file is None:
        return None, "❌ Load a trained model first.", None
    if image is None:
        return None, "❌ Provide an image to inspect.", None

    model_path = Path(getattr(model_file, "name", model_file))
    detector = PatchCore.load(model_path, device=device)
    if threshold_override is not None and threshold_override > 0:
        detector.threshold = float(threshold_override)

    result = detector.inspect(image)
    overlay = _heatmap_overlay(image, result.anomaly_map)

    verdict = "DEFECT" if result.is_defective else "OK"
    summary = (
        f"Verdict: {verdict}\n"
        f"Score: {result.score:.4f} (threshold {result.threshold:.4f})\n"
        f"Latency: {result.elapsed_ms:.1f} ms on {detector.device}\n"
        f"Within 2000 ms budget: {result.elapsed_ms < 2000.0}"
    )
    payload = {
        "verdict": verdict,
        "score": result.score,
        "threshold": result.threshold,
        "elapsed_ms": result.elapsed_ms,
        "device": str(detector.device),
    }
    return overlay, summary, payload


def build_interface() -> "gr.Blocks":
    with gr.Blocks(title="PatchCore Industrial Vision") as demo:
        gr.Markdown(
            "# PatchCore Industrial Vision\n"
            "Train an anomaly detector from a handful of *good* images and "
            "inspect new parts on CPU or GPU. Designed to stay within a "
            "2000 ms per-frame budget on industrial TPV workstations."
        )

        with gr.Tab("Train"):
            with gr.Row():
                with gr.Column():
                    train_files = gr.File(
                        file_count="multiple",
                        file_types=["image"],
                        label="Reference 'good' images",
                    )
                    backbone = gr.Dropdown(
                        choices=["resnet18", "resnet50"],
                        value="resnet18",
                        label="Backbone (resnet18 = fastest CPU)",
                    )
                    device = gr.Dropdown(
                        choices=_device_choices(),
                        value="auto",
                        label="Device",
                    )
                    image_size = gr.Slider(
                        128, 512, value=192, step=16,
                        label="Network input size (192 = default TPV preset)",
                    )
                    coreset_size = gr.Slider(
                        128, 4096, value=768, step=64,
                        label="Coreset size (768 = default TPV preset)",
                    )
                    quantile = gr.Slider(
                        0.5, 1.0, value=0.995, step=0.005,
                        label="Calibration quantile",
                    )
                    margin = gr.Slider(
                        1.0, 2.0, value=1.10, step=0.01,
                        label="Threshold margin (multiplier)",
                    )
                    output_name = gr.Textbox(
                        value="patchcore_model.pt",
                        label="Output file name",
                    )
                    train_btn = gr.Button("Train", variant="primary")
                with gr.Column():
                    train_log = gr.Textbox(label="Status", lines=5)
                    train_report = gr.JSON(label="Report")
                    train_artifact = gr.File(label="Trained model file")
            train_btn.click(
                _train_handler,
                inputs=[
                    train_files,
                    backbone,
                    device,
                    image_size,
                    coreset_size,
                    quantile,
                    margin,
                    output_name,
                ],
                outputs=[train_log, train_report, train_artifact],
            )

        with gr.Tab("Inspect (batch)"):
            gr.Markdown(
                "Drop the ~100 crops produced from a 4K frame. They are "
                "all forwarded as one mini-batch through the network and "
                "matched against the memory bank in a single matmul, so "
                "the whole frame stays inside the 2000 ms budget."
            )
            with gr.Row():
                with gr.Column():
                    batch_model_file = gr.File(label="Trained .pt model", file_types=[".pt"])
                    batch_files = gr.File(
                        file_count="multiple",
                        file_types=["image"],
                        label="Crops (e.g. 100 x 1024x1024)",
                    )
                    batch_device = gr.Dropdown(
                        choices=_device_choices(), value="cpu", label="Inference device"
                    )
                    batch_size_in = gr.Slider(
                        1, 64, value=8, step=1,
                        label="Mini-batch size for the network (8 = optimal CPU cache)",
                    )
                    batch_btn = gr.Button("Inspect batch", variant="primary")
                with gr.Column():
                    batch_summary = gr.Textbox(label="Frame summary", lines=4)
                    batch_payload = gr.JSON(label="Payload")
                    batch_table = gr.Dataframe(
                        headers=["File", "Score", "Threshold", "Verdict"],
                        label="Per-crop results",
                        wrap=True,
                    )
            batch_btn.click(
                _inspect_batch_handler,
                inputs=[batch_model_file, batch_files, batch_device, batch_size_in],
                outputs=[batch_summary, batch_payload, batch_table],
            )

        with gr.Tab("Inspect"):
            with gr.Row():
                with gr.Column():
                    model_file = gr.File(label="Trained .pt model", file_types=[".pt"])
                    inspect_device = gr.Dropdown(
                        choices=_device_choices(),
                        value="cpu",
                        label="Inference device",
                    )
                    threshold_override = gr.Number(
                        value=0.0,
                        label="Threshold override (0 = use trained)",
                    )
                    image_input = gr.Image(type="pil", label="Part image")
                    inspect_btn = gr.Button("Inspect", variant="primary")
                with gr.Column():
                    overlay_out = gr.Image(label="Anomaly heat-map", type="pil")
                    summary_out = gr.Textbox(label="Verdict", lines=4)
                    payload_out = gr.JSON(label="Payload")
            inspect_btn.click(
                _inspect_handler,
                inputs=[model_file, image_input, inspect_device, threshold_override],
                outputs=[overlay_out, summary_out, payload_out],
            )

        gr.Markdown(
            "**Tip for 4K cameras**: with the default ResNet-18 + 224 px "
            "network input + coreset ≤ 1024, the *full frame* (≈100 crops "
            "of 1024×1024) is processed in a single batched pass and stays "
            "comfortably under 2000 ms on commodity CPU TPVs. The 1024×1024 "
            "crops are downsampled internally to the network input size — "
            "the heat-map is upsampled back at the end."
        )
    return demo


def main() -> None:
    """Entry point used by the ``patchcore-vision`` console script."""
    demo = build_interface()
    demo.queue(max_size=8).launch(
        server_name=os.environ.get("PATCHCORE_HOST", "127.0.0.1"),
        server_port=int(os.environ.get("PATCHCORE_PORT", "7860")),
        share=False,
        show_error=True,
    )


if __name__ == "__main__":  # pragma: no cover
    main()
