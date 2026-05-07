"""Command-line entry points for the PatchCore industrial vision tool."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .model import PatchCore, PatchCoreConfig, benchmark_batch, benchmark_inference


def _iter_images(folder: Path) -> list[Path]:
    suffixes = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}
    return sorted(p for p in folder.rglob("*") if p.suffix.lower() in suffixes)


def cmd_train(args: argparse.Namespace) -> int:
    folder = Path(args.good_dir).expanduser().resolve()
    images = _iter_images(folder)
    if not images:
        print(f"No images found under {folder}", file=sys.stderr)
        return 2

    config = PatchCoreConfig(
        backbone=args.backbone,
        image_size=args.image_size,
        coreset_size=args.coreset_size,
        device=args.device,
    )
    detector = PatchCore(config)

    print(f"Training PatchCore on {len(images)} images using {detector.device}…")
    stats = detector.fit(images, batch_size=args.batch_size)
    threshold = detector.calibrate(images, quantile=args.quantile, margin=args.margin)
    bench = benchmark_inference(detector, iterations=args.bench_iterations)

    out = Path(args.output).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    detector.save(out)

    report = {
        "model_path": str(out),
        "threshold": threshold,
        "training_stats": stats,
        "benchmark": bench,
        "meets_2000ms_budget": bench["p95_ms"] < 2000.0,
    }
    print(json.dumps(report, indent=2))
    return 0


def cmd_inspect(args: argparse.Namespace) -> int:
    detector = PatchCore.load(args.model, device=args.device)
    result = detector.inspect(args.image)
    print(
        json.dumps(
            {
                "score": result.score,
                "threshold": result.threshold,
                "is_defective": result.is_defective,
                "elapsed_ms": result.elapsed_ms,
                "device": str(detector.device),
            },
            indent=2,
        )
    )
    return 0 if not result.is_defective else 1


def cmd_inspect_batch(args: argparse.Namespace) -> int:
    folder = Path(args.crops_dir).expanduser().resolve()
    images = _iter_images(folder)
    if not images:
        print(f"No images found under {folder}", file=sys.stderr)
        return 2

    detector = PatchCore.load(args.model, device=args.device)
    import time as _time
    t0 = _time.perf_counter()
    results = detector.inspect_many(
        images, batch_size=args.batch_size, return_maps=False
    )
    total_ms = (_time.perf_counter() - t0) * 1000.0
    defective = [str(p) for p, r in zip(images, results) if r.is_defective]

    report = {
        "device": str(detector.device),
        "num_crops": len(images),
        "total_ms": round(total_ms, 1),
        "per_crop_ms_mean": round(total_ms / max(len(images), 1), 2),
        "meets_2000ms_budget": total_ms < 2000.0,
        "defective_count": len(defective),
        "defective_files": defective if args.list_defective else None,
    }
    print(json.dumps(report, indent=2))
    return 0 if not defective else 1


def cmd_benchmark(args: argparse.Namespace) -> int:
    detector = PatchCore.load(args.model, device=args.device)
    if args.num_crops > 1:
        report = benchmark_batch(
            detector,
            num_crops=args.num_crops,
            crop_size=args.crop_size,
            batch_size=args.batch_size,
            iterations=args.iterations,
        )
        ok = report["meets_2000ms_budget"]
    else:
        report = benchmark_inference(detector, iterations=args.iterations)
        report["meets_2000ms_budget"] = report["p95_ms"] < 2000.0
        ok = report["meets_2000ms_budget"]
    print(json.dumps(report, indent=2))
    return 0 if ok else 1


def cmd_serve(args: argparse.Namespace) -> int:
    from . import app as gradio_app

    demo = gradio_app.build_interface()
    demo.queue(max_size=8).launch(
        server_name=args.host,
        server_port=args.port,
        share=False,
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="patchcore-vision",
        description="Industrial vision (PatchCore) toolkit – train, inspect, benchmark, serve.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_train = sub.add_parser("train", help="Fit a PatchCore model on a folder of good images")
    p_train.add_argument("good_dir", help="Folder containing 'good' reference images")
    p_train.add_argument("-o", "--output", default="patchcore_model.pt")
    p_train.add_argument("--backbone", default="resnet18", choices=["resnet18", "resnet50"])
    p_train.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    p_train.add_argument("--image-size", type=int, default=192,
                         help="Network input size; default 192 fits the TPV 2000ms budget.")
    p_train.add_argument("--coreset-size", type=int, default=768,
                         help="Memory bank size; default 768 fits the TPV 2000ms budget.")
    p_train.add_argument("--batch-size", type=int, default=8)
    p_train.add_argument("--quantile", type=float, default=0.995)
    p_train.add_argument("--margin", type=float, default=1.10)
    p_train.add_argument("--bench-iterations", type=int, default=5)
    p_train.set_defaults(func=cmd_train)

    p_inspect = sub.add_parser("inspect", help="Run inspection on a single image")
    p_inspect.add_argument("model")
    p_inspect.add_argument("image")
    p_inspect.add_argument("--device", default="cpu", choices=["auto", "cpu", "cuda"])
    p_inspect.set_defaults(func=cmd_inspect)

    p_ibatch = sub.add_parser(
        "inspect-batch",
        help="Inspect every image in a folder (e.g. ~100 1024x1024 crops from a 4K frame)",
    )
    p_ibatch.add_argument("model")
    p_ibatch.add_argument("crops_dir")
    p_ibatch.add_argument("--device", default="cpu", choices=["auto", "cpu", "cuda"])
    p_ibatch.add_argument("--batch-size", type=int, default=8)
    p_ibatch.add_argument("--list-defective", action="store_true")
    p_ibatch.set_defaults(func=cmd_inspect_batch)

    p_bench = sub.add_parser(
        "benchmark",
        help="Measure inference latency. Use --num-crops 100 --crop-size 1024 for the 4K scenario.",
    )
    p_bench.add_argument("model")
    p_bench.add_argument("--device", default="cpu", choices=["auto", "cpu", "cuda"])
    p_bench.add_argument("--iterations", type=int, default=10)
    p_bench.add_argument("--num-crops", type=int, default=1,
                         help="Number of synthetic crops per inspection round")
    p_bench.add_argument("--crop-size", type=int, default=1024)
    p_bench.add_argument("--batch-size", type=int, default=8)
    p_bench.set_defaults(func=cmd_benchmark)

    p_serve = sub.add_parser("serve", help="Launch the Gradio GUI")
    p_serve.add_argument("--host", default="127.0.0.1")
    p_serve.add_argument("--port", type=int, default=7860)
    p_serve.set_defaults(func=cmd_serve)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
