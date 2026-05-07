"""Industrial vision anomaly detection using PatchCore.

This package provides a lightweight, CPU-friendly PatchCore implementation
plus a Gradio-based GUI suitable for training and inspecting parts on
industrial point-of-sale (TPV) workstations.
"""

from .model import PatchCore, PatchCoreConfig, InspectionResult

__all__ = ["PatchCore", "PatchCoreConfig", "InspectionResult"]
