"""Draw bessel.svg and its PNG fallback from the integral representation of J_n.

uv run --extra dev examples/textbook/figures.py
"""

from pathlib import Path

import matplotlib
import numpy as np

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent


def bessel(n: int, x: np.ndarray) -> np.ndarray:
    """J_n(x) = (1/pi) * integral over [0, pi] of cos(n tau - x sin tau)."""
    tau = np.linspace(0, np.pi, 2001)
    return np.trapezoid(np.cos(n * tau[:, None] - x[None, :] * np.sin(tau[:, None])), tau, axis=0) / np.pi


# Text stays text in the SVG, so PowerPoint sets it in the deck's font.
plt.rcParams.update(
    {
        "svg.fonttype": "none",
        "font.family": "sans-serif",
        "font.sans-serif": ["Aptos", "Segoe UI", "DejaVu Sans"],
        "font.size": 14,
    }
)
x = np.linspace(0, 15, 600)
fig, ax = plt.subplots(figsize=(7, 3.6))
for n in range(3):
    ax.plot(x, bessel(n, x), label=f"n = {n}")
ax.axhline(0, color="0.6", lw=0.8)
ax.set(xlabel="x", ylabel="Jₙ(x)", xlim=(0, 15))
ax.legend(frameon=False)
fig.tight_layout()
fig.savefig(HERE / "bessel.svg", transparent=True)
fig.savefig(HERE / "bessel.png", dpi=200, transparent=True)
