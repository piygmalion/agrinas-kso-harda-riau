# -*- coding: utf-8 -*-
"""Satu grafik representatif tabel prioritas Polres (Unit II Harda)."""

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd

BASE = Path(r"C:\Users\Patron\Downloads\sawit lagi")
OUT = BASE / "output" / "grafik_tabel_prioritas_polres"
DOCS = BASE / "docs" / "prioritas"
OUT.mkdir(parents=True, exist_ok=True)
DOCS.mkdir(parents=True, exist_ok=True)

ROWS = [
    ("Rohul", "P1", 28.0, 51473.11, 19243.34, 37.4, 21),
    ("Bengkalis", "P1", 26.5, 49415.52, 6708.69, 13.6, 8),
    ("Rohil", "P1", 20.0, 9412.55, 7424.55, 78.9, 24),
    ("Kampar", "P2", 17.0, 27936.70, 3622.59, 13.0, 4),
    ("Kuansing", "P2", 10.5, 30072.65, 15201.65, 50.5, 11),
    ("Dumai", "P2", 9.5, 3456.70, 3456.70, 100.0, 3),
    ("Inhu", "P2", 9.0, 37837.62, 29139.73, 77.0, 28),
    ("Pelalawan", "P2", 3.0, 6269.63, 333.01, 5.3, 9),
    ("Inhil", "P3", 5.5, 3968.36, 3890.36, 98.0, 15),
    ("Siak", "P3", 1.5, 1102.68, 985.44, 89.4, 6),
    ("Kep. Meranti", "P4", 1.5, 2584.44, 0.00, 0.0, 1),
    ("Pekanbaru", "P4", 1.5, 0.00, 0.00, np.nan, 0),
]

LEVEL_COLOR = {"P1": "#C0392B", "P2": "#D4A017", "P3": "#2E7D4F", "P4": "#7F8C8D"}
LEVEL_LABEL = {
    "P1": "P1 Sangat Tinggi",
    "P2": "P2 Tinggi",
    "P3": "P3 Sedang",
    "P4": "P4 Rendah",
}


def main():
    df = pd.DataFrame(
        ROWS,
        columns=["polres", "level", "skor", "sita", "kso", "pct", "kebun_intelkam"],
    )
    # urut tabel asli (prioritas operasional)
    df = df.reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(14.5, 8.2), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")

    y = np.arange(len(df))
    # bar sita (background volume)
    bars_sita = ax.barh(
        y, df["sita"], height=0.62, color="#0F2A44", alpha=0.90, label="Luas Sita (Ha)", zorder=2
    )
    # overlay KSO
    ax.barh(
        y, df["kso"], height=0.32, color="#C45C26", alpha=0.95, label="Luas KSO (Ha)", zorder=3
    )

    # skor sebagai marker di ujung kanan area plot (sumbu sekunder visual)
    # posisi x marker = max sita scale * (skor/max skor) * 0.22 offset from right? 
    # Better: twin axis on top for skor as scatter aligned to same y
    ax2 = ax.twiny()
    ax2.set_xlim(0, 32)
    ax2.set_xlabel("Skor Risiko →", color="#8B1E1E", fontsize=10, fontweight="bold")
    ax2.tick_params(axis="x", colors="#8B1E1E")
    for i, r in df.iterrows():
        ax2.scatter(
            r["skor"], i,
            s=180 + r["kebun_intelkam"] * 8,
            c=LEVEL_COLOR[r["level"]],
            edgecolors="white", linewidths=1.2,
            zorder=5, marker="o",
        )
        ax2.text(
            r["skor"] + 0.45, i, f'{r["skor"]:g}',
            va="center", ha="left", fontsize=8, color="#8B1E1E", fontweight="bold", zorder=6,
        )

    # labels di kiri = polres + level
    ylabels = [f'{r.polres}  ({r.level})' for r in df.itertuples()]
    ax.set_yticks(y)
    ax.set_yticklabels(ylabels, fontsize=10)
    ax.invert_yaxis()

    # annotations volume di kanan bar sita
    xmax = max(df["sita"].max(), 1)
    ax.set_xlim(0, xmax * 1.38)
    for i, r in df.iterrows():
        pct = f'{r["pct"]:.1f}%' if pd.notna(r["pct"]) else "—"
        ax.text(
            max(r["sita"], r["kso"]) + xmax * 0.015,
            i,
            f'Sita {r["sita"]:,.0f} · KSO {r["kso"]:,.0f} · {pct} · {int(r["kebun_intelkam"])} kebun',
            va="center", fontsize=7.5, color="#1E242A", zorder=4,
        )

    ax.set_xlabel("Luas lahan (Ha)  —  batang navy = Sita · batang oranye = KSO", fontsize=10)
    ax.set_title(
        "Prioritas Polres Agrinas–KSO: Volume Lahan & Skor Risiko\n"
        "Unit II Harda · Ditreskrimum Polda Riau",
        fontsize=14, fontweight="bold", color="#0F2A44", pad=14,
    )

    # footer totals
    total_sita = df["sita"].sum()
    total_kso = df["kso"].sum()
    total_skor = df["skor"].sum()
    ax.text(
        0.0, -0.12,
        f"Total Sita {total_sita:,.2f} Ha  ·  Total KSO {total_kso:,.2f} Ha ({total_kso/total_sita*100:.1f}%)  ·  "
        f"Total Skor {total_skor:g}  ·  Titik = skor risiko (warna level P1–P4; ukuran ≈ jml kebun Intelkam)",
        transform=ax.transAxes, fontsize=8.5, color="#5A6672",
    )

    handles = [
        mpatches.Patch(color="#0F2A44", label="Luas Sita (Ha)"),
        mpatches.Patch(color="#C45C26", label="Luas KSO (Ha)"),
    ] + [mpatches.Patch(color=LEVEL_COLOR[k], label=LEVEL_LABEL[k]) for k in ("P1", "P2", "P3", "P4")]
    ax.legend(handles=handles, loc="lower right", fontsize=8, framealpha=0.96, ncol=3)

    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.grid(axis="x", alpha=0.22, zorder=0)
    ax2.grid(False)

    fig.tight_layout()
    out1 = BASE / "Grafik_Prioritas_Polres_Satu_Tampilan.png"
    out2 = OUT / "fig_prioritas_polres_satu_tampilan.png"
    out3 = DOCS / "fig_prioritas_polres_satu_tampilan.png"
    fig.savefig(out1, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    fig.savefig(out2, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    fig.savefig(out3, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK:", out1)


if __name__ == "__main__":
    main()
