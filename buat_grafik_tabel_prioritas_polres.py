# -*- coding: utf-8 -*-
"""Grafik dari tabel prioritas Polres (IDENTITAS + VOLUME KEBUN & LUAS).

Unit II Harda · Ditreskrimum Polda Riau
"""

from __future__ import annotations

import shutil
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

# Data dari tabel user (titik = ribuan, koma = desimal → float Python)
ROWS = [
    # polres, level, prioritas_harda, skor, kebun_matrik, kebun_intelkam, luas_sita, luas_kso, pct_kso
    ("Rohul", "P1", "Sangat tinggi", 28.0, 13, 21, 51473.11, 19243.34, 37.4),
    ("Bengkalis", "P1", "Sangat tinggi", 26.5, 10, 8, 49415.52, 6708.69, 13.6),
    ("Rohil", "P1", "Sangat tinggi", 20.0, 7, 24, 9412.55, 7424.55, 78.9),
    ("Kampar", "P2", "Tinggi", 17.0, 11, 4, 27936.70, 3622.59, 13.0),
    ("Kuansing", "P2", "Tinggi", 10.5, 11, 11, 30072.65, 15201.65, 50.5),
    ("Dumai", "P2", "Tinggi", 9.5, 3, 3, 3456.70, 3456.70, 100.0),
    ("Inhu", "P2", "Tinggi", 9.0, 28, 28, 37837.62, 29139.73, 77.0),
    ("Pelalawan", "P2", "Tinggi (dimensi TNTN)", 3.0, 9, 9, 6269.63, 333.01, 5.3),
    ("Inhil", "P3", "Sedang-tinggi", 5.5, 6, 15, 3968.36, 3890.36, 98.0),
    ("Siak", "P3", "Sedang", 1.5, 4, 6, 1102.68, 985.44, 89.4),
    ("Kep. Meranti", "P4", "Rendah (gap-fill)", 1.5, 1, 1, 2584.44, 0.00, 0.0),
    ("Pekanbaru", "P4", "Rendah (monitoring)", 1.5, 1, 0, 0.00, 0.00, np.nan),
]

LEVEL_COLOR = {
    "P1": "#C0392B",
    "P2": "#D4A017",
    "P3": "#2E7D4F",
    "P4": "#7F8C8D",
}

LEVEL_LABEL = {
    "P1": "P1 — Sangat Tinggi",
    "P2": "P2 — Tinggi",
    "P3": "P3 — Sedang",
    "P4": "P4 — Rendah",
}


def df_table() -> pd.DataFrame:
    df = pd.DataFrame(
        ROWS,
        columns=[
            "polres", "level", "prioritas_harda", "skor_risiko",
            "jml_kebun_matrik", "jml_kebun_intelkam",
            "luas_sita_ha", "luas_kso_ha", "pct_kso_vs_sita",
        ],
    )
    return df


def style(ax, title: str):
    ax.set_facecolor("#F4F6F8")
    ax.set_title(title, fontsize=12, fontweight="bold", color="#0F2A44", pad=10)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.grid(axis="x", alpha=0.25)


def legend_level(ax, loc="lower right"):
    handles = [mpatches.Patch(color=LEVEL_COLOR[k], label=LEVEL_LABEL[k]) for k in ("P1", "P2", "P3", "P4")]
    ax.legend(handles=handles, loc=loc, fontsize=8, framealpha=0.95)


def chart_skor(df: pd.DataFrame, path: Path):
    data = df.sort_values("skor_risiko", ascending=True)
    fig, ax = plt.subplots(figsize=(11, 6.5), facecolor="#F4F6F8")
    colors = [LEVEL_COLOR[l] for l in data["level"]]
    ax.barh(data["polres"], data["skor_risiko"], color=colors, height=0.7)
    for y, (skor, level) in enumerate(zip(data["skor_risiko"], data["level"])):
        ax.text(skor + 0.4, y, f"{skor:g}  ({level})", va="center", fontsize=8, color="#1E242A")
    ax.set_xlabel("Skor Risiko")
    ax.set_xlim(0, data["skor_risiko"].max() * 1.25)
    style(ax, "Skor Risiko per Polres\nUnit II Harda · Ditreskrimum Polda Riau")
    legend_level(ax)
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def chart_luas(df: pd.DataFrame, path: Path):
    data = df.sort_values("luas_sita_ha", ascending=True)
    y = np.arange(len(data))
    h = 0.38
    fig, ax = plt.subplots(figsize=(12, 7), facecolor="#F4F6F8")
    ax.barh(y + h / 2, data["luas_sita_ha"], h, label="Luas Sita (Ha)", color="#0F2A44")
    ax.barh(y - h / 2, data["luas_kso_ha"], h, label="Luas KSO (Ha)", color="#C45C26")
    ax.set_yticks(y)
    ax.set_yticklabels(data["polres"])
    for i, (sita, kso, pct) in enumerate(zip(data["luas_sita_ha"], data["luas_kso_ha"], data["pct_kso_vs_sita"])):
        ax.text(max(sita, kso) + 800, i, f"Sita {sita:,.0f} · KSO {kso:,.0f}" + (f" · {pct:.1f}%" if pd.notna(pct) else ""),
                va="center", fontsize=7, color="#1E242A")
    ax.set_xlabel("Luas (Ha)")
    ax.set_xlim(0, data["luas_sita_ha"].max() * 1.45)
    style(ax, "Luas Sita PKH vs Luas KSO per Polres")
    ax.legend(loc="lower right", fontsize=9)
    ax.grid(axis="x", alpha=0.25)
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def chart_pct_kso(df: pd.DataFrame, path: Path):
    data = df[df["pct_kso_vs_sita"].notna()].sort_values("pct_kso_vs_sita", ascending=True)
    fig, ax = plt.subplots(figsize=(11, 6.2), facecolor="#F4F6F8")
    colors = [LEVEL_COLOR[l] for l in data["level"]]
    ax.barh(data["polres"], data["pct_kso_vs_sita"], color=colors, height=0.7)
    for y, pct in enumerate(data["pct_kso_vs_sita"]):
        ax.text(pct + 1.2, y, f"{pct:.1f}%", va="center", fontsize=8)
    ax.axvline(40.3, color="#5A6672", ls="--", lw=1.2)
    ax.set_xlabel("% KSO vs Sita")
    ax.set_xlim(0, 115)
    style(ax, "% KSO terhadap Luas Sitaan PKH per Polres")
    handles = [mpatches.Patch(color=LEVEL_COLOR[k], label=LEVEL_LABEL[k]) for k in ("P1", "P2", "P3", "P4")]
    handles.append(plt.Line2D([0], [0], color="#5A6672", ls="--", label="Rata-rata Polda 40,3%"))
    ax.legend(handles=handles, loc="lower right", fontsize=8)
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def chart_kebun(df: pd.DataFrame, path: Path):
    data = df.sort_values("jml_kebun_intelkam", ascending=True)
    y = np.arange(len(data))
    h = 0.38
    fig, ax = plt.subplots(figsize=(11, 6.5), facecolor="#F4F6F8")
    ax.barh(y + h / 2, data["jml_kebun_matrik"], h, label="Jml Kebun (matrik)", color="#2C5F7C")
    ax.barh(y - h / 2, data["jml_kebun_intelkam"], h, label="Jml Kebun (Intelkam)", color="#B88A3D")
    ax.set_yticks(y)
    ax.set_yticklabels(data["polres"])
    for i, (a, b) in enumerate(zip(data["jml_kebun_matrik"], data["jml_kebun_intelkam"])):
        ax.text(max(a, b) + 0.4, i, f"{a} / {b}", va="center", fontsize=8)
    ax.set_xlabel("Jumlah kebun")
    ax.set_xlim(0, max(data["jml_kebun_matrik"].max(), data["jml_kebun_intelkam"].max()) * 1.25)
    style(ax, "Jumlah Kebun: Matrik vs Intelkam per Polres")
    ax.legend(loc="lower right", fontsize=9)
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def chart_dashboard(df: pd.DataFrame, path: Path):
    fig = plt.figure(figsize=(16, 11), facecolor="#F4F6F8")
    gs = fig.add_gridspec(2, 2, hspace=0.32, wspace=0.22)
    ax1 = fig.add_subplot(gs[0, 0])
    ax2 = fig.add_subplot(gs[0, 1])
    ax3 = fig.add_subplot(gs[1, 0])
    ax4 = fig.add_subplot(gs[1, 1])

    # 1 skor
    d1 = df.sort_values("skor_risiko", ascending=True)
    ax1.barh(d1["polres"], d1["skor_risiko"], color=[LEVEL_COLOR[l] for l in d1["level"]])
    style(ax1, "Skor Risiko")
    ax1.set_xlabel("Skor")

    # 2 luas
    d2 = df.sort_values("luas_sita_ha", ascending=True)
    y = np.arange(len(d2))
    ax2.barh(y + 0.18, d2["luas_sita_ha"], 0.35, color="#0F2A44", label="Sita")
    ax2.barh(y - 0.18, d2["luas_kso_ha"], 0.35, color="#C45C26", label="KSO")
    ax2.set_yticks(y)
    ax2.set_yticklabels(d2["polres"], fontsize=8)
    style(ax2, "Luas Sita vs KSO (Ha)")
    ax2.legend(fontsize=8, loc="lower right")

    # 3 % kso
    d3 = df[df["pct_kso_vs_sita"].notna()].sort_values("pct_kso_vs_sita", ascending=True)
    ax3.barh(d3["polres"], d3["pct_kso_vs_sita"], color=[LEVEL_COLOR[l] for l in d3["level"]])
    ax3.axvline(40.3, color="#5A6672", ls="--", lw=1)
    style(ax3, "% KSO vs Sita")
    ax3.set_xlabel("%")

    # 4 bubble: luas sita vs skor
    for level, g in df.groupby("level"):
        ax4.scatter(
            g["luas_sita_ha"], g["skor_risiko"],
            s=120 + g["jml_kebun_intelkam"] * 12,
            c=LEVEL_COLOR[level], label=LEVEL_LABEL[level],
            alpha=0.85, edgecolors="white", linewidths=0.8,
        )
        for _, r in g.iterrows():
            ax4.annotate(r["polres"], (r["luas_sita_ha"], r["skor_risiko"]),
                         textcoords="offset points", xytext=(5, 3), fontsize=7)
    style(ax4, "Posisi: Luas Sita × Skor Risiko\n(ukuran ≈ jml kebun Intelkam)")
    ax4.set_xlabel("Luas Sita (Ha)")
    ax4.set_ylabel("Skor Risiko")
    ax4.legend(fontsize=7, loc="upper left")
    ax4.grid(True, alpha=0.25)

    fig.suptitle(
        "Dashboard Tabel Prioritas Polres — Agrinas/KSO PKH\nUnit II Harda · Ditreskrimum Polda Riau",
        fontsize=14, fontweight="bold", color="#0F2A44", y=0.98,
    )
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def chart_priority_stacked_volume(df: pd.DataFrame, path: Path):
    """Ringkas volume per level prioritas."""
    g = df.groupby("level", sort=False).agg(
        luas_sita=("luas_sita_ha", "sum"),
        luas_kso=("luas_kso_ha", "sum"),
        skor=("skor_risiko", "sum"),
        n=("polres", "count"),
    ).reindex(["P1", "P2", "P3", "P4"])
    fig, axes = plt.subplots(1, 2, figsize=(12, 5), facecolor="#F4F6F8")
    x = np.arange(len(g))
    axes[0].bar(x, g["luas_sita"], color=[LEVEL_COLOR[i] for i in g.index], width=0.65)
    axes[0].set_xticks(x)
    axes[0].set_xticklabels([LEVEL_LABEL[i] for i in g.index], rotation=15, ha="right", fontsize=8)
    axes[0].set_ylabel("Ha")
    style(axes[0], "Total Luas Sita per Level Prioritas")
    for i, v in enumerate(g["luas_sita"]):
        axes[0].text(i, v + 1500, f"{v:,.0f}", ha="center", fontsize=8)

    axes[1].bar(x - 0.18, g["luas_sita"], 0.35, label="Sita", color="#0F2A44")
    axes[1].bar(x + 0.18, g["luas_kso"], 0.35, label="KSO", color="#C45C26")
    axes[1].set_xticks(x)
    axes[1].set_xticklabels(g.index, fontsize=10)
    axes[1].legend(fontsize=8)
    style(axes[1], "Sita vs KSO per Level Prioritas")
    fig.suptitle("Agregat Level Prioritas (P1–P4)", fontsize=13, fontweight="bold", color="#0F2A44")
    fig.tight_layout()
    fig.savefig(path, dpi=170, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print("OK", path.name)


def write_html_gallery(paths: list[Path]):
    cards = "\n".join(
        f'<a class="card" href="./{p.name}"><img src="./{p.name}" alt="{p.stem}"/><div>{p.stem.replace("_", " ")}</div></a>'
        for p in paths
    )
    html = f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8"/>
<title>Grafik Tabel Prioritas Polres — Unit II Harda</title>
<style>
body{{margin:0;font-family:Segoe UI,Calibri,sans-serif;background:#0F2A44;color:#fff}}
header{{padding:18px 24px;border-bottom:3px solid #C45C26}}
h1{{margin:0;font-size:1.2rem}} .sub{{color:#B8C5D0;font-size:.85rem;margin-top:4px}}
main{{padding:20px;display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px}}
.card{{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:12px;padding:10px;color:#fff;text-decoration:none}}
.card img{{width:100%;border-radius:8px;background:#F4F6F8}}
.card div{{margin-top:8px;font-size:.85rem;color:#D0D7DE}}
</style></head><body>
<header>
  <h1>Grafik Tabel Prioritas Polres</h1>
  <div class="sub">Unit II Harda · Ditreskrimum Polda Riau · IDENTITAS + VOLUME KEBUN & LUAS</div>
</header>
<main>{cards}</main>
</body></html>"""
    (DOCS / "index.html").write_text(html, encoding="utf-8")
    (OUT / "index.html").write_text(html, encoding="utf-8")


def update_portal():
    portal = BASE / "docs" / "index.html"
    if not portal.exists():
        return
    text = portal.read_text(encoding="utf-8")
    if "./prioritas/" in text:
        return
    card = """
      <a class="card" href="./prioritas/">
        <span class="tag">GRAFIK</span>
        <h2>Tabel Prioritas Polres</h2>
        <p>Skor risiko, luas sita vs KSO, % KSO, dan perbandingan kebun matrik–Intelkam.</p>
      </a>"""
    marker = 'href="./pengelola/"'
    idx = text.find(marker)
    if idx == -1:
        marker = 'href="./cluster/"'
        idx = text.find(marker)
    if idx != -1:
        end = text.find("</a>", idx)
        if end != -1:
            portal.write_text(text[: end + 4] + "\n" + card + text[end + 4 :], encoding="utf-8")
            print("OK portal + prioritas")


def main():
    df = df_table()
    df.to_csv(OUT / "tabel_prioritas_polres.csv", index=False, encoding="utf-8-sig")
    df.to_csv(DOCS / "tabel_prioritas_polres.csv", index=False, encoding="utf-8-sig")

    files = [
        OUT / "fig_skor_risiko_polres.png",
        OUT / "fig_luas_sita_vs_kso.png",
        OUT / "fig_pct_kso_vs_sita.png",
        OUT / "fig_kebun_matrik_vs_intelkam.png",
        OUT / "fig_dashboard_prioritas_polres.png",
        OUT / "fig_agregat_level_prioritas.png",
    ]
    chart_skor(df, files[0])
    chart_luas(df, files[1])
    chart_pct_kso(df, files[2])
    chart_kebun(df, files[3])
    chart_dashboard(df, files[4])
    chart_priority_stacked_volume(df, files[5])

    # also root-friendly copies
    for f in files:
        shutil.copy2(f, BASE / f.name.replace("fig_", "Grafik_Prioritas_"))
        shutil.copy2(f, DOCS / f.name)

    write_html_gallery([Path(p.name) for p in files])
    update_portal()

    print("\nTOTAL Sita:", f"{df['luas_sita_ha'].sum():,.2f} Ha")
    print("TOTAL KSO :", f"{df['luas_kso_ha'].sum():,.2f} Ha")
    print("TOTAL Skor:", df["skor_risiko"].sum())
    print("Output:", OUT)


if __name__ == "__main__":
    main()
