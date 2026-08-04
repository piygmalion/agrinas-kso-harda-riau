# -*- coding: utf-8 -*-
"""Generate notebook visualisasi analisis aktor KSO Harda."""
import json
from pathlib import Path

NB = Path(__file__).with_name("Analisis_Aktor_KSO_Agrinas_Unit_II_Harda.ipynb")


def md(source: str):
    lines = source.split("\n")
    return {"cell_type": "markdown", "metadata": {}, "source": [ln + "\n" for ln in lines]}


def code(source: str):
    lines = source.split("\n")
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": [ln + "\n" for ln in lines],
    }


cells = []

cells.append(md("""# Analisis Aktor Jaringan KSO Agrinas — Visualisasi

**Unit II Harda — Ditreskrimum Polda Riau**  
**Sumber:** `jaringan_kso_agrinas.gexf`  
**Produk:** Notebook analisis visual siap briefing  

> Deskripsi–sintesis analitik. Bukan verifikasi lapangan / perintah operasi.

Notebook ini menampilkan metrik jaringan, skor hub risiko, peta jaringan, heatmap aktor–Polres, dan ego-network aktor prioritas."""))

cells.append(md("""## 1. Setup"""))

cells.append(code("""from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(r"C:\\Users\\Patron\\Downloads\\sawit lagi")
sys.path.insert(0, str(ROOT / "notebooks"))

from analisis_aktor_kso_harda import (
    UNIT, JUDUL, GEXF, OUT,
    load_graph, compute_metrics, score_hubs,
    actor_polres_matrix, ego_polres_summary, communities,
    write_findings, PAM_SEED,
)
from viz_aktor_kso import (
    plot_skor_hub, plot_betweenness, plot_degree_vs_betweenness,
    plot_heatmap_aktor_polres, plot_band_distribution, plot_ego_polres,
    plot_network, plot_ego_aktor, plot_dashboard,
)

%matplotlib inline
plt.rcParams["figure.dpi"] = 120
plt.rcParams["savefig.dpi"] = 160
pd.set_option("display.max_colwidth", 70)
pd.set_option("display.width", 140)

print(JUDUL)
print(UNIT)
print("GEXF:", GEXF, "| exists:", GEXF.exists())
print("Output:", OUT)"""))

cells.append(md("""## 2. Load graf & hitung metrik"""))

cells.append(code("""G = load_graph()
metrics = compute_metrics(G)
hubs = score_hubs(metrics)
mat = actor_polres_matrix(G)
ego = ego_polres_summary(G, metrics)
comm = communities(G)
if not comm.empty:
    metrics = metrics.merge(comm, on="aktor_id", how="left")

print(f"Nodes: {G.number_of_nodes()} | Edges: {G.number_of_edges()}")
print(f"Aktor (non-generik): {(metrics.bucket.isin(['aktor','agrinas']) & (metrics.generic_flag!='YA')).sum()}")
print(f"Watchlist prioritas: {(hubs.prioritas_pantau=='YA').sum()}")
hubs.head(12)"""))

cells.append(md("""## 3. Dashboard visual ringkas

Cuplikan satu layar untuk briefing Kanit / Kasubdit."""))

cells.append(code("""fig = plot_dashboard(G, hubs, metrics, mat, ego)
fig.savefig(OUT / "fig_dashboard.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()"""))

cells.append(md("""## 4. Ranking skor hub risiko

Skor = struktural jaringan + boost dampak pilot R4 (Majuma/PAB/UTS/Riden).  
Band: **KRITIS ≥70 · TINGGI ≥50 · SEDANG ≥30 · RENDAH <30**."""))

cells.append(code("""fig = plot_skor_hub(hubs, top_n=12)
fig.savefig(OUT / "fig_degree_hub.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()

watch = hubs[hubs["prioritas_pantau"] == "YA"][
    ["aktor_id", "skor_hub", "skor_struktural", "boost_dampak", "band_hub",
     "jml_polres_terkait", "jml_estate_tetangga", "pam_non_bujp_flag", "polres_list"]
]
watch"""))

cells.append(md("""## 5. Betweenness — aktor penjembatan

Betweenness tinggi = aktor yang sering berada di jalur terpendek antar node → indikasi potensi penularan risiko antar kebun/Polres."""))

cells.append(code("""fig = plot_betweenness(metrics, top_n=12)
fig.savefig(OUT / "fig_betweenness.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()"""))

cells.append(md("""## 6. Peta posisi: degree × betweenness

Ukuran titik ~ skor hub. Label ditampilkan untuk aktor skor ≥40 atau berflag PAM."""))

cells.append(code("""fig = plot_degree_vs_betweenness(hubs)
fig.savefig(OUT / "fig_scatter_degree_betweenness.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()"""))

cells.append(md("""## 7. Distribusi band & beban per Polres"""))

cells.append(code("""fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), facecolor="#F4F6F8")
plot_band_distribution(hubs, ax=axes[0])
plot_ego_polres(ego, ax=axes[1])
fig.tight_layout()
fig.savefig(OUT / "fig_band_dan_ego_polres.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()
ego"""))

cells.append(md("""## 8. Heatmap aktor × Polres

Titik hitam = ada keterkaitan. Baris dengan banyak kolom = aktor multi-lokasi."""))

cells.append(code("""fig = plot_heatmap_aktor_polres(mat)
fig.savefig(OUT / "fig_heatmap_aktor_polres.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()
mat.sort_values("total_polres", ascending=False).head(15)"""))

cells.append(md("""## 9. Peta jaringan fokus hub risiko

Node: Agrinas (pusat), aktor hub, estate, Polres.  
Warna mengikuti tipologi/klaster/band."""))

cells.append(code("""fig = plot_network(G, hubs, focus_hubs=True, seed=42)
fig.savefig(OUT / "fig_network_overview.png", bbox_inches="tight", facecolor=fig.get_facecolor())
plt.show()"""))

cells.append(md("""## 10. Ego-network empat pilot bentrok (R4)

Inspeksi tetangga langsung: Majuma · PAB · UTS · Riden Jaya."""))

cells.append(code("""PILOT = [
    "PT Nusantara Sawit Majuma",
    "PT Palma Agung Betuah (PAB)",
    "PT Ujung Tanjung Sejahtera",
    "PT Riden Jaya Konstruksi",
]

for fokus in PILOT:
    fig = plot_ego_aktor(G, fokus)
    safe = fokus.replace("/", "-").replace(" ", "_")[:40]
    fig.savefig(OUT / f"fig_ego_{safe}.png", bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.show()"""))

cells.append(md("""## 11. Drill-down aktor (ubah `FOKUS` sesuai kebutuhan briefing)"""))

cells.append(code("""FOKUS = "Bernas Mulya Mandiri"  # ganti: Majuma / PAB / UTS / Berlian / dll.

fig = plot_ego_aktor(G, FOKUS)
plt.show()

row = hubs[hubs["aktor_id"] == FOKUS]
print("=== PROFIL ===")
print(row.to_string(index=False) if len(row) else "Tidak ada di tabel hub")

rows = []
if FOKUS in G:
    for nb in sorted(G.neighbors(FOKUS)):
        d = G.nodes[nb]
        ed = G.edges[FOKUS, nb]
        rows.append({
            "tetangga": nb,
            "ntype": d.get("ntype"),
            "klaster": d.get("klaster"),
            "polres": d.get("polres") or (nb if str(d.get("ntype","")).startswith("polres") else ""),
            "relasi": ed.get("relasi"),
            "layer": ed.get("layer"),
        })
pd.DataFrame(rows)"""))

cells.append(md("""## 12. Ekspor artefak Harda"""))

cells.append(code("""OUT.mkdir(parents=True, exist_ok=True)
metrics.to_csv(OUT / "tabel_aktor_metrics.csv", index=False, encoding="utf-8-sig")
hubs.to_csv(OUT / "tabel_hub_risiko.csv", index=False, encoding="utf-8-sig")
ego.to_csv(OUT / "tabel_ego_polres.csv", index=False, encoding="utf-8-sig")
mat.to_csv(OUT / "matriks_aktor_polres.csv", encoding="utf-8-sig")
write_findings(hubs, metrics, ego, OUT / "ringkasan_temuan.md")

print("Artefak tersimpan di:", OUT)
for p in sorted(OUT.glob("*")):
    print(f" - {p.name} ({p.stat().st_size:,} byte)")

print("\\n--- RINGKASAN TEMUAN ---")
print((OUT / "ringkasan_temuan.md").read_text(encoding="utf-8")[:1600])"""))

cells.append(md("""## 13. Checklist baca visual untuk analis

| Visual | Pertanyaan yang dijawab |
|---|---|
| Skor hub | Siapa masuk watchlist bulan ini? |
| Betweenness | Siapa menjembatani banyak klaster? |
| Scatter degree×betweenness | Hub konektif vs hub jembatan? |
| Heatmap aktor×Polres | Siapa multi-lokasi? |
| Peta jaringan | Bagaimana Agrinas–KSO–estate–Polres terhubung? |
| Ego pilot | Tetangga langsung kasus bentrok? |

**Batasan:** graf dari cuplikan dokumen tersedia; skor = alat prioritisasi analitik, bukan bukti pidana.

---
*Unit II Harda · Ditreskrimum Polda Riau*"""))

nb = {
    "nbformat": 4,
    "nbformat_minor": 5,
    "metadata": {
        "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
        "language_info": {"name": "python", "version": "3.13.0"},
        "authors": [{"name": "Unit II Harda Ditreskrimum Polda Riau"}],
        "title": "Analisis Aktor KSO Agrinas — Visualisasi Unit II Harda",
    },
    "cells": cells,
}

for c in nb["cells"]:
    if c["source"] and c["source"][-1] == "\n":
        c["source"] = c["source"][:-1]

NB.write_text(json.dumps(nb, ensure_ascii=False, indent=1), encoding="utf-8")
print("OK:", NB)
