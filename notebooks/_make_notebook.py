# -*- coding: utf-8 -*-
"""Generate .ipynb tanpa dependensi nbformat."""
import json
from pathlib import Path

NB = Path(__file__).with_name("Analisis_Aktor_KSO_Agrinas_Unit_II_Harda.ipynb")


def md(source: str):
    return {"cell_type": "markdown", "metadata": {}, "source": [line + "\n" for line in source.split("\n")]}


def code(source: str):
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": [line + "\n" for line in source.split("\n")],
    }


cells = []

cells.append(md("""# Analisis Aktor Jaringan KSO Agrinas

**Unit II Harda — Ditreskrimum Polda Riau**  
**Produk:** Notebook analitik operasional (siap pakai)  
**Sumber graf:** `jaringan_kso_agrinas.gexf`  
**Tanggal kerangka:** 4 Agustus 2026  

> Sifat dokumen: **deskripsi–sintesis analitik**. Bukan verifikasi lapangan dan bukan perintah operasi.

### Tujuan notebook
1. Memetakan aktor penerima KSO / PAM dalam skema Agrinas.
2. Menghitung metrik jaringan (degree, betweenness, dll.).
3. Menyusun **skor hub risiko** untuk watchlist Unit II Harda.
4. Menyajikan ego-network per Polres dan matriks aktor–Polres.
5. Mengekspor tabel & gambar siap briefing."""))

cells.append(md("""## 0. Cara pakai (Unit Harda)

1. Pastikan file berikut ada di folder kerja `sawit lagi/`:
   - `jaringan_kso_agrinas.gexf`
   - `analisis_lanjutan_5_prioritas.xlsx` (opsional, untuk catatan R2)
2. Jalankan semua cell berurutan (**Run All**).
3. Hasil otomatis tersimpan di `output/aktor_kso/`.
4. Untuk refresh bulanan: regenerate GEXF dari pipeline graph, lalu Run All ulang.

**Alternatif CLI (tanpa UI notebook):**
```bash
python notebooks/analisis_aktor_kso_harda.py
```"""))

cells.append(md("""## 1. Setup & identitas analisis"""))

cells.append(code("""from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd

# pastikan package lokal notebook terdeteksi
ROOT = Path(r"C:\\Users\\Patron\\Downloads\\sawit lagi")
sys.path.insert(0, str(ROOT / "notebooks"))

from analisis_aktor_kso_harda import (
    UNIT, JUDUL, TANGGAL, BASE, GEXF, OUT,
    load_graph, compute_metrics, score_hubs,
    actor_polres_matrix, ego_polres_summary, communities,
    fig_degree_hub, fig_betweenness, fig_network_overview,
    write_findings, run_analysis, PAM_SEED, GENERIC,
    W_DEGREE, W_BETWEEN, W_MULTI_POLRES, W_PAM_FLAG, W_ESTATE_MERAH,
)

pd.set_option("display.max_columns", 50)
pd.set_option("display.max_colwidth", 60)
pd.set_option("display.width", 140)

print(JUDUL)
print(UNIT)
print("Tanggal kerangka :", TANGGAL)
print("Sumber GEXF      :", GEXF)
print("Folder output    :", OUT)
print("GEXF exists      :", GEXF.exists())"""))

cells.append(md("""## 2. Load graf & audit struktur

Graf multipartite longgar:
- **Agrinas** (pusat skema)
- **Aktor** (penerima KSO / PAM / multi-lokasi)
- **Estate** (kebun; warna klaster merah/kuning/hijau)
- **Polres** (wilayah; band risiko R1)"""))

cells.append(code("""import networkx as nx

G = load_graph()

n_aktor = sum(1 for _, d in G.nodes(data=True) if d.get("ntype") == "agrinas" or str(d.get("ntype","")).startswith("aktor"))
n_estate = sum(1 for _, d in G.nodes(data=True) if str(d.get("ntype","")).startswith("estate"))
n_polres = sum(1 for _, d in G.nodes(data=True) if str(d.get("ntype","")).startswith("polres"))
n_hub = sum(1 for _, d in G.nodes(data=True) if d.get("hub_risiko"))

audit = pd.DataFrame([
    {"metrik": "nodes", "nilai": G.number_of_nodes()},
    {"metrik": "edges", "nilai": G.number_of_edges()},
    {"metrik": "aktor (+Agrinas)", "nilai": n_aktor},
    {"metrik": "estate", "nilai": n_estate},
    {"metrik": "polres", "nilai": n_polres},
    {"metrik": "flag hub_risiko", "nilai": n_hub},
    {"metrik": "komponen terhubung", "nilai": nx.number_connected_components(G)},
    {"metrik": "density", "nilai": round(nx.density(G), 4)},
])
audit"""))

cells.append(md("""### 2.1 Catatan kualitas data
- Label generik **`KSO`** dan **`Belum ada KSO`** = gap isian satker, **bukan** badan hukum.
- Atribut `hub_risiko` berasal dari seed R2 (multi-lokasi / PAM bentrok).
- Estate tanpa klaster tetap dihitung, tetapi tidak menambah skor “merah”."""))

cells.append(code("""# Node generik / gap
gap_nodes = [n for n in G.nodes if n in GENERIC]
print("Node gap/generik:", gap_nodes)

# Distribusi ntype
ntype_count = (
    pd.Series({n: d.get("ntype") for n, d in G.nodes(data=True)})
    .value_counts()
    .rename_axis("ntype")
    .reset_index(name="jumlah")
)
ntype_count"""))

cells.append(md("""## 3. Metrik jaringan per node

| Metrik | Arti analitik Harda |
|---|---|
| **Degree** | Banyak koneksi langsung (kebun/Polres/aktor) |
| **Betweenness** | Peran “jembatan” — potensi penularan risiko lintas klaster |
| **Closeness** | Kedekatan rata-rata ke seluruh jaringan |
| **Eigenvector** | Terhubung ke node yang juga sentral |"""))

cells.append(code("""metrics = compute_metrics(G)
print("Baris metrik:", len(metrics))
metrics.head(15)"""))

cells.append(code("""# Fokus aktor saja (tanpa estate/polres), buang label generik
aktor_m = metrics[(metrics["bucket"].isin(["aktor", "agrinas"])) & (metrics["generic_flag"] != "YA")].copy()
aktor_m.sort_values("degree", ascending=False)[
    ["aktor_id", "ntype", "degree", "betweenness", "closeness", "eigenvector",
     "jml_polres_terkait", "jml_estate_tetangga", "jml_estate_merah_tetangga", "pam_non_bujp_flag"]
].head(20)"""))

cells.append(md("""## 4. Skor hub risiko Unit II Harda

$$
\\text{Skor} = 100 \\times (0.25\\hat{D} + 0.25\\hat{B} + 0.20\\hat{S} + 0.15\\,PAM + 0.10\\hat{R} + 0.05\\,U)
$$

- $\\hat{D}$ degree ternormalisasi **antar penerima KSO** (Agrinas tidak ikut menormalisasi)  
- $\\hat{B}$ betweenness ternormalisasi  
- $\\hat{S}$ span = Polres + 0.5×estate (multi-lokasi)  
- $PAM$ = 1 jika seed PAM non-BUJP / bentrok  
- $\\hat{R}$ tetangga estate merah  
- $U$ = 1 jika di koridor utara (Rohul/Rohil/Bengkalis/Dumai)  

Ditambah **boost dampak** (poin) untuk 4 pilot R4: Majuma +35, PAB +28, UTS +25, Riden +18 — agar skor tidak semata degree jaringan.

**Band:** KRITIS ≥70 · TINGGI ≥50 · SEDANG ≥30 · RENDAH <30 · Agrinas = REFERENSI  

> **Prioritas pantau** = band ≥ SEDANG, atau PAM flag, atau multi-Polres/multi-estate."""))

cells.append(code("""print("Bobot:", {
    "degree": W_DEGREE,
    "betweenness": W_BETWEEN,
    "multi_polres": W_MULTI_POLRES,
    "pam_flag": W_PAM_FLAG,
    "estate_merah": W_ESTATE_MERAH,
})
print("Seed PAM/bentrok:", sorted(PAM_SEED))

hubs = score_hubs(metrics)
hubs"""))

cells.append(code("""# Watchlist operasional (prioritas pantau = YA)
watch = hubs[hubs["prioritas_pantau"] == "YA"].copy()
print(f"Jumlah watchlist: {len(watch)}")
watch[["aktor_id", "skor_hub", "skor_struktural", "boost_dampak", "band_hub",
       "degree", "betweenness", "jml_polres_terkait", "pam_non_bujp_flag", "polres_list"]]"""))

cells.append(md("""## 5. Matriks aktor × Polres

Menunjukkan aktor yang menjangkau lebih dari satu satker — indikasi **multi-lokasi** / potensi pola penunjukan berulang."""))

cells.append(code("""mat = actor_polres_matrix(G)
mat"""))

cells.append(code("""multi = mat[mat["total_polres"] >= 2].sort_values("total_polres", ascending=False)
print("Aktor multi-Polres:")
multi"""))

cells.append(md("""## 6. Ego-summary per Polres

Ringkasan beban aktor/PAM di tiap satker — untuk briefing Kasubdit / Kanit."""))

cells.append(code("""ego = ego_polres_summary(G, metrics)
ego"""))

cells.append(md("""## 7. Komunitas (modularity)

Klaster struktural kasar — membantu melihat “pulau” jaringan yang saling lebih rapat (bukan klaster Intelkam merah/kuning/hijau)."""))

cells.append(code("""comm = communities(G)
if comm.empty:
    print("Komunitas tidak terhitung.")
else:
    metrics_c = metrics.merge(comm, on="aktor_id", how="left")
    # ringkas komunitas berisi hub
    hub_ids = set(hubs["aktor_id"])
    ringkas = (
        metrics_c[metrics_c["aktor_id"].isin(hub_ids)]
        .groupby(["komunitas", "ukuran_komunitas"])
        .agg(aktor=("aktor_id", lambda s: ", ".join(sorted(s)[:6])), n=("aktor_id", "count"))
        .reset_index()
        .sort_values("ukuran_komunitas", ascending=False)
    )
    metrics = metrics_c  # simpan komunitas ke metrics aktif
    ringkas"""))

cells.append(md("""## 8. Visualisasi"""))

cells.append(code("""fig_degree_hub(hubs, OUT / "fig_degree_hub.png")
img = plt.imread(OUT / "fig_degree_hub.png")
plt.figure(figsize=(11, 6))
plt.imshow(img)
plt.axis("off")
plt.title("Ranking skor hub risiko")
plt.show()"""))

cells.append(code("""fig_betweenness(metrics, OUT / "fig_betweenness.png")
img = plt.imread(OUT / "fig_betweenness.png")
plt.figure(figsize=(11, 6))
plt.imshow(img)
plt.axis("off")
plt.title("Betweenness — aktor penjembatan")
plt.show()"""))

cells.append(code("""fig_network_overview(G, hubs, OUT / "fig_network_overview.png")
img = plt.imread(OUT / "fig_network_overview.png")
plt.figure(figsize=(12, 8))
plt.imshow(img)
plt.axis("off")
plt.title("Overview jaringan hub risiko")
plt.show()"""))

cells.append(md("""## 9. Drill-down aktor prioritas

Ubah `FOKUS` untuk inspeksi tetangga satu aktor (ego network tekstual)."""))

cells.append(code("""FOKUS = "PT Nusantara Sawit Majuma"  # ganti sesuai kebutuhan briefing

assert FOKUS in G, f"{FOKUS} tidak ada di graf"
row = hubs[hubs["aktor_id"] == FOKUS]
print("=== PROFIL ===")
profil = row if len(row) else metrics[metrics["aktor_id"] == FOKUS]
print(profil.to_string(index=False))

print("\\n=== TETANGGA ===")
rows = []
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

cells.append(md("""### 9.1 Batch profil 4 hub bentrok koridor utara"""))

cells.append(code("""PILOT = [
    "PT Nusantara Sawit Majuma",      # Rohul — MD
    "PT Palma Agung Betuah (PAB)",    # Bengkalis — SIS
    "PT Ujung Tanjung Sejahtera",     # Rohil — UTS
    "PT Riden Jaya Konstruksi",       # Dumai — DMMP
]

pilot_rows = []
for a in PILOT:
    if a not in G:
        pilot_rows.append({"aktor": a, "status": "TIDAK ADA DI GRAF"})
        continue
    estates = [nb for nb in G.neighbors(a) if str(G.nodes[nb].get("ntype","")).startswith("estate")]
    polres = sorted({
        *(nb for nb in G.neighbors(a) if str(G.nodes[nb].get("ntype","")).startswith("polres")),
        *(G.nodes[e].get("polres") for e in estates if G.nodes[e].get("polres")),
    })
    h = hubs[hubs["aktor_id"] == a]
    pilot_rows.append({
        "aktor": a,
        "skor_hub": float(h["skor_hub"].iloc[0]) if len(h) else None,
        "band_hub": h["band_hub"].iloc[0] if len(h) else None,
        "estate": ", ".join(estates) if estates else "—",
        "polres": ", ".join([p for p in polres if p]),
        "degree": G.degree(a),
        "status": "OK",
    })
pd.DataFrame(pilot_rows)"""))

cells.append(md("""## 10. Ekspor produk Harda

Menyimpan seluruh artefak ke `output/aktor_kso/` dan menulis ringkasan temuan markdown."""))

cells.append(code("""# Simpan tabel
metrics.to_csv(OUT / "tabel_aktor_metrics.csv", index=False, encoding="utf-8-sig")
hubs.to_csv(OUT / "tabel_hub_risiko.csv", index=False, encoding="utf-8-sig")
ego.to_csv(OUT / "tabel_ego_polres.csv", index=False, encoding="utf-8-sig")
mat.to_csv(OUT / "matriks_aktor_polres.csv", encoding="utf-8-sig")
write_findings(hubs, metrics, ego, OUT / "ringkasan_temuan.md")

print("Tersimpan di:", OUT)
print("\\n".join(f" - {p.name}" for p in sorted(OUT.glob("*"))))

# tampilkan ringkasan temuan
print("\\n" + (OUT / "ringkasan_temuan.md").read_text(encoding="utf-8")[:1800])"""))

cells.append(md("""## 11. Checklist interpretasi untuk analis Harda

| Pertanyaan | Cara baca di notebook |
|---|---|
| Siapa yang harus dipantau bulan ini? | Tabel `watch` / `tabel_hub_risiko.csv` (`prioritas_pantau = YA`) |
| Apakah risiko menular antar kebun? | Betweenness tinggi + multi-Polres |
| Polres mana yang padat PAM? | `ego` → kolom `aktor_pam_flag` |
| Apa yang masih gelap di data? | Node generik `KSO` / `Belum ada KSO` → tindak gap-fill satker |
| Apakah sama dengan klaster Intelkam? | Tidak otomatis — cocokkan dengan R3 validasi klaster |

### Batasan
- Graf dibangun dari cuplikan dokumen yang tersedia; coverage ≠ 130 kebun penuh.
- Skor hub adalah **alat prioritisasi analitik**, bukan bukti pidana.
- TNTN Pelalawan harus dibaca terpisah dari tipologi KSO murni.

---
*Unit II Harda · Ditreskrimum Polda Riau · Notebook analisis aktor KSO Agrinas*"""))

cells.append(md("""## Appendix — jalankan pipeline penuh satu panggilan

Cell ini opsional; berguna untuk refresh cepat tanpa menjalankan cell di atas satu per satu."""))

cells.append(code("""# Uncomment untuk refresh penuh:
# result = run_analysis(verbose=True)
# result["hubs"].head(10)"""))

# strip trailing newlines artifacts: nbformat prefers sources as arrays; last line newline ok
# Fix display() for environments without IPython
# Replace display with print-friendly in community cell - actually jupyter has display

nb = {
    "nbformat": 4,
    "nbformat_minor": 5,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "name": "python",
            "version": "3.11",
        },
        "authors": [{"name": "Unit II Harda Ditreskrimum Polda Riau"}],
        "title": "Analisis Aktor KSO Agrinas — Unit II Harda",
    },
    "cells": cells,
}

# Clean source arrays: remove final extra newline-only duplication by ensuring each source line ends with \n
# but last cell lines from split already have \n on each including empty trailing from split("\n") on ending newline
for c in nb["cells"]:
    src = c["source"]
    if src and src[-1] == "\n":
        c["source"] = src[:-1]
    # ensure at least one line
    if not c["source"]:
        c["source"] = [""]

NB.write_text(json.dumps(nb, ensure_ascii=False, indent=1), encoding="utf-8")
print("OK:", NB)
