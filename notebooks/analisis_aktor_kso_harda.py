# -*- coding: utf-8 -*-
"""
Analisis Aktor Jaringan KSO Agrinas
Unit II Harda — Ditreskrimum Polda Riau

Skrip inti yang dipakai notebook. Dapat dijalankan mandiri:
  python notebooks/analisis_aktor_kso_harda.py

Output (folder output/aktor_kso/):
  - tabel_aktor_metrics.csv
  - tabel_hub_risiko.csv
  - tabel_ego_polres.csv
  - matriks_aktor_polres.csv
  - fig_degree_hub.png
  - fig_network_overview.png
  - fig_betweenness.png
  - ringkasan_temuan.md
"""

from __future__ import annotations

import warnings
from pathlib import Path

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════
# 0. KONFIGURASI UNIT
# ═══════════════════════════════════════════════════════════════

BASE = Path(r"C:\Users\Patron\Downloads\sawit lagi")
GEXF = BASE / "jaringan_kso_agrinas.gexf"
XLSX_R = BASE / "analisis_lanjutan_5_prioritas.xlsx"
XLSX_M = BASE / "matriks_agrinas_kso_12_polres.xlsx"
OUT = BASE / "output" / "aktor_kso"
OUT.mkdir(parents=True, exist_ok=True)

UNIT = "Unit II Harda — Ditreskrimum Polda Riau"
JUDUL = "Analisis Aktor Jaringan KSO Agrinas"
TANGGAL = "4 Agustus 2026"

# Bobot skor hub risiko Harda (0–100)
W_DEGREE = 0.25
W_BETWEEN = 0.25
W_MULTI_POLRES = 0.20
W_PAM_FLAG = 0.20
W_ESTATE_MERAH = 0.10

PAM_SEED = {
    "PT Nusantara Sawit Majuma",
    "PT Palma Agung Betuah (PAB)",
    "PT Ujung Tanjung Sejahtera",
    "PT Riden Jaya Konstruksi",
    "Makmur Jaya Sentosa",
    "Poktan Riau Jaya Makmur",
    "Torus",
    "Togos",
    "Togas",
}

# Boost dampak aktual (R4 pilot) — melengkapi skor struktural jaringan
# nilai dalam poin skor (0–100)
DAMPAK_BOOST = {
    "PT Nusantara Sawit Majuma": 35,      # 1 MD + bentrok PAM
    "PT Palma Agung Betuah (PAB)": 28,    # eskalasi berulang SIS
    "PT Ujung Tanjung Sejahtera": 25,     # 7 luka
    "PT Riden Jaya Konstruksi": 18,       # take-over / bentrok
}

GENERIC = {"KSO", "Belum ada KSO"}

COLOR = {
    "agrinas": "#0F2A44",
    "aktor_pam": "#B32D2D",
    "aktor_multi": "#C45C26",
    "aktor": "#2C5F7C",
    "aktor_generic": "#8A95A1",
    "estate_merah": "#B32D2D",
    "estate_kuning": "#C48A14",
    "estate_hijau": "#2E7D4F",
    "estate": "#5A6672",
    "polres_kritis": "#8B1E1E",
    "polres_tinggi": "#C45C26",
    "polres_sedang": "#B88A3D",
    "polres_rendah": "#3D6B5A",
}


def load_sheet(path: Path, name: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def load_graph() -> nx.Graph:
    G = nx.read_gexf(GEXF)
    # normalisasi atribut
    for n, d in G.nodes(data=True):
        d["ntype"] = d.get("ntype") or "estate"
        d["hub_risiko"] = str(d.get("hub_risiko", "")).lower() in ("true", "1", "ya")
        d["label"] = (d.get("label") or n).replace("\n", " ").strip()
        d["polres"] = d.get("polres") or ""
        d["klaster"] = d.get("klaster") or ""
        d["band"] = d.get("band") or ""
        d["peran"] = d.get("peran") or ""
        try:
            d["jml_polres"] = int(float(d.get("jml_polres") or 0))
        except Exception:
            d["jml_polres"] = 0
        try:
            d["jml_estate"] = int(float(d.get("jml_estate") or 0))
        except Exception:
            d["jml_estate"] = 0
    return G


def is_aktor(ntype: str) -> bool:
    return ntype == "agrinas" or ntype.startswith("aktor")


def is_estate(ntype: str) -> bool:
    return ntype.startswith("estate")


def is_polres(ntype: str) -> bool:
    return ntype.startswith("polres")


# ═══════════════════════════════════════════════════════════════
# 1. METRIK JARINGAN
# ═══════════════════════════════════════════════════════════════

def compute_metrics(G: nx.Graph) -> pd.DataFrame:
    deg = dict(G.degree())
    # betweenness pada graf penuh
    btw = nx.betweenness_centrality(G, normalized=True)
    # closeness bisa gagal jika disconnected — gunakan per komponen
    clo = {}
    for comp in nx.connected_components(G):
        sub = G.subgraph(comp)
        clo.update(nx.closeness_centrality(sub))
    try:
        eig = nx.eigenvector_centrality_numpy(G)
    except Exception:
        eig = {n: 0.0 for n in G.nodes()}

    rows = []
    for n, d in G.nodes(data=True):
        # estate tetangga merah
        merah_n = 0
        polres_set = set()
        estate_set = set()
        for nb in G.neighbors(n):
            nd = G.nodes[nb]
            if is_estate(nd.get("ntype", "")):
                estate_set.add(nb)
                if "merah" in (nd.get("ntype") or "") or "merah" in (nd.get("klaster") or "").lower():
                    merah_n += 1
                if nd.get("polres"):
                    polres_set.add(nd["polres"])
            if is_polres(nd.get("ntype", "")):
                polres_set.add(nb)

        ntype = d.get("ntype", "")
        pam_flag = 1 if (n in PAM_SEED or ntype == "aktor_pam") else 0
        multi = max(d.get("jml_polres", 0), len(polres_set))

        rows.append(
            {
                "aktor_id": n,
                "label": d.get("label", n),
                "ntype": ntype,
                "bucket": (
                    "agrinas"
                    if ntype == "agrinas"
                    else "aktor"
                    if is_aktor(ntype)
                    else "estate"
                    if is_estate(ntype)
                    else "polres"
                    if is_polres(ntype)
                    else "lain"
                ),
                "hub_risiko_flag": "YA" if d.get("hub_risiko") else "TIDAK",
                "pam_non_bujp_flag": "YA" if pam_flag else "TIDAK",
                "generic_flag": "YA" if n in GENERIC else "TIDAK",
                "degree": deg.get(n, 0),
                "betweenness": round(btw.get(n, 0.0), 4),
                "closeness": round(clo.get(n, 0.0), 4),
                "eigenvector": round(float(eig.get(n, 0.0)), 4),
                "jml_estate_tetangga": len(estate_set),
                "jml_polres_terkait": multi,
                "jml_estate_merah_tetangga": merah_n,
                "polres_list": ", ".join(sorted(polres_set)),
                "peran": d.get("peran", ""),
                "klaster": d.get("klaster", ""),
                "band_polres": d.get("band", ""),
            }
        )
    df = pd.DataFrame(rows)
    return df.sort_values(["bucket", "degree"], ascending=[True, False]).reset_index(drop=True)


def score_hubs(df: pd.DataFrame) -> pd.DataFrame:
    """Skor komposit untuk penerima KSO/PAM.

    Agrinas dihitung terpisah sebagai referensi struktural (tidak menormalisasi
    ulang skala degree/betweenness penerima KSO).
    """
    all_aktor = df[df["bucket"].isin(["aktor", "agrinas"])].copy()
    all_aktor = all_aktor[all_aktor["generic_flag"] != "YA"].copy()

    # Pool penskalaan: hanya penerima KSO/PAM (bukan Agrinas pusat)
    pool = all_aktor[all_aktor["aktor_id"] != "Agrinas Palma Nusantara"].copy()

    def norm(s):
        s = s.astype(float)
        if len(s) == 0 or s.max() == s.min():
            return pd.Series(np.zeros(len(s)), index=s.index)
        return (s - s.min()) / (s.max() - s.min())

    # Multi-lokasi: gabungan polres + estate (penting untuk Bernas/Berlian/dll.)
    pool["span"] = pool["jml_polres_terkait"].astype(float) + 0.5 * pool["jml_estate_tetangga"].astype(float)

    pool["n_degree"] = norm(pool["degree"])
    pool["n_between"] = norm(pool["betweenness"])
    pool["n_multi"] = norm(pool["span"])
    pool["n_pam"] = (pool["pam_non_bujp_flag"] == "YA").astype(float)
    pool["n_merah"] = norm(pool["jml_estate_merah_tetangga"])

    # Bonus koridor utara (lokasi bentrok historis) — ringan, tidak mendominasi
    utara = {"Rohul", "Rohil", "Bengkalis", "Dumai"}
    pool["n_utara"] = pool["polres_list"].fillna("").apply(
        lambda s: 1.0 if any(p in s.split(", ") for p in utara) else 0.0
    )

    # Redistribusi bobot: sisipkan 5% koridor utara dari PAM
    w_pam = W_PAM_FLAG - 0.05
    w_utara = 0.05

    pool["skor_struktural"] = (
        W_DEGREE * pool["n_degree"]
        + W_BETWEEN * pool["n_between"]
        + W_MULTI_POLRES * pool["n_multi"]
        + w_pam * pool["n_pam"]
        + W_ESTATE_MERAH * pool["n_merah"]
        + w_utara * pool["n_utara"]
    ) * 100

    pool["boost_dampak"] = pool["aktor_id"].map(DAMPAK_BOOST).fillna(0.0)
    pool["skor_hub"] = (pool["skor_struktural"] + pool["boost_dampak"]).clip(upper=100)

    def band(x):
        if x >= 70:
            return "KRITIS"
        if x >= 50:
            return "TINGGI"
        if x >= 30:
            return "SEDANG"
        return "RENDAH"

    pool["band_hub"] = pool["skor_hub"].apply(band)
    pool["skor_hub"] = pool["skor_hub"].round(1)
    pool["skor_struktural"] = pool["skor_struktural"].round(1)

    # Agrinas sebagai baris referensi (bukan target pantau oknum)
    agr = all_aktor[all_aktor["aktor_id"] == "Agrinas Palma Nusantara"].copy()
    if len(agr):
        agr["skor_hub"] = 100.0
        agr["skor_struktural"] = 100.0
        agr["boost_dampak"] = 0.0
        agr["band_hub"] = "REFERENSI"
        agr["span"] = agr["jml_polres_terkait"]
        aktor = pd.concat([agr, pool], ignore_index=True)
    else:
        aktor = pool

    aktor["prioritas_pantau"] = np.where(
        (aktor["aktor_id"] != "Agrinas Palma Nusantara")
        & (
            aktor["band_hub"].isin(["KRITIS", "TINGGI", "SEDANG"])
            | (aktor["pam_non_bujp_flag"] == "YA")
            | (aktor["jml_polres_terkait"] >= 2)
            | (aktor["jml_estate_tetangga"] >= 2)
        ),
        "YA",
        "TIDAK",
    )

    if "skor_struktural" not in aktor.columns:
        aktor["skor_struktural"] = aktor["skor_hub"]
        aktor["boost_dampak"] = 0.0
    else:
        aktor["boost_dampak"] = aktor.get("boost_dampak", 0.0)

    cols = [
        "aktor_id",
        "label",
        "ntype",
        "skor_hub",
        "skor_struktural",
        "boost_dampak",
        "band_hub",
        "prioritas_pantau",
        "degree",
        "betweenness",
        "jml_polres_terkait",
        "jml_estate_tetangga",
        "jml_estate_merah_tetangga",
        "pam_non_bujp_flag",
        "hub_risiko_flag",
        "polres_list",
        "peran",
    ]
    # pastikan kolom ada di agrinas rows
    for c in cols:
        if c not in aktor.columns:
            aktor[c] = 0 if c in ("skor_struktural", "boost_dampak") else ""
    return aktor[cols].sort_values(["skor_hub", "degree"], ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════
# 2. EGO / MATRIKS
# ═══════════════════════════════════════════════════════════════

def actor_polres_matrix(G: nx.Graph) -> pd.DataFrame:
    records = []
    for n, d in G.nodes(data=True):
        if not is_aktor(d.get("ntype", "")) or n in GENERIC:
            continue
        polres = set()
        for nb in G.neighbors(n):
            nd = G.nodes[nb]
            if is_polres(nd.get("ntype", "")):
                polres.add(nb)
            if is_estate(nd.get("ntype", "")) and nd.get("polres"):
                polres.add(nd["polres"])
        for p in polres:
            records.append({"aktor": n, "polres": p, "link": 1})
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    mat = df.pivot_table(index="aktor", columns="polres", values="link", aggfunc="max", fill_value=0)
    mat["total_polres"] = mat.sum(axis=1)
    return mat.sort_values("total_polres", ascending=False)


def ego_polres_summary(G: nx.Graph, metrics: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for n, d in G.nodes(data=True):
        if not is_polres(d.get("ntype", "")):
            continue
        actors, estates, merah, pam = [], [], 0, 0
        for nb in G.neighbors(n):
            nd = G.nodes[nb]
            if is_aktor(nd.get("ntype", "")) and nb not in GENERIC:
                actors.append(nb)
                if nb in PAM_SEED or nd.get("ntype") == "aktor_pam":
                    pam += 1
            if is_estate(nd.get("ntype", "")):
                estates.append(nb)
                if "merah" in (nd.get("ntype") or ""):
                    merah += 1
            # estates linked indirectly via actors already counted as neighbors of polres
        # also pull estates connected to this polres
        for nb in list(G.neighbors(n)):
            pass
        rows.append(
            {
                "polres": n,
                "band_risiko": d.get("band", ""),
                "jml_aktor_langsung": len(set(actors)),
                "jml_estate_langsung": len(set(estates)),
                "estate_merah": merah,
                "aktor_pam_flag": pam,
                "aktor_contoh": ", ".join(sorted(set(actors))[:5]),
            }
        )
    return pd.DataFrame(rows).sort_values(
        ["band_risiko", "aktor_pam_flag", "estate_merah"],
        ascending=[True, False, False],
        key=lambda s: s if s.name != "band_risiko" else s.map({"KRITIS": 0, "TINGGI": 1, "SEDANG": 2, "RENDAH": 3}),
    ).reset_index(drop=True)


def communities(G: nx.Graph) -> pd.DataFrame:
    try:
        from networkx.algorithms import community as nx_comm

        comms = list(nx_comm.greedy_modularity_communities(G))
    except Exception:
        return pd.DataFrame(columns=["aktor_id", "komunitas", "ukuran_komunitas"])
    rows = []
    for i, c in enumerate(sorted(comms, key=len, reverse=True), 1):
        for n in c:
            rows.append({"aktor_id": n, "komunitas": f"C{i:02d}", "ukuran_komunitas": len(c)})
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════
# 3. VISUALISASI
# ═══════════════════════════════════════════════════════════════

def fig_degree_hub(hubs: pd.DataFrame, path: Path):
    top = hubs[~hubs["band_hub"].isin(["REFERENSI"])].head(12).iloc[::-1]
    fig, ax = plt.subplots(figsize=(10, 6), facecolor="#F4F6F8")
    colors = [
        "#B32D2D" if b == "KRITIS" else "#C45C26" if b == "TINGGI" else "#B88A3D" if b == "SEDANG" else "#2C5F7C"
        for b in top["band_hub"]
    ]
    ax.barh(top["label"], top["skor_hub"], color=colors)
    ax.set_xlabel("Skor hub risiko (0–100)")
    ax.set_title("Ranking Hub Risiko Aktor KSO (ex-Agrinas pusat)\nUnit II Harda · Ditreskrimum Polda Riau")
    ax.set_xlim(0, 100)
    for i, (skor, band) in enumerate(zip(top["skor_hub"], top["band_hub"])):
        ax.text(skor + 1, i, f"{skor:.0f} · {band}", va="center", fontsize=8, color="#1E242A")
    ax.set_facecolor("#F4F6F8")
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)


def fig_betweenness(metrics: pd.DataFrame, path: Path):
    aktor = metrics[(metrics["bucket"] == "aktor") & (metrics["generic_flag"] != "YA")].nlargest(12, "betweenness").iloc[::-1]
    fig, ax = plt.subplots(figsize=(10, 6), facecolor="#F4F6F8")
    ax.barh(aktor["label"], aktor["betweenness"], color="#2C5F7C")
    ax.set_xlabel("Betweenness centrality (normalized)")
    ax.set_title("Aktor Penjembatan (Betweenness) — potensi penularan risiko antar kebun/Polres")
    ax.set_facecolor("#F4F6F8")
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)


def fig_network_overview(G: nx.Graph, hubs: pd.DataFrame, path: Path):
    # fokus: agrinas + top hubs + tetangga estate/polres
    top_ids = set(hubs.head(15)["aktor_id"]) | {"Agrinas Palma Nusantara"}
    keep = set(top_ids)
    for n in list(top_ids):
        if n in G:
            keep.update(G.neighbors(n))
    H = G.subgraph([n for n in keep if n in G]).copy()

    pos = nx.spring_layout(H, seed=42, k=0.7)
    fig, ax = plt.subplots(figsize=(14, 10), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")

    # edges
    nx.draw_networkx_edges(H, pos, ax=ax, alpha=0.25, width=1.0, edge_color="#5A6672")

    for ntype, nodes in _group_nodes(H).items():
        sizes = []
        for n in nodes:
            base = 800 if ntype == "agrinas" else 500 if ntype.startswith("polres") else 350 if ntype.startswith("aktor") else 180
            sizes.append(base + H.degree(n) * 40)
        nx.draw_networkx_nodes(
            H, pos, nodelist=nodes, ax=ax,
            node_color=COLOR.get(ntype, "#5A6672"),
            node_size=sizes, alpha=0.9, linewidths=0.8, edgecolors="white",
        )

    labels = {}
    for n, d in H.nodes(data=True):
        if d.get("hub_risiko") or d.get("ntype") == "agrinas" or is_polres(d.get("ntype", "")):
            labels[n] = (d.get("label") or n)[:26]
    nx.draw_networkx_labels(H, pos, labels=labels, ax=ax, font_size=7, font_color="#1E242A")
    ax.set_title("Overview Jaringan Hub Risiko KSO Agrinas\n(Agrinas · aktor prioritas · estate · Polres)")
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)


def _group_nodes(G):
    g = {}
    for n, d in G.nodes(data=True):
        g.setdefault(d.get("ntype", "estate"), []).append(n)
    return g


# ═══════════════════════════════════════════════════════════════
# 4. TEMUAN
# ═══════════════════════════════════════════════════════════════

def write_findings(hubs: pd.DataFrame, metrics: pd.DataFrame, ego: pd.DataFrame, path: Path):
    top = hubs[hubs["prioritas_pantau"] == "YA"].head(8)
    lines = [
        f"# {JUDUL}",
        f"**{UNIT}** · {TANGGAL}",
        "",
        "## Ringkasan metodologi",
        "- Graf sumber: `jaringan_kso_agrinas.gexf` (aktor–estate–Polres).",
        "- Metrik: degree, betweenness, closeness, eigenvector.",
        "- Skor hub Harda = 25% degree + 25% betweenness + 20% span(multi) + 15% PAM + 10% estate merah + 5% koridor utara.",
        "- Normalisasi tanpa Agrinas pusat; Agrinas berlabel REFERENSI.",
        "- Band: KRITIS ≥70 · TINGGI ≥50 · SEDANG ≥30 · RENDAH <30.",
        "",
        "## Temuan utama",
        "1. **Agrinas** adalah pusat struktural skema (degree tertinggi) — dipakai sebagai referensi, bukan target 'oknum'.",
        "2. Risiko analitik berada pada **aktor penerima KSO/PAM** yang menjembatani banyak kebun atau terkait bentrok.",
        "3. Koridor utara (Rohul–Rohil–Bengkalis–Dumai) memusatkan hub PAM berbendera bentrok (Majuma, PAB, UTS, Riden Jaya).",
        "4. Hub multi-lokasi (Bernas Mulya Mandiri, Berlian NP, Maju Serempak, PT Runggu) memperbesar peluang penolakan berulang lintas estate.",
        "5. Label generik `KSO` / `Belum ada KSO` di graf menandai **gap data satker**, bukan entitas hukum — jangan dihitung sebagai hub operasional.",
        "",
        "## Prioritas pantau aktor",
        "| Rank | Aktor | Skor | Struktural | Boost dampak | Band | Polres | PAM |",
        "|---:|---|---:|---:|---:|---|---|---|",
    ]
    for i, r in enumerate(top.itertuples(index=False), 1):
        lines.append(
            f"| {i} | {r.aktor_id} | {r.skor_hub} | {getattr(r, 'skor_struktural', r.skor_hub)} | "
            f"{getattr(r, 'boost_dampak', 0)} | {r.band_hub} | {r.polres_list or '—'} | {r.pam_non_bujp_flag} |"
        )
    lines += [
        "",
        "## Implikasi kerja Unit II Harda",
        "1. Perbarui watchlist aktor bulanan dari `tabel_hub_risiko.csv`.",
        "2. Cocokkan hub PAM dengan LP/bentrok (R3/R4) — terutama Rohul, Rohil, Bengkalis, Dumai, Kampar.",
        "3. Minta satker melengkapi nama di balik label generik KSO (gap-fill R5).",
        "4. Analisis ego-Polres: satker dengan `aktor_pam_flag` > 0 wajib masuk briefing mingguan.",
        "5. Jangan agregasi TNTN Pelalawan ke skor KSO murni.",
        "",
        "## File keluaran",
        "- `tabel_aktor_metrics.csv`",
        "- `tabel_hub_risiko.csv`",
        "- `tabel_ego_polres.csv`",
        "- `matriks_aktor_polres.csv`",
        "- `fig_degree_hub.png`, `fig_betweenness.png`, `fig_network_overview.png`",
        "",
        "*Dokumen analitik — bukan perintah operasi.*",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


# ═══════════════════════════════════════════════════════════════
# 5. MAIN
# ═══════════════════════════════════════════════════════════════

def run_analysis(verbose: bool = True):
    if verbose:
        print(f"=== {JUDUL} ===")
        print(UNIT)
        print(f"Sumber: {GEXF.name}")

    G = load_graph()
    if verbose:
        print(f"Graph: {G.number_of_nodes()} nodes · {G.number_of_edges()} edges")

    metrics = compute_metrics(G)
    hubs = score_hubs(metrics)
    ego = ego_polres_summary(G, metrics)
    mat = actor_polres_matrix(G)
    comm = communities(G)
    if not comm.empty:
        metrics = metrics.merge(comm, on="aktor_id", how="left")

    # enrich hubs dengan komunitas
    if "komunitas" in metrics.columns:
        hubs = hubs.merge(metrics[["aktor_id", "komunitas", "ukuran_komunitas"]], on="aktor_id", how="left")

    # optional: join Excel hub sheet
    try:
        hub_xls = load_sheet(XLSX_R, "R2_hub_aktor")
        hub_xls = hub_xls.rename(columns={"aktor": "aktor_id"})
        hubs = hubs.merge(
            hub_xls[["aktor_id", "catatan", "flag_hub_risiko"]],
            on="aktor_id",
            how="left",
        )
    except Exception:
        pass

    metrics.to_csv(OUT / "tabel_aktor_metrics.csv", index=False, encoding="utf-8-sig")
    hubs.to_csv(OUT / "tabel_hub_risiko.csv", index=False, encoding="utf-8-sig")
    ego.to_csv(OUT / "tabel_ego_polres.csv", index=False, encoding="utf-8-sig")
    mat.to_csv(OUT / "matriks_aktor_polres.csv", encoding="utf-8-sig")

    fig_degree_hub(hubs, OUT / "fig_degree_hub.png")
    fig_betweenness(metrics, OUT / "fig_betweenness.png")
    fig_network_overview(G, hubs, OUT / "fig_network_overview.png")
    write_findings(hubs, metrics, ego, OUT / "ringkasan_temuan.md")

    if verbose:
        print("\n=== TOP 10 HUB (ex-Agrinas generik) ===")
        show = hubs[hubs["aktor_id"] != "Agrinas Palma Nusantara"].head(10)
        print(show[["aktor_id", "skor_hub", "band_hub", "pam_non_bujp_flag", "polres_list"]].to_string(index=False))
        print(f"\nOutput -> {OUT}")

    return {
        "G": G,
        "metrics": metrics,
        "hubs": hubs,
        "ego": ego,
        "matrix": mat,
        "out": OUT,
    }


if __name__ == "__main__":
    run_analysis(verbose=True)
