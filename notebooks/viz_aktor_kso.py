# -*- coding: utf-8 -*-
"""Visualisasi notebook analisis aktor KSO — Unit II Harda."""

from __future__ import annotations

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd
from matplotlib.patches import Patch

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

BAND_COLOR = {
    "KRITIS": "#B32D2D",
    "TINGGI": "#C45C26",
    "SEDANG": "#B88A3D",
    "RENDAH": "#2C5F7C",
    "REFERENSI": "#0F2A44",
}


def _style(ax, title: str):
    ax.set_facecolor("#F4F6F8")
    ax.set_title(title, fontsize=12, fontweight="bold", color="#0F2A44", pad=10)
    ax.grid(axis="x", alpha=0.25)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


def plot_skor_hub(hubs: pd.DataFrame, top_n: int = 12, ax=None):
    data = hubs[hubs["band_hub"] != "REFERENSI"].head(top_n).iloc[::-1]
    if ax is None:
        fig, ax = plt.subplots(figsize=(10, 5.8), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    colors = [BAND_COLOR.get(b, "#2C5F7C") for b in data["band_hub"]]
    ax.barh(data["label"].str.slice(0, 34), data["skor_hub"], color=colors, height=0.7)
    for i, (skor, band, boost) in enumerate(zip(data["skor_hub"], data["band_hub"], data.get("boost_dampak", [0]*len(data)))):
        note = f"{skor:.0f}  {band}"
        if float(boost or 0) > 0:
            note += f"  (+{float(boost):.0f} dampak)"
        ax.text(min(skor + 1.2, 96), i, note, va="center", fontsize=8, color="#1E242A")
    ax.set_xlim(0, 100)
    ax.set_xlabel("Skor hub risiko (0–100)")
    _style(ax, "Ranking skor hub risiko aktor KSO (ex-Agrinas)")
    legend = [Patch(color=c, label=b) for b, c in BAND_COLOR.items() if b != "REFERENSI"]
    ax.legend(handles=legend, loc="lower right", fontsize=8, framealpha=0.95)
    fig.tight_layout()
    return fig


def plot_betweenness(metrics: pd.DataFrame, top_n: int = 12, ax=None):
    data = metrics[(metrics["bucket"] == "aktor") & (metrics["generic_flag"] != "YA")].nlargest(top_n, "betweenness").iloc[::-1]
    if ax is None:
        fig, ax = plt.subplots(figsize=(10, 5.5), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    ax.barh(data["label"].str.slice(0, 34), data["betweenness"], color="#2C5F7C", height=0.7)
    ax.set_xlabel("Betweenness centrality")
    _style(ax, "Aktor penjembatan — potensi penularan risiko lintas kebun/Polres")
    fig.tight_layout()
    return fig


def plot_degree_vs_betweenness(hubs: pd.DataFrame, ax=None):
    data = hubs[hubs["band_hub"] != "REFERENSI"].copy()
    if ax is None:
        fig, ax = plt.subplots(figsize=(9, 6), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    for band, g in data.groupby("band_hub"):
        ax.scatter(
            g["degree"], g["betweenness"],
            s=80 + g["skor_hub"] * 2,
            c=BAND_COLOR.get(band, "#2C5F7C"),
            label=band, alpha=0.85, edgecolors="white", linewidths=0.8,
        )
        for _, r in g.iterrows():
            if r["skor_hub"] >= 40 or r["pam_non_bujp_flag"] == "YA":
                ax.annotate(str(r["label"])[:22], (r["degree"], r["betweenness"]),
                            textcoords="offset points", xytext=(5, 4), fontsize=7, color="#1E242A")
    ax.set_xlabel("Degree")
    ax.set_ylabel("Betweenness")
    ax.legend(fontsize=8)
    _style(ax, "Pemetaan posisi aktor: konektivitas vs peran jembatan")
    ax.grid(True, alpha=0.25)
    fig.tight_layout()
    return fig


def plot_heatmap_aktor_polres(mat: pd.DataFrame, ax=None):
    if mat is None or mat.empty:
        fig, ax = plt.subplots(figsize=(8, 3), facecolor="#F4F6F8")
        ax.text(0.5, 0.5, "Matriks kosong", ha="center")
        ax.axis("off")
        return fig
    m = mat.drop(columns=["total_polres"], errors="ignore")
    if ax is None:
        fig, ax = plt.subplots(figsize=(max(8, len(m.columns) * 0.9), max(5, len(m) * 0.38)), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    im = ax.imshow(m.values, cmap="YlOrRd", aspect="auto", vmin=0, vmax=1)
    ax.set_xticks(range(len(m.columns)))
    ax.set_yticks(range(len(m.index)))
    ax.set_xticklabels(m.columns, rotation=40, ha="right", fontsize=9)
    ax.set_yticklabels([str(i)[:32] for i in m.index], fontsize=8)
    ax.set_title("Matriks keterkaitan aktor × Polres", fontsize=12, fontweight="bold", color="#0F2A44")
    for i in range(m.shape[0]):
        for j in range(m.shape[1]):
            if m.values[i, j]:
                ax.text(j, i, "●", ha="center", va="center", color="#0F2A44", fontsize=10)
    fig.colorbar(im, ax=ax, fraction=0.03, pad=0.02, label="terhubung")
    fig.tight_layout()
    return fig


def plot_band_distribution(hubs: pd.DataFrame, ax=None):
    data = hubs[hubs["band_hub"] != "REFERENSI"]["band_hub"].value_counts().reindex(
        ["KRITIS", "TINGGI", "SEDANG", "RENDAH"]
    ).fillna(0)
    if ax is None:
        fig, ax = plt.subplots(figsize=(7, 4.5), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    colors = [BAND_COLOR[b] for b in data.index]
    bars = ax.bar(data.index, data.values, color=colors, width=0.65)
    for b, v in zip(bars, data.values):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.1, str(int(v)), ha="center", fontsize=10, fontweight="bold")
    ax.set_ylabel("Jumlah aktor")
    _style(ax, "Distribusi band hub risiko")
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    return fig


def plot_ego_polres(ego: pd.DataFrame, ax=None):
    data = ego.copy()
    order = {"KRITIS": 0, "TINGGI": 1, "SEDANG": 2, "RENDAH": 3}
    data["_o"] = data["band_risiko"].map(order).fillna(9)
    data = data.sort_values(["_o", "aktor_pam_flag", "estate_merah"], ascending=[True, False, False])
    if ax is None:
        fig, ax = plt.subplots(figsize=(10, 5), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    x = np.arange(len(data))
    w = 0.35
    ax.bar(x - w / 2, data["jml_aktor_langsung"], w, label="Aktor langsung", color="#2C5F7C")
    ax.bar(x + w / 2, data["aktor_pam_flag"], w, label="Aktor PAM/flag", color="#B32D2D")
    ax.set_xticks(x)
    ax.set_xticklabels(data["polres"], rotation=30, ha="right")
    ax.set_ylabel("Jumlah")
    ax.legend(fontsize=8)
    _style(ax, "Beban aktor per Polres (ego-summary)")
    fig.tight_layout()
    return fig


def plot_network(G: nx.Graph, hubs: pd.DataFrame, focus_hubs: bool = True, seed: int = 42, ax=None):
    if focus_hubs:
        top = set(hubs[hubs["band_hub"] != "REFERENSI"].head(12)["aktor_id"]) | {"Agrinas Palma Nusantara"}
        keep = set(top)
        for n in list(top):
            if n in G:
                keep.update(G.neighbors(n))
        H = G.subgraph([n for n in keep if n in G]).copy()
        title = "Peta jaringan fokus hub risiko (aktor · estate · Polres)"
    else:
        H = G
        title = "Peta jaringan penuh KSO Agrinas"

    pos = nx.spring_layout(H, seed=seed, k=0.65 if focus_hubs else 0.4)
    if ax is None:
        fig, ax = plt.subplots(figsize=(12, 8.5), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    ax.set_facecolor("#F4F6F8")

    nx.draw_networkx_edges(H, pos, ax=ax, alpha=0.28, width=1.0, edge_color="#6A7680")

    groups = {}
    for n, d in H.nodes(data=True):
        groups.setdefault(d.get("ntype", "estate"), []).append(n)

    for ntype, nodes in groups.items():
        sizes = []
        for n in nodes:
            if ntype == "agrinas":
                sizes.append(2200)
            elif str(ntype).startswith("polres"):
                sizes.append(900 + H.degree(n) * 50)
            elif str(ntype).startswith("aktor"):
                sizes.append(500 + H.degree(n) * 70)
            else:
                sizes.append(180 + H.degree(n) * 40)
        nx.draw_networkx_nodes(
            H, pos, nodelist=nodes, ax=ax,
            node_color=COLOR.get(ntype, "#5A6672"),
            node_size=sizes, alpha=0.92, linewidths=0.9, edgecolors="white",
        )

    labels = {}
    for n, d in H.nodes(data=True):
        nt = d.get("ntype", "")
        if nt == "agrinas" or d.get("hub_risiko") or str(nt).startswith("polres") or "merah" in str(nt):
            labels[n] = (d.get("label") or n)[:24]
    nx.draw_networkx_labels(H, pos, labels=labels, ax=ax, font_size=7.5, font_color="#1E242A")

    legend = [
        Patch(color=COLOR["agrinas"], label="Agrinas"),
        Patch(color=COLOR["aktor_pam"], label="Hub PAM/risiko"),
        Patch(color=COLOR["aktor_multi"], label="Hub multi-lokasi"),
        Patch(color=COLOR["estate_merah"], label="Estate merah"),
        Patch(color=COLOR["estate_kuning"], label="Estate kuning"),
        Patch(color=COLOR["estate_hijau"], label="Estate hijau"),
        Patch(color=COLOR["polres_kritis"], label="Polres KRITIS/TINGGI"),
    ]
    ax.legend(handles=legend, loc="lower left", fontsize=8, framealpha=0.95, ncol=2)
    ax.set_title(title, fontsize=13, fontweight="bold", color="#0F2A44")
    ax.axis("off")
    fig.tight_layout()
    return fig


def plot_ego_aktor(G: nx.Graph, fokus: str, ax=None):
    if fokus not in G:
        fig, ax = plt.subplots(figsize=(8, 3), facecolor="#F4F6F8")
        ax.text(0.5, 0.5, f"{fokus} tidak ada di graf", ha="center")
        ax.axis("off")
        return fig
    nodes = {fokus} | set(G.neighbors(fokus))
    # satu hop tambahan dari estate ke polres sudah termasuk neighbor
    H = G.subgraph(nodes).copy()
    pos = nx.spring_layout(H, seed=7, k=1.1)
    if ax is None:
        fig, ax = plt.subplots(figsize=(9, 6.5), facecolor="#F4F6F8")
    else:
        fig = ax.figure
    ax.set_facecolor("#F4F6F8")
    nx.draw_networkx_edges(H, pos, ax=ax, alpha=0.4, width=1.4, edge_color="#5A6672")
    for n, d in H.nodes(data=True):
        nt = d.get("ntype", "estate")
        size = 1800 if n == fokus else 700 if str(nt).startswith("polres") else 500
        nx.draw_networkx_nodes(
            H, pos, nodelist=[n], ax=ax,
            node_color=COLOR.get(nt, "#5A6672"),
            node_size=size, alpha=0.95, edgecolors="white", linewidths=1.2,
        )
    labels = {n: (d.get("label") or n)[:26] for n, d in H.nodes(data=True)}
    nx.draw_networkx_labels(H, pos, labels=labels, ax=ax, font_size=8, font_color="#1E242A")
    ax.set_title(f"Ego-network: {fokus}", fontsize=12, fontweight="bold", color="#0F2A44")
    ax.axis("off")
    fig.tight_layout()
    return fig


def plot_dashboard(G, hubs, metrics, mat, ego):
    """Satu halaman dashboard visual untuk briefing."""
    fig = plt.figure(figsize=(16, 12), facecolor="#F4F6F8")
    gs = fig.add_gridspec(2, 2, hspace=0.28, wspace=0.2)
    ax1 = fig.add_subplot(gs[0, 0])
    ax2 = fig.add_subplot(gs[0, 1])
    ax3 = fig.add_subplot(gs[1, 0])
    ax4 = fig.add_subplot(gs[1, 1])

    # reuse plotters on axes — recreate simplified inline to avoid double fig
    data = hubs[hubs["band_hub"] != "REFERENSI"].head(10).iloc[::-1]
    ax1.barh(data["label"].str.slice(0, 28), data["skor_hub"], color=[BAND_COLOR.get(b, "#2C5F7C") for b in data["band_hub"]])
    ax1.set_xlim(0, 100)
    _style(ax1, "Top skor hub")

    data2 = metrics[(metrics["bucket"] == "aktor") & (metrics["generic_flag"] != "YA")].nlargest(10, "betweenness").iloc[::-1]
    ax2.barh(data2["label"].str.slice(0, 28), data2["betweenness"], color="#2C5F7C")
    _style(ax2, "Top betweenness")

    band = hubs[hubs["band_hub"] != "REFERENSI"]["band_hub"].value_counts().reindex(["KRITIS", "TINGGI", "SEDANG", "RENDAH"]).fillna(0)
    ax3.bar(band.index, band.values, color=[BAND_COLOR[b] for b in band.index])
    _style(ax3, "Distribusi band")
    ax3.grid(axis="y", alpha=0.25)

    # mini network
    top = set(hubs[hubs["band_hub"] != "REFERENSI"].head(8)["aktor_id"]) | {"Agrinas Palma Nusantara"}
    keep = set(top)
    for n in list(top):
        if n in G:
            keep.update(list(G.neighbors(n))[:6])
    H = G.subgraph([n for n in keep if n in G]).copy()
    pos = nx.spring_layout(H, seed=42, k=0.7)
    nx.draw_networkx_edges(H, pos, ax=ax4, alpha=0.25, width=0.8)
    for n, d in H.nodes(data=True):
        nx.draw_networkx_nodes(
            H, pos, nodelist=[n], ax=ax4,
            node_color=COLOR.get(d.get("ntype", "estate"), "#5A6672"),
            node_size=280 if d.get("ntype") == "agrinas" else 120, alpha=0.9,
        )
    ax4.set_title("Cuplikan jaringan hub", fontsize=12, fontweight="bold", color="#0F2A44")
    ax4.axis("off")
    fig.suptitle("Dashboard Analisis Aktor KSO Agrinas — Unit II Harda", fontsize=14, fontweight="bold", color="#0F2A44", y=0.98)
    return fig
