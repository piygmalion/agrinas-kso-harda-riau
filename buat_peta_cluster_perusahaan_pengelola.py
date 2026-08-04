# -*- coding: utf-8 -*-
"""Peta jaringan & grafik Cluster per Perusahaan Pengelola (KSO/Agrinas).

Unit II Harda · Ditreskrimum Polda Riau

Output:
  - Peta_Cluster_Perusahaan_Pengelola_Jaringan.png
  - Peta_Cluster_Perusahaan_Pengelola_Ranking.png
  - Peta_Cluster_Perusahaan_Pengelola.html
  - output/cluster_perusahaan_pengelola/*
  - docs/pengelola/ (untuk GitHub Pages)
"""

from __future__ import annotations

import json
import math
import re
import shutil
from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import networkx as nx
import numpy as np
import openpyxl
from matplotlib.lines import Line2D

BASE = Path(r"C:\Users\Patron\Downloads\sawit lagi")
XLSX = BASE / "matriks_agrinas_kso_12_polres.xlsx"
OUT = BASE / "output" / "cluster_perusahaan_pengelola"
DOCS = BASE / "docs" / "pengelola"
OUT.mkdir(parents=True, exist_ok=True)
DOCS.mkdir(parents=True, exist_ok=True)

# Alias → nama kanonik
ALIASES = {
    "agrinas": "Agrinas Palma Nusantara",
    "agrinas palma nusantara": "Agrinas Palma Nusantara",
    "pt agrinas palma nusantara": "Agrinas Palma Nusantara",
    "dikelola langsung agrinas": "Agrinas Palma Nusantara",
    "bernas mulya mandiri": "Bernas Mulya Mandiri",
    "bernas mulya mandiri / osten panjaitan": "Bernas Mulya Mandiri",
    "mitra personal pt. bernas mulya mandiri": "Bernas Mulya Mandiri",
    "pt nusantara sawit majuma": "PT Nusantara Sawit Majuma",
    "nusantara sawit majuma": "PT Nusantara Sawit Majuma",
    "majuma": "PT Nusantara Sawit Majuma",
    "pt palma agung betuah (pab)": "PT Palma Agung Betuah (PAB)",
    "pt palma agung betuah": "PT Palma Agung Betuah (PAB)",
    "pab": "PT Palma Agung Betuah (PAB)",
    "pt ujung tanjung sejahtera": "PT Ujung Tanjung Sejahtera",
    "pt riden jaya konstruksi": "PT Riden Jaya Konstruksi",
    "berlian nusantara perkasa": "Berlian Nusantara Perkasa",
    "maju serempak": "Maju Serempak",
    "pt runggu": "PT Runggu",
    "poktan riau jaya makmur": "Poktan Riau Jaya Makmur",
    "poktan berkah tani sejahtera": "Poktan Berkah Tani Sejahtera",
    "agus s lubis": "Agus S Lubis",
    "osten panjaitan / agus s lubis": "Agus S Lubis / Osten Panjaitan",
}

GAP_LABELS = {
    "",
    "-",
    "—",
    "kso",
    "belum ada kso",
    "(kosong/tidak disebut)",
    "(dikelola langsung / belum kso)",
}

CLUSTER_DEF = {
    "Agrinas langsung": {
        "color": "#0F2A44",
        "match": lambda n, m: n == "Agrinas Palma Nusantara",
    },
    "Hub multi-lokasi": {
        "color": "#C45C26",
        "match": lambda n, m: m["jml_polres"] >= 2 and n != "Agrinas Palma Nusantara" and not m["is_gap"],
    },
    "Hub PAM / bentrok": {
        "color": "#B32D2D",
        "names": {
            "PT Nusantara Sawit Majuma",
            "PT Palma Agung Betuah (PAB)",
            "PT Ujung Tanjung Sejahtera",
            "PT Riden Jaya Konstruksi",
            "Poktan Riau Jaya Makmur",
            "Makmur Jaya Sentosa",
            "CV Makmur Jaya Sentosa",
        },
    },
    "Koperasi / Poktan": {
        "color": "#2E7D4F",
        "match": lambda n, m: (
            not m["is_gap"]
            and n != "Agrinas Palma Nusantara"
            and m["jml_polres"] < 2
            and (
                n.lower().startswith(("poktan", "koperasi", "kop.", "kud ", "gapoktan"))
                or "poktan" in n.lower()
                or "koperasi" in n.lower()
                or "kud " in n.lower()
            )
        ),
    },
    "Mitra KSO / CV / PT": {
        "color": "#2C5F7C",
        "match": lambda n, m: (not m["is_gap"] and n != "Agrinas Palma Nusantara"),
    },
    "Gap / belum KSO": {
        "color": "#8A95A1",
        "match": lambda n, m: m["is_gap"],
    },
}


def canon(name: str) -> tuple[str, bool]:
    raw = (name or "").strip()
    if not raw or raw in ("-", "—"):
        return "Belum ada KSO / dikelola langsung", True
    key = re.sub(r"\s+", " ", raw).strip().lower()
    if key in GAP_LABELS or key == "kso":
        return "Belum ada KSO / dikelola langsung", True
    # label agregat / gap dari satker (bukan badan hukum)
    # agregat satker yang menyesatkan jika dihitung sebagai badan hukum
    if any(p in key for p in (
        "nama tidak lengkap",
        "belum teridentifikasi",
        "kso penunjukan",
    )) or re.search(r"\b\d+\s*kso\b", key):
        return "Belum ada KSO / dikelola langsung", True
    if any(p in key for p in ("belum ada kso", "dikelola sendiri", "penguasaan agrinas", "self/kso")):
        if "agrinas" in key and "penunjukan" not in key:
            return "Agrinas Palma Nusantara", False
        return "Belum ada KSO / dikelola langsung", True
    if key in ("self",):
        return "Belum ada KSO / dikelola langsung", True
    if key in ALIASES:
        return ALIASES[key], False
    # soft match agrinas
    if "agrinas" in key and "palma" in key:
        return "Agrinas Palma Nusantara", False
    if key == "agrinas" or key.startswith("agrinas ") or key.endswith(" agrinas"):
        return "Agrinas Palma Nusantara", False
    if "bernas mulya" in key:
        return "Bernas Mulya Mandiri", False
    return raw, False


def load_kebun():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["kebun"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[0])
    return [dict(zip(hdr, r)) for r in rows[1:] if any(c is not None for c in r)]


def assign_cluster(name: str, meta: dict) -> str:
    # order matters
    if meta["is_gap"]:
        return "Gap / belum KSO"
    if name in CLUSTER_DEF["Hub PAM / bentrok"]["names"]:
        return "Hub PAM / bentrok"
    if name == "Agrinas Palma Nusantara":
        return "Agrinas langsung"
    if meta["jml_polres"] >= 2:
        return "Hub multi-lokasi"
    if CLUSTER_DEF["Koperasi / Poktan"]["match"](name, meta):
        return "Koperasi / Poktan"
    return "Mitra KSO / CV / PT"


def build_aggregates(kebun):
    agg = defaultdict(lambda: {
        "estates": [],
        "ha": 0.0,
        "ha_known": 0,
        "polres": set(),
        "merah": 0,
        "kuning": 0,
        "hijau": 0,
        "raw_names": set(),
        "is_gap": False,
    })

    for r in kebun:
        pengelola_raw = r.get("penerima_kso")
        name, is_gap = canon(pengelola_raw)
        d = agg[name]
        d["is_gap"] = d["is_gap"] or is_gap
        d["raw_names"].add((pengelola_raw or "").strip() or "(kosong)")
        estate = (r.get("eks_perusahaan") or "").strip() or "(estate n/a)"
        pol = (r.get("polres") or "").strip()
        if pol:
            d["polres"].add(pol)
        ha = r.get("luas_sita_ha")
        hav = None
        try:
            if ha is not None and str(ha).strip() != "":
                v = float(ha)
                if 0 < v < 200000:
                    hav = v
                    d["ha"] += v
                    d["ha_known"] += 1
        except Exception:
            pass
        kl = (r.get("klaster") or "").lower()
        if "merah" in kl:
            d["merah"] += 1
        elif "kuning" in kl:
            d["kuning"] += 1
        elif "hijau" in kl:
            d["hijau"] += 1
        d["estates"].append({
            "estate": estate,
            "polres": pol,
            "ha": hav,
            "klaster": r.get("klaster") or "",
            "lokasi": r.get("lokasi") or "",
        })

    rows = []
    for name, d in agg.items():
        meta = {
            "jml_polres": len(d["polres"]),
            "jml_estate": len(d["estates"]),
            "ha": d["ha"],
            "is_gap": d["is_gap"],
        }
        cluster = assign_cluster(name, meta)
        # if PAM also multi-lokasi, keep PAM label (already prioritized)
        rows.append({
            "pengelola": name,
            "cluster": cluster,
            "ha": d["ha"],
            "jml_estate": len(d["estates"]),
            "jml_polres": len(d["polres"]),
            "polres_list": ", ".join(sorted(d["polres"])),
            "merah": d["merah"],
            "kuning": d["kuning"],
            "hijau": d["hijau"],
            "is_gap": d["is_gap"],
            "estates": d["estates"],
            "raw_names": sorted(d["raw_names"]),
            "color": CLUSTER_DEF[cluster]["color"],
        })

    rows.sort(key=lambda x: (-x["ha"], -x["jml_estate"], x["pengelola"]))
    return rows


def build_graph(rows, top_n: int = 18):
    """Agrinas/Polda center → cluster → top pengelola → sample estates/polres."""
    G = nx.Graph()
    G.add_node("SKEMA AGRINAS–KSO", ntype="root", label="Skema\nAgrinas–KSO", ha=0, color="#0F2A44")

    # cluster nodes
    cluster_ha = defaultdict(float)
    cluster_members = defaultdict(list)
    for r in rows:
        cluster_ha[r["cluster"]] += r["ha"]
        cluster_members[r["cluster"]].append(r)

    for cname, conf in CLUSTER_DEF.items():
        G.add_node(
            cname,
            ntype="cluster",
            label=f"{cname}\n{cluster_ha[cname]:,.0f} Ha",
            ha=cluster_ha[cname],
            color=conf["color"],
            n_members=len(cluster_members[cname]),
        )
        G.add_edge("SKEMA AGRINAS–KSO", cname, layer="root-cluster", relasi="tipologi")

    # top pengelola overall + ensure PAM hubs included
    pam = CLUSTER_DEF["Hub PAM / bentrok"]["names"]
    selected = []
    seen = set()
    for r in rows:
        if r["pengelola"] in pam or len(selected) < top_n:
            if r["pengelola"] not in seen and not (r["is_gap"] and r["jml_estate"] < 2 and r["ha"] < 1000):
                # always include gap aggregate if significant
                selected.append(r)
                seen.add(r["pengelola"])
        if len(selected) >= top_n + 4:
            break
    # force include gap node if exists
    for r in rows:
        if r["is_gap"] and r["pengelola"] not in seen:
            selected.append(r)
            break

    for r in selected:
        pid = r["pengelola"]
        G.add_node(
            pid,
            ntype="pengelola",
            label=f"{pid[:30] + ('…' if len(pid)>30 else '')}\n{r['ha']:,.0f} Ha · {r['jml_estate']} kebun",
            ha=r["ha"],
            color=r["color"],
            cluster=r["cluster"],
            jml_estate=r["jml_estate"],
            jml_polres=r["jml_polres"],
            polres_list=r["polres_list"],
        )
        G.add_edge(r["cluster"], pid, layer="cluster-pengelola", relasi="anggota_cluster")

        # link sample estates (top 3 by ha)
        estates = sorted(
            [e for e in r["estates"] if e.get("ha")],
            key=lambda x: x["ha"] or 0,
            reverse=True,
        )[:3]
        if not estates:
            estates = r["estates"][:2]
        for e in estates:
            eid = f"{pid}::{e['estate']}"
            kl = (e.get("klaster") or "").lower()
            if "merah" in kl:
                ecol = "#B32D2D"
            elif "kuning" in kl:
                ecol = "#C48A14"
            elif "hijau" in kl:
                ecol = "#2E7D4F"
            else:
                ecol = "#7A8792"
            G.add_node(
                eid,
                ntype="estate",
                label=(e["estate"][:26] + ("…" if len(e["estate"]) > 26 else "")),
                ha=e.get("ha") or 0,
                color=ecol,
                polres=e.get("polres") or "",
            )
            G.add_edge(pid, eid, layer="pengelola-estate", relasi="kelola")

            # polres node (shared)
            pol = e.get("polres")
            if pol:
                if pol not in G:
                    G.add_node(pol, ntype="polres", label=pol, ha=0, color="#B88A3D")
                G.add_edge(eid, pol, layer="estate-polres", relasi="wilayah")

    return G, selected


def layout(G):
    pos = {"SKEMA AGRINAS–KSO": (0.0, 0.0)}
    clusters = [n for n, d in G.nodes(data=True) if d.get("ntype") == "cluster"]
    # order by ha
    clusters.sort(key=lambda n: G.nodes[n].get("ha", 0), reverse=True)
    for i, c in enumerate(clusters):
        ang = -math.pi / 2 + i * (2 * math.pi / max(len(clusters), 1))
        pos[c] = (2.3 * math.cos(ang), 2.3 * math.sin(ang))

    for c in clusters:
        members = [n for n in G.neighbors(c) if G.nodes[n].get("ntype") == "pengelola"]
        cx, cy = pos[c]
        base = math.atan2(cy, cx)
        for j, m in enumerate(members):
            spread = (j - (len(members) - 1) / 2) * 0.42
            ang = base + spread
            pos[m] = (4.2 * math.cos(ang), 4.2 * math.sin(ang))
            estates = [n for n in G.neighbors(m) if G.nodes[n].get("ntype") == "estate"]
            for k, e in enumerate(estates):
                e_spread = (k - (len(estates) - 1) / 2) * 0.22
                pos[e] = (5.5 * math.cos(ang + e_spread), 5.5 * math.sin(ang + e_spread))

    # polres lightly placed
    for n, d in G.nodes(data=True):
        if d.get("ntype") == "polres" and n not in pos:
            pos[n] = (6.2, 0.0)

    return nx.spring_layout(G, pos=pos, seed=42, k=0.85, iterations=45)


def draw_network(G, path: Path):
    pos = layout(G)
    fig, ax = plt.subplots(figsize=(18, 14), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")

    for layer, style, alpha, w in [
        ("root-cluster", "solid", 0.45, 2.2),
        ("cluster-pengelola", "solid", 0.4, 1.6),
        ("pengelola-estate", "dashed", 0.3, 1.0),
        ("estate-polres", "dotted", 0.22, 0.8),
    ]:
        el = [(u, v) for u, v, d in G.edges(data=True) if d.get("layer") == layer]
        nx.draw_networkx_edges(G, pos, edgelist=el, ax=ax, style=style, alpha=alpha, width=w, edge_color="#5A6672")

    groups = defaultdict(list)
    for n, d in G.nodes(data=True):
        groups[d.get("ntype")].append(n)

    def sizes(nodes, base, scale):
        return [base + min(G.nodes[n].get("ha") or 0, 50000) / scale for n in nodes]

    if "root" in groups:
        nx.draw_networkx_nodes(G, pos, nodelist=groups["root"], ax=ax, node_color="#0F2A44",
                               node_size=3800, edgecolors="white", linewidths=2)
    if "cluster" in groups:
        nx.draw_networkx_nodes(
            G, pos, nodelist=groups["cluster"], ax=ax,
            node_color=[G.nodes[n]["color"] for n in groups["cluster"]],
            node_size=sizes(groups["cluster"], 1400, 40),
            edgecolors="white", linewidths=1.5, alpha=0.95,
        )
    if "pengelola" in groups:
        nx.draw_networkx_nodes(
            G, pos, nodelist=groups["pengelola"], ax=ax,
            node_color=[G.nodes[n]["color"] for n in groups["pengelola"]],
            node_size=sizes(groups["pengelola"], 700, 25),
            edgecolors="white", linewidths=1.2, alpha=0.95,
        )
    if "estate" in groups:
        nx.draw_networkx_nodes(
            G, pos, nodelist=groups["estate"], ax=ax,
            node_color=[G.nodes[n].get("color", "#7A8792") for n in groups["estate"]],
            node_size=[200 + min(G.nodes[n].get("ha") or 0, 8000) / 30 for n in groups["estate"]],
            node_shape="s", edgecolors="white", linewidths=0.6, alpha=0.85,
        )
    if "polres" in groups:
        nx.draw_networkx_nodes(
            G, pos, nodelist=groups["polres"], ax=ax,
            node_color="#B88A3D", node_size=500, node_shape="h",
            edgecolors="white", linewidths=1.0, alpha=0.9,
        )

    labels = {}
    for n, d in G.nodes(data=True):
        nt = d.get("ntype")
        if nt in ("root", "cluster", "pengelola"):
            labels[n] = d.get("label", n)
        elif nt == "polres":
            labels[n] = n
        elif nt == "estate" and (d.get("ha") or 0) >= 5000:
            labels[n] = d.get("label", n)
    nx.draw_networkx_labels(G, pos, labels=labels, ax=ax, font_size=7.2, font_color="#1E242A")

    ax.set_title(
        "Peta Jaringan Cluster per Perusahaan Pengelola\nSkema Agrinas–KSO · Wilayah Hukum Polda Riau",
        fontsize=16, fontweight="bold", color="#0F2A44", pad=14,
    )
    ax.text(
        0.5, -0.02,
        "Unit II Harda · Ditreskrimum Polda Riau  |  Node size ≈ luas sitaan dikelola  |  "
        "Kotak=estate · hexagon=Polres  |  Nama digabung (alias Agrinas/Bernas/dll.)  |  Agustus 2026",
        transform=ax.transAxes, ha="center", fontsize=9, color="#5A6672",
    )
    legend = [
        mpatches.Patch(color=conf["color"], label=name) for name, conf in CLUSTER_DEF.items()
    ] + [
        Line2D([0], [0], color="#5A6672", lw=1, ls="--", label="Pengelola → estate"),
        Line2D([0], [0], color="#5A6672", lw=1, ls=":", label="Estate → Polres"),
    ]
    ax.legend(handles=legend, loc="lower left", fontsize=8, framealpha=0.95, ncol=2)
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"OK PNG: {path}")


def draw_ranking(rows, path: Path, top_n: int = 20):
    data = [r for r in rows if not (r["is_gap"] and r["ha"] < 500)][:top_n]
    # put gap aggregate if in top by ha
    gap = next((r for r in rows if r["is_gap"]), None)
    if gap and gap not in data and gap["ha"] > 0:
        data = data[: top_n - 1] + [gap]
        data.sort(key=lambda x: -x["ha"])

    fig, ax = plt.subplots(figsize=(12, 8), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")
    labels = [r["pengelola"][:40] + ("…" if len(r["pengelola"]) > 40 else "") for r in data]
    vals = [r["ha"] for r in data]
    colors = [r["color"] for r in data]
    y = np.arange(len(data))
    ax.barh(y, vals, color=colors, height=0.72)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    for i, r in enumerate(data):
        ax.text(
            r["ha"] + (max(vals) * 0.01 if vals else 1),
            i,
            f"{r['ha']:,.0f} Ha · {r['jml_estate']} kebun · {r['jml_polres']} Polres · {r['cluster']}",
            va="center", fontsize=7.5, color="#1E242A",
        )
    ax.set_xlabel("Luas sitaan PKH dikelola (Ha)")
    ax.set_xlim(0, (max(vals) * 1.45) if vals else 1)
    ax.set_title(
        "Ranking Perusahaan Pengelola — Konsentrasi Lahan Sitaan PKH\nUnit II Harda · Ditreskrimum Polda Riau",
        fontsize=13, fontweight="bold", color="#0F2A44",
    )
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.grid(axis="x", alpha=0.25)
    legend = [mpatches.Patch(color=c["color"], label=n) for n, c in CLUSTER_DEF.items()]
    ax.legend(handles=legend, loc="lower right", fontsize=8, framealpha=0.95)
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"OK PNG: {path}")


def draw_cluster_pie(rows, path: Path):
    from collections import OrderedDict
    tot = OrderedDict()
    for r in rows:
        tot[r["cluster"]] = tot.get(r["cluster"], 0.0) + r["ha"]
    labels = list(tot.keys())
    sizes = list(tot.values())
    colors = [CLUSTER_DEF[k]["color"] for k in labels]
    fig, ax = plt.subplots(figsize=(9, 7), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")
    explode = [0.03] * len(labels)
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, colors=colors, autopct=lambda p: f"{p:.1f}%" if p >= 3 else "",
        startangle=120, explode=explode, pctdistance=0.75,
        wedgeprops=dict(width=0.45, edgecolor="white"),
    )
    for t in autotexts:
        t.set_fontsize(8)
        t.set_color("#1E242A")
    ax.legend(
        [f"{l} ({v:,.0f} Ha)" for l, v in tot.items()],
        loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=9,
    )
    ax.set_title(
        "Komposisi Cluster Pengelola menurut Luas Sitaan PKH",
        fontsize=13, fontweight="bold", color="#0F2A44",
    )
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"OK PNG: {path}")


def write_html(G, rows, selected, path: Path):
    nodes, edges = [], []
    idmap = {n: i for i, n in enumerate(G.nodes())}
    for n, d in G.nodes(data=True):
        nt = d.get("ntype")
        shape = {"root": "dot", "cluster": "dot", "pengelola": "box", "estate": "diamond", "polres": "hexagon"}.get(nt, "dot")
        size = {
            "root": 38,
            "cluster": 26,
            "pengelola": 14 + min(d.get("ha") or 0, 40000) / 2500,
            "estate": 8 + min(d.get("ha") or 0, 8000) / 900,
            "polres": 14,
        }.get(nt, 10)
        title = [str(n) if nt != "estate" else n.split("::", 1)[-1], f"tipe: {nt}", f"Ha: {d.get('ha', 0):,.0f}"]
        if d.get("cluster"):
            title.append(f"cluster: {d['cluster']}")
        if d.get("polres_list"):
            title.append(f"polres: {d['polres_list']}")
        nodes.append({
            "id": idmap[n],
            "label": (d.get("label") or n).replace("\n", " ")[:42],
            "title": " | ".join(title),
            "color": d.get("color", "#5A6672"),
            "shape": shape,
            "size": size,
            "font": {"color": "#1E242A", "size": 11 if nt != "estate" else 10},
        })
    for u, v, d in G.edges(data=True):
        edges.append({
            "from": idmap[u], "to": idmap[v],
            "dashes": d.get("layer") in ("pengelola-estate", "estate-polres"),
            "width": 2.0 if "cluster" in (d.get("layer") or "") else 1.2,
            "color": {"color": "#6A7680"},
            "title": f"{d.get('relasi')} ({d.get('layer')})",
        })

    rank_html = "".join(
        f'<div class="hub" style="border-left-color:{r["color"]}"><b>{r["pengelola"]}</b><br/>'
        f'<small>{r["ha"]:,.0f} Ha · {r["jml_estate"]} kebun · {r["jml_polres"]} Polres · {r["cluster"]}</small></div>'
        for r in rows[:15]
    )

    html = f"""<!DOCTYPE html>
<html lang="id"><head>
<meta charset="utf-8"/>
<title>Cluster per Perusahaan Pengelola — Agrinas–KSO</title>
<script src="https://unpkg.com/vis-network@9.1.9/standalone/umd/vis-network.min.js"></script>
<style>
body{{margin:0;font-family:Segoe UI,Calibri,sans-serif;background:#0F2A44;color:#fff;display:grid;grid-template-rows:auto 1fr;height:100vh}}
header{{padding:14px 20px;border-bottom:3px solid #C45C26}}
h1{{margin:0;font-size:1.15rem}} .sub{{color:#B8C5D0;font-size:.82rem;margin-top:4px}}
.shell{{display:grid;grid-template-columns:1fr 340px;min-height:0}}
#net{{background:#F4F6F8}}
aside{{padding:14px;overflow:auto;background:rgba(15,42,68,.96)}}
h2{{font-size:.75rem;color:#C45C26;text-transform:uppercase;letter-spacing:.08em}}
.hub{{background:rgba(0,0,0,.22);border-left:4px solid #C45C26;padding:8px 10px;margin:6px 0;border-radius:0 8px 8px 0;font-size:12px}}
.note{{font-size:11px;color:#9AA8B5;line-height:1.45}}
a{{color:#F0C27A}}
</style></head><body>
<header>
  <h1>Cluster per Perusahaan Pengelola — Lahan Sitaan PKH</h1>
  <div class="sub">Unit II Harda · Ditreskrimum Polda Riau · Skema → Cluster → Pengelola → Estate → Polres</div>
</header>
<div class="shell">
  <div id="net"></div>
  <aside>
    <h2>Top pengelola (Ha)</h2>
    {rank_html}
    <h2>Cara baca</h2>
    <div class="note">
      Cluster: Agrinas langsung · Hub multi-lokasi · Hub PAM/bentrok · Koperasi/Poktan · Mitra KSO · Gap/belum KSO.<br/>
      Ukuran node ≈ luas sitaan yang dikelola.<br/>
      Nama pengelola sudah dinormalisasi (alias Agrinas/Bernas/dll.).
    </div>
  </aside>
</div>
<script>
const nodes=new vis.DataSet({json.dumps(nodes, ensure_ascii=False)});
const edges=new vis.DataSet({json.dumps(edges, ensure_ascii=False)});
new vis.Network(document.getElementById('net'), {{nodes, edges}}, {{
  physics:{{barnesHut:{{gravitationalConstant:-13000, springLength:130, springConstant:0.03}}, stabilization:{{iterations:170}}}},
  interaction:{{hover:true, navigationButtons:true}}
}});
</script></body></html>"""
    path.write_text(html, encoding="utf-8")
    print(f"OK HTML: {path}")


def export_csv(rows):
    import csv
    p1 = OUT / "tabel_cluster_perusahaan_pengelola.csv"
    with p1.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([
            "pengelola", "cluster", "ha_sitaan", "jml_estate", "jml_polres",
            "polres_list", "merah", "kuning", "hijau", "alias_mentah",
        ])
        for r in rows:
            w.writerow([
                r["pengelola"], r["cluster"], f'{r["ha"]:.2f}', r["jml_estate"], r["jml_polres"],
                r["polres_list"], r["merah"], r["kuning"], r["hijau"], " | ".join(r["raw_names"][:5]),
            ])
    p2 = OUT / "tabel_ringkas_cluster_pengelola.csv"
    from collections import defaultdict
    roll = defaultdict(lambda: {"ha": 0.0, "n": 0, "estate": 0})
    for r in rows:
        roll[r["cluster"]]["ha"] += r["ha"]
        roll[r["cluster"]]["n"] += 1
        roll[r["cluster"]]["estate"] += r["jml_estate"]
    with p2.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["cluster", "jml_pengelola", "jml_estate", "ha_total"])
        for c, v in sorted(roll.items(), key=lambda x: -x[1]["ha"]):
            w.writerow([c, v["n"], v["estate"], f'{v["ha"]:.2f}'])
    print(f"OK CSV: {p1.name}, {p2.name}")


def sync_docs():
    for f in [
        "Peta_Cluster_Perusahaan_Pengelola_Jaringan.png",
        "Peta_Cluster_Perusahaan_Pengelola_Ranking.png",
        "Peta_Cluster_Perusahaan_Pengelola_Komposisi.png",
        "Peta_Cluster_Perusahaan_Pengelola.html",
    ]:
        src = BASE / f
        if src.exists():
            shutil.copy2(src, DOCS / ( "index.html" if f.endswith(".html") else f))
    # copy html as index
    html = BASE / "Peta_Cluster_Perusahaan_Pengelola.html"
    if html.exists():
        shutil.copy2(html, DOCS / "index.html")
    for f in OUT.glob("*.csv"):
        shutil.copy2(f, DOCS / f.name)
    for f in OUT.glob("*.gexf"):
        shutil.copy2(f, DOCS / f.name)
    print(f"OK docs sync: {DOCS}")


def update_portal():
    portal = BASE / "docs" / "index.html"
    if not portal.exists():
        return
    text = portal.read_text(encoding="utf-8")
    card = """
      <a class="card" href="./pengelola/">
        <span class="tag">INTERAKTIF</span>
        <h2>Cluster per Perusahaan Pengelola</h2>
        <p>
          Tipologi pengelola Agrinas/KSO/PAM/poktan, ranking luas sitaan,
          dan jaringan pengelola → estate → Polres.
        </p>
      </a>"""
    if "./pengelola/" not in text:
        # insert after cluster card block
        marker = 'href="./cluster/"'
        idx = text.find(marker)
        if idx != -1:
            # find end of that </a>
            end = text.find("</a>", idx)
            if end != -1:
                text = text[: end + 4] + "\n" + card + text[end + 4 :]
                portal.write_text(text, encoding="utf-8")
                print("OK portal updated")
                return
    print("Portal already has pengelola link or marker missing")


def main():
    kebun = load_kebun()
    rows = build_aggregates(kebun)
    G, selected = build_graph(rows, top_n=18)

    print("=== TOP PENGELOLA ===")
    for r in rows[:15]:
        print(
            f"{r['ha']:10,.1f} Ha | {r['jml_estate']:2d} kebun | {r['jml_polres']} Polres | "
            f"{r['cluster'][:22]:22s} | {r['pengelola'][:50]}"
        )

    draw_network(G, BASE / "Peta_Cluster_Perusahaan_Pengelola_Jaringan.png")
    draw_ranking(rows, BASE / "Peta_Cluster_Perusahaan_Pengelola_Ranking.png")
    draw_cluster_pie(rows, BASE / "Peta_Cluster_Perusahaan_Pengelola_Komposisi.png")
    write_html(G, rows, selected, BASE / "Peta_Cluster_Perusahaan_Pengelola.html")

    # output copies
    shutil.copy2(BASE / "Peta_Cluster_Perusahaan_Pengelola_Jaringan.png", OUT / "fig_jaringan_pengelola.png")
    shutil.copy2(BASE / "Peta_Cluster_Perusahaan_Pengelola_Ranking.png", OUT / "fig_ranking_pengelola.png")
    shutil.copy2(BASE / "Peta_Cluster_Perusahaan_Pengelola_Komposisi.png", OUT / "fig_komposisi_cluster_pengelola.png")
    shutil.copy2(BASE / "Peta_Cluster_Perusahaan_Pengelola.html", OUT / "Peta_Cluster_Perusahaan_Pengelola.html")

    export_csv(rows)
    H = G.copy()
    for n, d in H.nodes(data=True):
        for k, v in list(d.items()):
            d[k] = "" if v is None else str(v)
    nx.write_gexf(H, OUT / "cluster_perusahaan_pengelola.gexf")
    sync_docs()
    update_portal()
    print(f"Graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")
    print(f"Pengelola unik (kanonik): {len(rows)}")


if __name__ == "__main__":
    main()
