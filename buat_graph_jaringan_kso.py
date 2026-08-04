# -*- coding: utf-8 -*-
"""Peta jaringan KSO Agrinas — Unit II Harda Ditreskrimum Polda Riau.

Output:
  - Peta_Jaringan_KSO_Agrinas.png          (graf utama)
  - Peta_Jaringan_KSO_Agrinas_hub.png      (fokus hub risiko)
  - Peta_Jaringan_KSO_Agrinas.html         (interaktif)
  - jaringan_kso_nodes.csv / edges.csv
  - jaringan_kso_agrinas.gexf / .graphml
"""

from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import networkx as nx
import openpyxl
from matplotlib.lines import Line2D

BASE = Path(r"C:\Users\Patron\Downloads\sawit lagi")
XLSX_R = BASE / "analisis_lanjutan_5_prioritas.xlsx"
XLSX_M = BASE / "matriks_agrinas_kso_12_polres.xlsx"

# Polres risk band (from R1)
POLRES_BAND = {
    "Rohul": "KRITIS",
    "Rohil": "TINGGI",
    "Bengkalis": "TINGGI",
    "Kampar": "SEDANG",
    "Inhu": "SEDANG",
    "Dumai": "SEDANG",
    "Kuansing": "SEDANG",
    "Inhil": "RENDAH",
    "Pelalawan": "RENDAH",
    "Siak": "RENDAH",
    "Kep. Meranti": "RENDAH",
    "Pekanbaru": "RENDAH",
}

# Known PAM non-BUJP / bentrok hubs
PAM_HUBS = {
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

# Generic/noise actor labels to soften
GENERIC_ACTORS = {"KSO", "Belum ada KSO"}

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
    "edge_kso": "#5A6672",
    "edge_kelola": "#0F2A44",
    "edge_polres": "#A0AAB4",
}


def load_sheet(path, name):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])
    return [dict(zip(headers, r)) for r in rows[1:] if any(c is not None for c in r)]


def short(name: str, n: int = 28) -> str:
    name = (name or "").strip()
    if len(name) <= n:
        return name
    return name[: n - 1] + "…"


def build_graph():
    edges_r2 = load_sheet(XLSX_R, "R2_edges")
    hubs = load_sheet(XLSX_R, "R2_hub_aktor")
    kebun = load_sheet(XLSX_M, "kebun")
    pelalawan = load_sheet(XLSX_R, "R5_pelalawan_kso")
    skor_kebun = load_sheet(XLSX_R, "R1_skor_kebun")

    # estate -> klaster lookup
    klaster = {}
    for r in kebun:
        e = (r.get("eks_perusahaan") or "").strip()
        if e:
            klaster[e] = (r.get("klaster") or "").strip()
    for r in skor_kebun:
        e = (r.get("eks_perusahaan") or "").strip()
        if e and e not in klaster:
            klaster[e] = (r.get("klaster_intelkam") or "").strip()

    hub_flag = {h["aktor"]: h for h in hubs}

    G = nx.Graph()

    # Central Agrinas node
    G.add_node(
        "Agrinas Palma Nusantara",
        ntype="agrinas",
        label="Agrinas Palma\nNusantara",
        hub_risiko=True,
        jml_polres=5,
        jml_estate=10,
    )

    def ensure_aktor(a: str):
        if a in G and G.nodes[a].get("ntype") in ("aktor", "agrinas", "aktor_pam", "aktor_multi", "aktor_generic"):
            return
        if a == "Agrinas Palma Nusantara":
            return
        info = hub_flag.get(a, {})
        if a in PAM_HUBS:
            ntype = "aktor_pam"
        elif a in GENERIC_ACTORS:
            ntype = "aktor_generic"
        elif info.get("flag_hub_risiko") == "YA" and (info.get("jml_polres") or 0) >= 2:
            ntype = "aktor_multi"
        elif info.get("flag_hub_risiko") == "YA":
            ntype = "aktor_pam" if "pam" in str(info.get("peran", "")).lower() or a in PAM_HUBS else "aktor_multi"
        else:
            ntype = "aktor"
        G.add_node(
            a,
            ntype=ntype,
            label=short(a, 26),
            hub_risiko=info.get("flag_hub_risiko") == "YA" or a in PAM_HUBS,
            jml_polres=info.get("jml_polres") or 1,
            jml_estate=info.get("jml_estate_terkait") or 0,
            peran=info.get("peran") or "",
        )

    def ensure_estate(e: str, pol: str):
        if not e:
            return
        if e in G and G.nodes[e].get("ntype", "").startswith("estate"):
            return
        k = (klaster.get(e) or "").lower()
        if "merah" in k:
            ntype = "estate_merah"
        elif "kuning" in k:
            ntype = "estate_kuning"
        elif "hijau" in k:
            ntype = "estate_hijau"
        else:
            ntype = "estate"
        G.add_node(e, ntype=ntype, label=short(e, 24), polres=pol, klaster=klaster.get(e, ""))

    def ensure_polres(p: str):
        if not p or (p in G and G.nodes[p].get("ntype", "").startswith("polres")):
            return
        band = POLRES_BAND.get(p, "RENDAH")
        ntype = {
            "KRITIS": "polres_kritis",
            "TINGGI": "polres_tinggi",
            "SEDANG": "polres_sedang",
            "RENDAH": "polres_rendah",
        }[band]
        G.add_node(p, ntype=ntype, label=p, band=band)

    # Primary edges from R2
    for r in edges_r2:
        a = (r.get("source_aktor") or "").strip()
        p = (r.get("target_polres") or "").strip()
        e = (r.get("estate") or "").strip()
        rel = (r.get("relasi") or "penerima_kso").strip()
        if not a or not p:
            continue
        ensure_aktor(a)
        ensure_polres(p)
        if e:
            ensure_estate(e, p)
            G.add_edge(a, e, relasi=rel, layer="aktor-estate")
            G.add_edge(e, p, relasi="wilayah", layer="estate-polres")
        else:
            G.add_edge(a, p, relasi=rel, layer="aktor-polres")
        # Link all KSO/aktor to Agrinas hub (titip kelola)
        if a != "Agrinas Palma Nusantara":
            G.add_edge("Agrinas Palma Nusantara", a, relasi="skema_agrinas", layer="agrinas-aktor")

    # Enrich Pelalawan KSO not fully in R2 edges
    for r in pelalawan:
        e = (r.get("eks_perusahaan") or "").strip()
        a = (r.get("penerima_kso") or "").strip()
        if not e or not a or a == "BELUM TERIDENTIFIKASI":
            continue
        if a.lower() == "agrinas":
            a = "Agrinas Palma Nusantara"
        ensure_aktor(a)
        ensure_polres("Pelalawan")
        ensure_estate(e, "Pelalawan")
        G.add_edge(a, e, relasi="penerima_kso", layer="aktor-estate")
        G.add_edge(e, "Pelalawan", relasi="wilayah", layer="estate-polres")
        if a != "Agrinas Palma Nusantara":
            G.add_edge("Agrinas Palma Nusantara", a, relasi="skema_agrinas", layer="agrinas-aktor")

    # Force PAM hubs even if sparse estate
    for a, pol in [
        ("Makmur Jaya Sentosa", "Kampar"),
        ("Torus", "Rohul"),
        ("Togos", "Rohul"),
        ("Togas", "Rohul"),
    ]:
        ensure_aktor(a)
        ensure_polres(pol)
        G.add_edge(a, pol, relasi="pam_or_kso_flagged", layer="aktor-polres")
        G.add_edge("Agrinas Palma Nusantara", a, relasi="skema_agrinas", layer="agrinas-aktor")

    return G, hubs


def node_color(ntype: str) -> str:
    return COLOR.get(ntype, "#5A6672")


def node_size(G, n) -> float:
    d = G.degree(n)
    ntype = G.nodes[n].get("ntype", "")
    if ntype == "agrinas":
        return 3200
    if ntype.startswith("polres"):
        return 1400 + d * 80
    if ntype.startswith("aktor"):
        base = 900 + d * 120
        if G.nodes[n].get("hub_risiko"):
            base += 300
        return base
    return 500 + d * 60


def layout_positions(G):
    """Hierarchical-ish spring: Agrinas center, polres outer ring, actors mid, estates near actors."""
    # Seed positions by type
    polres = [n for n, d in G.nodes(data=True) if d.get("ntype", "").startswith("polres")]
    aktor = [n for n, d in G.nodes(data=True) if d.get("ntype", "").startswith("aktor") or d.get("ntype") == "agrinas"]
    estate = [n for n, d in G.nodes(data=True) if d.get("ntype", "").startswith("estate")]

    pos = {}
    # Place polres on outer circle by risk band order
    order = sorted(
        polres,
        key=lambda p: {"KRITIS": 0, "TINGGI": 1, "SEDANG": 2, "RENDAH": 3}.get(G.nodes[p].get("band"), 9),
    )
    import math

    for i, p in enumerate(order):
        ang = 2 * math.pi * i / max(len(order), 1) - math.pi / 2
        pos[p] = (3.6 * math.cos(ang), 3.6 * math.sin(ang))

    pos["Agrinas Palma Nusantara"] = (0.0, 0.0)

    # Actors in mid ring near their polres neighbors
    for a in aktor:
        if a == "Agrinas Palma Nusantara":
            continue
        neigh_p = [x for x in G.neighbors(a) if G.nodes[x].get("ntype", "").startswith("polres")]
        if not neigh_p:
            # use estate's polres
            for e in G.neighbors(a):
                if G.nodes[e].get("ntype", "").startswith("estate"):
                    pp = G.nodes[e].get("polres")
                    if pp in pos:
                        neigh_p.append(pp)
        if neigh_p and neigh_p[0] in pos:
            px, py = pos[neigh_p[0]]
            # pull toward center
            pos[a] = (px * 0.55, py * 0.55)
        else:
            pos[a] = (0.8, 0.8)

    # Estates near their actor
    for e in estate:
        acts = [x for x in G.neighbors(e) if G.nodes[x].get("ntype", "").startswith("aktor") or x == "Agrinas Palma Nusantara"]
        pols = [x for x in G.neighbors(e) if G.nodes[x].get("ntype", "").startswith("polres")]
        if acts and acts[0] in pos:
            ax, ay = pos[acts[0]]
            if pols and pols[0] in pos:
                px, py = pos[pols[0]]
                pos[e] = (ax * 0.55 + px * 0.35, ay * 0.55 + py * 0.35)
            else:
                pos[e] = (ax * 1.15, ay * 1.15)
        elif pols and pols[0] in pos:
            px, py = pos[pols[0]]
            pos[e] = (px * 0.75, py * 0.75)
        else:
            pos[e] = (1.0, -1.0)

    # Refine with spring
    pos = nx.spring_layout(G, pos=pos, seed=42, k=0.55, iterations=55, weight=None)
    return pos


def draw_network(G, out_png: Path, title: str, focus_hubs: bool = False):
    if focus_hubs:
        keep = set()
        for n, d in G.nodes(data=True):
            if d.get("ntype") == "agrinas" or d.get("hub_risiko") or d.get("ntype", "").startswith("polres"):
                keep.add(n)
        # include neighbors of hub risiko (estates)
        add = set()
        for n in list(keep):
            if G.nodes[n].get("hub_risiko") or G.nodes[n].get("ntype") == "agrinas":
                add.update(G.neighbors(n))
        keep |= add
        H = G.subgraph(keep).copy()
    else:
        H = G

    pos = layout_positions(H)
    fig, ax = plt.subplots(figsize=(18, 14), facecolor="#F4F6F8")
    ax.set_facecolor("#F4F6F8")

    # Edges by layer
    e_agr = [(u, v) for u, v, d in H.edges(data=True) if d.get("layer") == "agrinas-aktor"]
    e_ae = [(u, v) for u, v, d in H.edges(data=True) if d.get("layer") == "aktor-estate"]
    e_ep = [(u, v) for u, v, d in H.edges(data=True) if d.get("layer") == "estate-polres"]
    e_ap = [(u, v) for u, v, d in H.edges(data=True) if d.get("layer") == "aktor-polres"]

    nx.draw_networkx_edges(H, pos, edgelist=e_agr, ax=ax, width=1.6, alpha=0.35, edge_color=COLOR["edge_kelola"])
    nx.draw_networkx_edges(H, pos, edgelist=e_ae, ax=ax, width=1.4, alpha=0.55, edge_color=COLOR["edge_kso"])
    nx.draw_networkx_edges(H, pos, edgelist=e_ep, ax=ax, width=1.0, alpha=0.35, edge_color=COLOR["edge_polres"], style="dashed")
    nx.draw_networkx_edges(H, pos, edgelist=e_ap, ax=ax, width=1.2, alpha=0.45, edge_color="#8A95A1", style="dotted")

    # Nodes by type
    groups = defaultdict(list)
    for n, d in H.nodes(data=True):
        groups[d.get("ntype", "estate")].append(n)

    for ntype, nodes in groups.items():
        nx.draw_networkx_nodes(
            H,
            pos,
            nodelist=nodes,
            ax=ax,
            node_color=node_color(ntype),
            node_size=[node_size(H, n) for n in nodes],
            alpha=0.92,
            linewidths=1.2,
            edgecolors="white",
        )

    labels = {n: H.nodes[n].get("label", short(n, 22)) for n in H.nodes()}
    # Only label agrinas, polres, hub actors, merah estates (reduce clutter)
    show = {}
    for n, d in H.nodes(data=True):
        nt = d.get("ntype", "")
        if nt == "agrinas" or nt.startswith("polres") or d.get("hub_risiko") or nt == "estate_merah":
            show[n] = labels[n]
        elif focus_hubs and nt.startswith("estate"):
            show[n] = labels[n]
        elif not focus_hubs and nt.startswith("aktor") and H.degree(n) >= 3:
            show[n] = labels[n]
    nx.draw_networkx_labels(H, pos, labels=show, ax=ax, font_size=7.5, font_color="#1E242A", font_family="DejaVu Sans")

    ax.set_title(title, fontsize=16, fontweight="bold", color="#0F2A44", pad=14)
    ax.text(
        0.5,
        -0.02,
        "Unit II Harda · Ditreskrimum Polda Riau  |  Sumber: R2_edges + matriks kebun + Pelalawan KSO  |  Agustus 2026",
        transform=ax.transAxes,
        ha="center",
        fontsize=9,
        color="#5A6672",
    )

    legend = [
        mpatches.Patch(color=COLOR["agrinas"], label="Agrinas (pusat skema)"),
        mpatches.Patch(color=COLOR["aktor_pam"], label="Hub PAM/KSO risiko (bentrok)"),
        mpatches.Patch(color=COLOR["aktor_multi"], label="Hub multi-lokasi"),
        mpatches.Patch(color=COLOR["aktor"], label="Penerima KSO"),
        mpatches.Patch(color=COLOR["estate_merah"], label="Estate merah"),
        mpatches.Patch(color=COLOR["estate_kuning"], label="Estate kuning"),
        mpatches.Patch(color=COLOR["estate_hijau"], label="Estate hijau"),
        mpatches.Patch(color=COLOR["polres_kritis"], label="Polres KRITIS"),
        mpatches.Patch(color=COLOR["polres_tinggi"], label="Polres TINGGI"),
        Line2D([0], [0], color=COLOR["edge_kelola"], lw=2, label="Agrinas ↔ aktor"),
        Line2D([0], [0], color=COLOR["edge_kso"], lw=2, label="Aktor ↔ estate"),
        Line2D([0], [0], color=COLOR["edge_polres"], lw=1.5, ls="--", label="Estate ↔ Polres"),
    ]
    ax.legend(handles=legend, loc="lower left", fontsize=8, framealpha=0.95, ncol=2)
    ax.set_axis_off()
    fig.tight_layout()
    fig.savefig(out_png, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"OK PNG: {out_png}")


def export_tables(G):
    import csv

    nodes_path = BASE / "jaringan_kso_nodes.csv"
    edges_path = BASE / "jaringan_kso_edges.csv"
    with nodes_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["id", "label", "ntype", "hub_risiko", "degree", "polres", "klaster", "band", "jml_polres", "jml_estate"])
        for n, d in sorted(G.nodes(data=True), key=lambda x: (-x[1].get("hub_risiko", False), x[0])):
            w.writerow(
                [
                    n,
                    d.get("label", n),
                    d.get("ntype"),
                    "YA" if d.get("hub_risiko") else "TIDAK",
                    G.degree(n),
                    d.get("polres", ""),
                    d.get("klaster", ""),
                    d.get("band", ""),
                    d.get("jml_polres", ""),
                    d.get("jml_estate", ""),
                ]
            )
    with edges_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["source", "target", "relasi", "layer"])
        for u, v, d in G.edges(data=True):
            w.writerow([u, v, d.get("relasi", ""), d.get("layer", "")])
    print(f"OK CSV: {nodes_path.name}, {edges_path.name}")


def export_graph_formats(G):
    # sanitize for GEXF (networkx needs string attrs mostly)
    H = G.copy()
    for n, d in H.nodes(data=True):
        for k, v in list(d.items()):
            if isinstance(v, bool):
                d[k] = "true" if v else "false"
            elif v is None:
                d[k] = ""
            else:
                d[k] = str(v)
    gexf = BASE / "jaringan_kso_agrinas.gexf"
    graphml = BASE / "jaringan_kso_agrinas.graphml"
    nx.write_gexf(H, gexf)
    nx.write_graphml(H, graphml)
    print(f"OK GEXF/GraphML: {gexf.name}, {graphml.name}")


def write_html(G, out_html: Path):
    """Self-contained interactive graph with vis-network CDN."""
    nodes = []
    edges = []
    group_map = {
        "agrinas": {"color": COLOR["agrinas"], "shape": "dot", "size": 36},
        "aktor_pam": {"color": COLOR["aktor_pam"], "shape": "dot", "size": 24},
        "aktor_multi": {"color": COLOR["aktor_multi"], "shape": "dot", "size": 22},
        "aktor": {"color": COLOR["aktor"], "shape": "dot", "size": 18},
        "aktor_generic": {"color": COLOR["aktor_generic"], "shape": "dot", "size": 14},
        "estate_merah": {"color": COLOR["estate_merah"], "shape": "diamond", "size": 16},
        "estate_kuning": {"color": COLOR["estate_kuning"], "shape": "diamond", "size": 14},
        "estate_hijau": {"color": COLOR["estate_hijau"], "shape": "diamond", "size": 12},
        "estate": {"color": COLOR["estate"], "shape": "diamond", "size": 12},
        "polres_kritis": {"color": COLOR["polres_kritis"], "shape": "box", "size": 28},
        "polres_tinggi": {"color": COLOR["polres_tinggi"], "shape": "box", "size": 26},
        "polres_sedang": {"color": COLOR["polres_sedang"], "shape": "box", "size": 22},
        "polres_rendah": {"color": COLOR["polres_rendah"], "shape": "box", "size": 18},
    }

    idmap = {n: i for i, n in enumerate(G.nodes())}
    for n, d in G.nodes(data=True):
        nt = d.get("ntype", "estate")
        style = group_map.get(nt, group_map["estate"])
        title_bits = [n, f"tipe: {nt}", f"degree: {G.degree(n)}"]
        if d.get("polres"):
            title_bits.append(f"polres: {d['polres']}")
        if d.get("klaster"):
            title_bits.append(f"klaster: {d['klaster']}")
        if d.get("band"):
            title_bits.append(f"band: {d['band']}")
        if d.get("hub_risiko"):
            title_bits.append("HUB RISIKO")
        nodes.append(
            {
                "id": idmap[n],
                "label": d.get("label", short(n, 22)).replace("\n", " "),
                "title": " | ".join(title_bits),
                "color": style["color"],
                "shape": style["shape"],
                "size": style["size"] + min(G.degree(n), 8),
                "font": {"color": "#1E242A", "size": 12},
            }
        )

    for u, v, d in G.edges(data=True):
        layer = d.get("layer", "")
        dashes = layer in ("estate-polres", "aktor-polres")
        width = 2.2 if layer == "agrinas-aktor" else 1.4
        edges.append(
            {
                "from": idmap[u],
                "to": idmap[v],
                "title": f"{d.get('relasi','')} ({layer})",
                "color": {"color": COLOR["edge_kelola"] if layer == "agrinas-aktor" else COLOR["edge_kso"]},
                "dashes": dashes,
                "width": width,
            }
        )

    hub_list = [
        n for n, d in G.nodes(data=True)
        if d.get("hub_risiko") and d.get("ntype") != "agrinas"
    ]
    hub_list.sort(key=lambda n: -G.degree(n))

    html = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8"/>
<title>Peta Jaringan KSO Agrinas — Unit II Harda Polda Riau</title>
<script src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
<style>
  body {{ margin:0; font-family: Calibri, Segoe UI, sans-serif; background:#0F2A44; color:#F4F6F8; }}
  header {{ padding:16px 24px; border-bottom:3px solid #C45C26; }}
  h1 {{ margin:0 0 4px; font-size:22px; }}
  .sub {{ color:#B8C5D0; font-size:13px; }}
  #layout {{ display:flex; height: calc(100vh - 88px); }}
  #mynetwork {{ flex:1; background:#F4F6F8; }}
  aside {{ width:320px; padding:16px; overflow:auto; background:#1A3A5C; }}
  aside h2 {{ font-size:14px; color:#C45C26; margin:12px 0 8px; }}
  .hub {{ background:#0F2A44; border-left:4px solid #B32D2D; padding:8px 10px; margin:6px 0; font-size:12px; }}
  .leg span {{ display:inline-block; width:12px; height:12px; margin-right:6px; border-radius:2px; }}
  .leg div {{ margin:4px 0; font-size:12px; }}
</style>
</head>
<body>
<header>
  <h1>Peta Jaringan KSO Agrinas</h1>
  <div class="sub">Unit II Harda · Ditreskrimum Polda Riau · Analisis berbasis R2 edges / matriks kebun · Agustus 2026</div>
</header>
<div id="layout">
  <div id="mynetwork"></div>
  <aside>
    <h2>Cara baca</h2>
    <div class="leg">
      <div><span style="background:{COLOR['agrinas']}"></span>Agrinas (pusat skema)</div>
      <div><span style="background:{COLOR['aktor_pam']}"></span>Hub PAM/KSO risiko</div>
      <div><span style="background:{COLOR['aktor_multi']}"></span>Hub multi-lokasi</div>
      <div><span style="background:{COLOR['aktor']}"></span>Penerima KSO</div>
      <div><span style="background:{COLOR['estate_merah']}"></span>Estate merah / kuning / hijau</div>
      <div><span style="background:{COLOR['polres_kritis']}"></span>Polres (kotak, warna = band risiko)</div>
    </div>
    <h2>Hub risiko prioritas</h2>
    {''.join(f'<div class="hub"><b>{short(h,40)}</b><br/>degree {G.degree(h)} · {G.nodes[h].get("ntype")}</div>' for h in hub_list[:12])}
    <h2>Statistik</h2>
    <div class="sub">Node: {G.number_of_nodes()} · Edge: {G.number_of_edges()} · Hub risiko: {sum(1 for _,d in G.nodes(data=True) if d.get('hub_risiko'))}</div>
    <p style="font-size:11px;color:#9AA8B5;margin-top:16px">Geser/zoom untuk eksplorasi. Hover node untuk detail. Garis putus = wilayah Polres.</p>
  </aside>
</div>
<script>
const nodes = new vis.DataSet({json.dumps(nodes, ensure_ascii=False)});
const edges = new vis.DataSet({json.dumps(edges, ensure_ascii=False)});
const container = document.getElementById('mynetwork');
const data = {{ nodes, edges }};
const options = {{
  physics: {{
    barnesHut: {{ gravitationalConstant: -12000, springLength: 120, springConstant: 0.03 }},
    stabilization: {{ iterations: 180 }}
  }},
  interaction: {{ hover: true, tooltipDelay: 80, navigationButtons: true, keyboard: true }},
  edges: {{ smooth: {{ type: 'continuous' }} }},
}};
new vis.Network(container, data, options);
</script>
</body>
</html>
"""
    out_html.write_text(html, encoding="utf-8")
    print(f"OK HTML: {out_html}")


def main():
    G, hubs = build_graph()
    print(f"Graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")
    print(f"Hubs (YA): {sum(1 for _, d in G.nodes(data=True) if d.get('hub_risiko'))}")

    draw_network(
        G,
        BASE / "Peta_Jaringan_KSO_Agrinas.png",
        "Peta Jaringan KSO Agrinas — Aktor · Estate · Polres\n(Unit II Harda Ditreskrimum Polda Riau)",
        focus_hubs=False,
    )
    draw_network(
        G,
        BASE / "Peta_Jaringan_KSO_Agrinas_hub.png",
        "Fokus Hub Risiko KSO/PAM — Multi-lokasi & Bentrok\n(Unit II Harda Ditreskrimum Polda Riau)",
        focus_hubs=True,
    )
    export_tables(G)
    export_graph_formats(G)
    write_html(G, BASE / "Peta_Jaringan_KSO_Agrinas.html")

    # Print top hubs for console
    print("\n=== TOP HUB (by degree) ===")
    ranked = sorted(
        [(n, G.degree(n), d.get("ntype")) for n, d in G.nodes(data=True) if d.get("ntype", "").startswith("aktor") or d.get("ntype") == "agrinas"],
        key=lambda x: -x[1],
    )
    for n, deg, nt in ranked[:15]:
        print(f"  {deg:2d}  [{nt}] {n}")


if __name__ == "__main__":
    main()
