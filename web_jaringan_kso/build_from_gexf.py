# -*- coding: utf-8 -*-
"""Konversi jaringan_kso_agrinas.gexf → graph-data.json untuk webpage interaktif."""

from __future__ import annotations

import json
from pathlib import Path

import networkx as nx

BASE = Path(r"C:\Users\Patron\Downloads\sawit lagi")
GEXF = BASE / "jaringan_kso_agrinas.gexf"
OUT_DIR = BASE / "web_jaringan_kso"
OUT_JSON = OUT_DIR / "graph-data.json"

COLORS = {
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

SHAPES = {
    "agrinas": "dot",
    "aktor_pam": "dot",
    "aktor_multi": "dot",
    "aktor": "dot",
    "aktor_generic": "dot",
    "estate_merah": "diamond",
    "estate_kuning": "diamond",
    "estate_hijau": "diamond",
    "estate": "diamond",
    "polres_kritis": "box",
    "polres_tinggi": "box",
    "polres_sedang": "box",
    "polres_rendah": "box",
}

GROUP_LABEL = {
    "agrinas": "Agrinas (pusat skema)",
    "aktor_pam": "Hub PAM/KSO risiko",
    "aktor_multi": "Hub multi-lokasi",
    "aktor": "Penerima KSO",
    "aktor_generic": "Label generik / gap data",
    "estate_merah": "Estate merah",
    "estate_kuning": "Estate kuning",
    "estate_hijau": "Estate hijau",
    "estate": "Estate (klaster n/a)",
    "polres_kritis": "Polres KRITIS",
    "polres_tinggi": "Polres TINGGI",
    "polres_sedang": "Polres SEDANG",
    "polres_rendah": "Polres RENDAH",
}


def ntype_bucket(ntype: str) -> str:
    if ntype.startswith("aktor") or ntype == "agrinas":
        return "aktor"
    if ntype.startswith("estate"):
        return "estate"
    if ntype.startswith("polres"):
        return "polres"
    return "lain"


def main():
    G = nx.read_gexf(GEXF)
    nodes = []
    edges = []
    idmap = {n: i for i, n in enumerate(G.nodes())}

    for n, d in G.nodes(data=True):
        ntype = d.get("ntype") or "estate"
        hub = str(d.get("hub_risiko", "")).lower() in ("true", "1", "ya")
        label = (d.get("label") or n).replace("\n", " ").strip()
        degree = G.degree(n)
        size_base = {
            "agrinas": 38,
            "aktor_pam": 26,
            "aktor_multi": 24,
            "aktor": 18,
            "aktor_generic": 15,
            "estate_merah": 18,
            "estate_kuning": 15,
            "estate_hijau": 13,
            "estate": 12,
            "polres_kritis": 30,
            "polres_tinggi": 26,
            "polres_sedang": 22,
            "polres_rendah": 18,
        }.get(ntype, 14)

        nodes.append(
            {
                "id": idmap[n],
                "key": n,
                "label": label,
                "ntype": ntype,
                "bucket": ntype_bucket(ntype),
                "groupLabel": GROUP_LABEL.get(ntype, ntype),
                "hub_risiko": hub,
                "degree": degree,
                "polres": d.get("polres") or "",
                "klaster": d.get("klaster") or "",
                "band": d.get("band") or "",
                "peran": d.get("peran") or "",
                "jml_polres": d.get("jml_polres") or "",
                "jml_estate": d.get("jml_estate") or "",
                "color": COLORS.get(ntype, "#5A6672"),
                "shape": SHAPES.get(ntype, "dot"),
                "size": size_base + min(degree, 10),
            }
        )

    for u, v, d in G.edges(data=True):
        layer = d.get("layer") or ""
        relasi = d.get("relasi") or ""
        edges.append(
            {
                "id": f"{idmap[u]}-{idmap[v]}",
                "from": idmap[u],
                "to": idmap[v],
                "fromKey": u,
                "toKey": v,
                "relasi": relasi,
                "layer": layer,
                "dashes": layer in ("estate-polres", "aktor-polres"),
                "width": 2.4 if layer == "agrinas-aktor" else 1.4,
            }
        )

    payload = {
        "meta": {
            "title": "Peta Jaringan KSO Agrinas",
            "unit": "Unit II Harda · Ditreskrimum Polda Riau",
            "source": "jaringan_kso_agrinas.gexf",
            "generated": "2026-08-04",
            "nodes": len(nodes),
            "edges": len(edges),
            "hubs": sum(1 for n in nodes if n["hub_risiko"]),
        },
        "legend": [{"ntype": k, "label": v, "color": COLORS[k]} for k, v in GROUP_LABEL.items() if k in COLORS],
        "nodes": nodes,
        "edges": edges,
    }

    OUT_DIR.mkdir(exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    OUT_JSON.write_text(text, encoding="utf-8")
    # JS embed agar bisa dibuka tanpa server (file://)
    js_path = OUT_DIR / "graph-data.js"
    js_path.write_text("window.GRAPH_DATA = " + text + ";\n", encoding="utf-8")
    # salin GEXF sumber ke folder web untuk referensi
    import shutil

    shutil.copy2(GEXF, OUT_DIR / "jaringan_kso_agrinas.gexf")
    print(f"OK: {OUT_JSON} ({len(nodes)} nodes, {len(edges)} edges)")
    print(f"OK: {js_path}")
    print(f"OK: {OUT_DIR / 'jaringan_kso_agrinas.gexf'}")


if __name__ == "__main__":
    main()
