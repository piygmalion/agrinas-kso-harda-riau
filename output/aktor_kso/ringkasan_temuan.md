# Analisis Aktor Jaringan KSO Agrinas
**Unit II Harda — Ditreskrimum Polda Riau** · 4 Agustus 2026

## Ringkasan metodologi
- Graf sumber: `jaringan_kso_agrinas.gexf` (aktor–estate–Polres).
- Metrik: degree, betweenness, closeness, eigenvector.
- Skor hub Harda = 25% degree + 25% betweenness + 20% span(multi) + 15% PAM + 10% estate merah + 5% koridor utara.
- Normalisasi tanpa Agrinas pusat; Agrinas berlabel REFERENSI.
- Band: KRITIS ≥70 · TINGGI ≥50 · SEDANG ≥30 · RENDAH <30.

## Temuan utama
1. **Agrinas** adalah pusat struktural skema (degree tertinggi) — dipakai sebagai referensi, bukan target 'oknum'.
2. Risiko analitik berada pada **aktor penerima KSO/PAM** yang menjembatani banyak kebun atau terkait bentrok.
3. Koridor utara (Rohul–Rohil–Bengkalis–Dumai) memusatkan hub PAM berbendera bentrok (Majuma, PAB, UTS, Riden Jaya).
4. Hub multi-lokasi (Bernas Mulya Mandiri, Berlian NP, Maju Serempak, PT Runggu) memperbesar peluang penolakan berulang lintas estate.
5. Label generik `KSO` / `Belum ada KSO` di graf menandai **gap data satker**, bukan entitas hukum — jangan dihitung sebagai hub operasional.

## Prioritas pantau aktor
| Rank | Aktor | Skor | Struktural | Boost dampak | Band | Polres | PAM |
|---:|---|---:|---:|---:|---|---|---|
| 1 | Bernas Mulya Mandiri | 75.0 | 75.0 | 0.0 | KRITIS | Inhu, Rohul | TIDAK |
| 2 | PT Nusantara Sawit Majuma | 74.1 | 39.1 | 35.0 | KRITIS | Rohul | YA |
| 3 | PT Ujung Tanjung Sejahtera | 73.6 | 48.6 | 25.0 | KRITIS | Rohil | YA |
| 4 | PT Palma Agung Betuah (PAB) | 70.6 | 42.6 | 28.0 | KRITIS | Bengkalis | YA |
| 5 | PT Riden Jaya Konstruksi | 51.5 | 33.5 | 18.0 | TINGGI | Dumai | YA |
| 6 | Berlian Nusantara Perkasa | 40.8 | 40.8 | 0.0 | SEDANG | Pelalawan | TIDAK |
| 7 | Poktan Riau Jaya Makmur | 40.5 | 40.5 | 0.0 | SEDANG | Kampar | YA |
| 8 | Poktan Berkah Tani Sejahtera | 35.7 | 35.7 | 0.0 | SEDANG | Inhil | TIDAK |

## Implikasi kerja Unit II Harda
1. Perbarui watchlist aktor bulanan dari `tabel_hub_risiko.csv`.
2. Cocokkan hub PAM dengan LP/bentrok (R3/R4) — terutama Rohul, Rohil, Bengkalis, Dumai, Kampar.
3. Minta satker melengkapi nama di balik label generik KSO (gap-fill R5).
4. Analisis ego-Polres: satker dengan `aktor_pam_flag` > 0 wajib masuk briefing mingguan.
5. Jangan agregasi TNTN Pelalawan ke skor KSO murni.

## File keluaran
- `tabel_aktor_metrics.csv`
- `tabel_hub_risiko.csv`
- `tabel_ego_polres.csv`
- `matriks_aktor_polres.csv`
- `fig_degree_hub.png`, `fig_betweenness.png`, `fig_network_overview.png`

*Dokumen analitik — bukan perintah operasi.*