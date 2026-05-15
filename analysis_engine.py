# ============================================================
# ANALISIS TERPADU POSKO EXISTING VS POSKO USULAN
# ============================================================
#
# Alur:
# 1. Hitung jumlah k ideal dengan Silhouette Score
# 2. Lakukan K-Means dengan jumlah k terbaik
# 3. Bentuk posko usulan dari centroid K-Means
# 4. Hitung jarak posko existing ke gangguan yang ditangani
# 5. Deteksi dan buang outlier jarak existing
# 6. Untuk data valid non-outlier, alokasikan gangguan ke posko usulan terdekat
# 7. Bandingkan parameter jarak dan estimasi response time
# 8. Evaluasi pemerataan beban kerja existing vs usulan

import re
import json
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler


# ------------------------------------------------------------
# PARAMETER
# ------------------------------------------------------------
FILE_GANGGUAN_VALID = "data_gangguan_valid_14052026 dari arcgis.xlsx"
SHEET_GANGGUAN_VALID = "data_gangguan_valid_14052026"

FILE_MASTER_BERSIH = "DATA GABUNGAN MINIMALIS - BERSIH NORMALISASI.xlsx"
SHEET_MASTER_BERSIH = "Data Bersih"

OUTPUT_EXCEL = "hasil_analisis_terpadu_posko.xlsx"
OUTPUT_GRAFIK_SILHOUETTE = "hasil_analisis_terpadu_silhouette.png"
OUTPUT_GRAFIK_BEBAN = "hasil_analisis_terpadu_beban_kerja.png"
OUTPUT_MAP_HTML = "hasil_kmeans_map.html"
OUTPUT_DIR_VISUAL_KMEANS = "visualisasi_kmeans_terpadu"

KOLOM_KODE_GANGGUAN = "NO_LAPORAN"
KOLOM_LATITUDE = "LAT_PLGN"
KOLOM_LONGITUDE = "LONG_PLGN"
KOLOM_REGU = "NAMA_REGU"
KOLOM_POSKO = "NAMA_POSKO"
KOLOM_RPT = "RPT"

TARGET_JUMLAH_TIM_EXISTING = 28
K_MIN = 15
K_MAX = TARGET_JUMLAH_TIM_EXISTING
RANDOM_STATE = 42
N_INIT = 10
SILHOUETTE_SAMPLE_SIZE = 10000

OUTLIER_IQR_MULTIPLIER = 1.5
BATAS_MIN_OUTLIER_EXISTING_KM = 15.0
KECEPATAN_KM_PER_JAM = 30
CHUNK_SIZE = 5000
MAX_TITIK_VISUALISASI = 100000

_PROGRESS_CALLBACK = None


# ------------------------------------------------------------
# POSKO EXISTING
# ------------------------------------------------------------
POSKO_EXISTING = [
    {"UNIT_POSKO": "ULP KOTA 51", "POSKO_LATITUDE": -6.721328, "POSKO_LONGITUDE": 108.572009},
    {"UNIT_POSKO": "ULP KOTA 52", "POSKO_LATITUDE": -6.721328, "POSKO_LONGITUDE": 108.572009},
    {"UNIT_POSKO": "ULP SUMBER 51", "POSKO_LATITUDE": -6.758961, "POSKO_LONGITUDE": 108.480259},
    {"UNIT_POSKO": "ULP SUMBER 52", "POSKO_LATITUDE": -6.758961, "POSKO_LONGITUDE": 108.480259},
    {"UNIT_POSKO": "ULP KUNINGAN 51", "POSKO_LATITUDE": -6.977561, "POSKO_LONGITUDE": 108.484565},
    {"UNIT_POSKO": "ULP KUNINGAN 52", "POSKO_LATITUDE": -6.977561, "POSKO_LONGITUDE": 108.484565},
    {"UNIT_POSKO": "ULP CILEDUG 51", "POSKO_LATITUDE": -6.907618, "POSKO_LONGITUDE": 108.749176},
    {"UNIT_POSKO": "ULP CILEDUG 52", "POSKO_LATITUDE": -6.907618, "POSKO_LONGITUDE": 108.749176},
    {"UNIT_POSKO": "ULP CILIMUS 51", "POSKO_LATITUDE": -6.877434, "POSKO_LONGITUDE": 108.496182},
    {"UNIT_POSKO": "ULP CILIMUS 52", "POSKO_LATITUDE": -6.877434, "POSKO_LONGITUDE": 108.496182},
    {"UNIT_POSKO": "MUNDU", "POSKO_LATITUDE": -6.751810915201957, "POSKO_LONGITUDE": 108.5902556329763},
    {"UNIT_POSKO": "KAPETAKAN", "POSKO_LATITUDE": -6.626170590077665, "POSKO_LONGITUDE": 108.52377980846067},
    {"UNIT_POSKO": "PLERED", "POSKO_LATITUDE": -6.704882, "POSKO_LONGITUDE": 108.506095},
    {"UNIT_POSKO": "ARJAWINANGUN", "POSKO_LATITUDE": -6.644440, "POSKO_LONGITUDE": 108.410443},
    {"UNIT_POSKO": "GEGESIK", "POSKO_LATITUDE": -6.593858, "POSKO_LONGITUDE": 108.426905},
    {"UNIT_POSKO": "PALIMANAN", "POSKO_LATITUDE": -6.709856, "POSKO_LONGITUDE": 108.439076},
    {"UNIT_POSKO": "KARANGSEMBUNG", "POSKO_LATITUDE": -6.854646, "POSKO_LONGITUDE": 108.643082},
    {"UNIT_POSKO": "LOSARI", "POSKO_LATITUDE": -6.840699, "POSKO_LONGITUDE": 108.802556},
    {"UNIT_POSKO": "SINDANGLAUT", "POSKO_LATITUDE": -6.830417, "POSKO_LONGITUDE": 108.622643},
    {"UNIT_POSKO": "PANGENAN", "POSKO_LATITUDE": -6.818578, "POSKO_LONGITUDE": 108.705260},
    {"UNIT_POSKO": "JALAKSANA", "POSKO_LATITUDE": -6.920232, "POSKO_LONGITUDE": 108.486959},
    {"UNIT_POSKO": "MANDIRANCAN", "POSKO_LATITUDE": -6.805028, "POSKO_LONGITUDE": 108.470044},
    {"UNIT_POSKO": "CINIRU", "POSKO_LATITUDE": -7.050376, "POSKO_LONGITUDE": 108.489688},
    {"UNIT_POSKO": "MALEBER", "POSKO_LATITUDE": -7.008398, "POSKO_LONGITUDE": 108.565704},
    {"UNIT_POSKO": "CIWIGEBANG", "POSKO_LATITUDE": -6.972867, "POSKO_LONGITUDE": 108.586270},
    {"UNIT_POSKO": "CIKIJING", "POSKO_LATITUDE": -7.020716, "POSKO_LONGITUDE": 108.391902},
    {"UNIT_POSKO": "LURAGUNG", "POSKO_LATITUDE": -7.037613, "POSKO_LONGITUDE": 108.698346},
    {"UNIT_POSKO": "CIBINGBIN", "POSKO_LATITUDE": -7.057459, "POSKO_LONGITUDE": 108.757218},
]

ALIAS_UNIT = {
    "KOTA": "ULP KOTA 51",
    "CIREBON KOTA": "ULP KOTA 51",
    "ULP CIREBON KOTA": "ULP KOTA 51",
    "POSKO ULP CIREBON KOTA": "ULP KOTA 51",
    "SUMBER": "ULP SUMBER 51",
    "ULP SUMBER": "ULP SUMBER 51",
    "POSKO ULP SUMBER": "ULP SUMBER 51",
    "KUNINGAN": "ULP KUNINGAN 51",
    "ULP KUNINGAN": "ULP KUNINGAN 51",
    "POSKO ULP KUNINGAN": "ULP KUNINGAN 51",
    "CILEDUG": "ULP CILEDUG 51",
    "ULP CILEDUG": "ULP CILEDUG 51",
    "POSKO ULP CILEDUG": "ULP CILEDUG 51",
    "CILIMUS": "ULP CILIMUS 51",
    "ULP CILIMUS": "ULP CILIMUS 51",
    "POSKO ULP CILIMUS": "ULP CILIMUS 51",
    "CIAWIGEBANG": "CIWIGEBANG",
}


def log_progress(tahap, total_tahap, pesan):
    persen = tahap / total_tahap * 100
    if _PROGRESS_CALLBACK is not None:
        _PROGRESS_CALLBACK(tahap, total_tahap, pesan)
    print(f"[{persen:6.2f}%] {pesan}", flush=True)


def log_sub_progress(nomor, total, pesan, interval=None):
    if total == 0:
        return
    if interval is None:
        interval = max(1, total // 20)
    if nomor == 1 or nomor == total or nomor % interval == 0:
        persen = nomor / total * 100
        print(f"         [{persen:6.2f}%] {pesan}", flush=True)


def normalisasi_teks(nilai):
    if pd.isna(nilai):
        return ""
    return re.sub(r"\s+", " ", str(nilai).upper().strip())


def unit_dari_regu(nama_regu):
    teks = normalisasi_teks(nama_regu)
    if not teks:
        return None

    match = re.match(r"^(?P<unit>.+?)\s+(?P<kode>\d+(?:\.\d+)?)$", teks)
    if match:
        unit_dasar = match.group("unit").strip()
        kode = match.group("kode").replace(".", "")
        if unit_dasar in {"KOTA", "SUMBER", "KUNINGAN", "CILEDUG", "CILIMUS"} and kode in {"51", "52"}:
            return f"ULP {unit_dasar} {kode}"
        return ALIAS_UNIT.get(unit_dasar, unit_dasar)

    return ALIAS_UNIT.get(teks, teks)


def unit_dari_posko(nama_posko):
    teks = normalisasi_teks(nama_posko)
    if not teks:
        return None
    return ALIAS_UNIT.get(teks, None)


def tentukan_unit_posko(row):
    unit = unit_dari_regu(row.get(f"{KOLOM_REGU}_MASTER"))
    if unit:
        return unit

    unit = unit_dari_posko(row.get(f"{KOLOM_POSKO}_MASTER"))
    if unit:
        return unit

    unit = unit_dari_regu(row.get(KOLOM_REGU))
    if unit:
        return unit

    return unit_dari_posko(row.get(KOLOM_POSKO))


def haversine_km(lat1, lon1, lat2, lon2):
    radius_bumi_km = 6371.0088
    lat1 = np.radians(lat1)
    lon1 = np.radians(lon1)
    lat2 = np.radians(lat2)
    lon2 = np.radians(lon2)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    c = 2 * np.arcsin(np.sqrt(a))
    return radius_bumi_km * c


def hitung_k_ideal(x_scaled):
    hasil = []
    k_values = list(range(K_MIN, K_MAX + 1))

    for nomor, k in enumerate(k_values, start=1):
        log_sub_progress(nomor, len(k_values), f"Uji Silhouette k={k}", interval=1)
        kmeans = KMeans(n_clusters=k, random_state=RANDOM_STATE, n_init=N_INIT)
        labels = kmeans.fit_predict(x_scaled)

        sample_size = min(SILHOUETTE_SAMPLE_SIZE, len(x_scaled))
        score = silhouette_score(
            x_scaled,
            labels,
            sample_size=sample_size,
            random_state=RANDOM_STATE,
        )
        hasil.append({"Jumlah Cluster": k, "Silhouette Score": score})

    result_df = pd.DataFrame(hasil)
    best_row = result_df.loc[result_df["Silhouette Score"].idxmax()]
    return result_df, int(best_row["Jumlah Cluster"]), float(best_row["Silhouette Score"])


def buat_grafik_silhouette(result_df):
    plt.figure(figsize=(8, 5))
    plt.plot(result_df["Jumlah Cluster"], result_df["Silhouette Score"], marker="o")
    plt.xlabel("Jumlah Cluster (k)")
    plt.ylabel("Silhouette Score")
    plt.title("Perbandingan Silhouette Score")
    plt.grid(True)
    plt.tight_layout()
    plt.savefig(OUTPUT_GRAFIK_SILHOUETTE, dpi=200)
    plt.close()


def sampling_visualisasi(data_plot):
    if len(data_plot) <= MAX_TITIK_VISUALISASI:
        return data_plot
    return data_plot.sample(MAX_TITIK_VISUALISASI, random_state=RANDOM_STATE)


def simpan_visualisasi_kmeans(data_plot, label_col, judul, nama_file, catatan=None):
    output_dir = Path(OUTPUT_DIR_VISUAL_KMEANS)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / nama_file

    data_visual = sampling_visualisasi(data_plot[[KOLOM_LATITUDE, KOLOM_LONGITUDE, label_col]].dropna().copy())
    centroid = (
        data_plot
        .groupby(label_col, as_index=False)
        .agg(
            CENTROID_LATITUDE=(KOLOM_LATITUDE, "mean"),
            CENTROID_LONGITUDE=(KOLOM_LONGITUDE, "mean"),
            JUMLAH_TITIK=(label_col, "size"),
        )
        .sort_values(label_col)
    )

    plt.figure(figsize=(11, 8))
    scatter = plt.scatter(
        data_visual[KOLOM_LONGITUDE],
        data_visual[KOLOM_LATITUDE],
        c=data_visual[label_col],
        cmap="tab20",
        s=7,
        alpha=0.55,
        linewidths=0,
    )
    plt.scatter(
        centroid["CENTROID_LONGITUDE"],
        centroid["CENTROID_LATITUDE"],
        marker="x",
        c="black",
        s=90,
        linewidths=2,
        label="Centroid",
    )

    for _, row in centroid.iterrows():
        plt.annotate(
            str(int(row[label_col])),
            (row["CENTROID_LONGITUDE"], row["CENTROID_LATITUDE"]),
            xytext=(4, 4),
            textcoords="offset points",
            fontsize=7,
            color="black",
        )

    plt.xlabel("Longitude")
    plt.ylabel("Latitude")
    plt.title(judul)
    if catatan:
        plt.figtext(0.01, 0.01, catatan, fontsize=8)
    plt.grid(True, alpha=0.25)
    plt.legend(loc="best")
    plt.colorbar(scatter, label="Cluster")
    plt.tight_layout()
    plt.savefig(output_path, dpi=200)
    plt.close()
    return str(output_path)


def buat_peta_cluster_html(data_plot, posko_usulan, nama_file):
    titik_cluster = (
        data_plot[[KOLOM_LATITUDE, KOLOM_LONGITUDE, "USULAN_ID"]]
        .dropna()
        .copy()
    )
    titik_cluster["USULAN_ID"] = titik_cluster["USULAN_ID"].astype(int)
    points = [
        [float(row[KOLOM_LATITUDE]), float(row[KOLOM_LONGITUDE]), int(row["USULAN_ID"])]
        for row in titik_cluster.to_dict(orient="records")
    ]

    centroid = posko_usulan[["USULAN_LATITUDE", "USULAN_LONGITUDE", "USULAN_ID", "NAMA_POSKO_USULAN"]].dropna().copy()
    centroid["USULAN_ID"] = centroid["USULAN_ID"].astype(int)
    centroids = [
        [
            float(row["USULAN_LATITUDE"]),
            float(row["USULAN_LONGITUDE"]),
            int(row["USULAN_ID"]),
            str(row["NAMA_POSKO_USULAN"]),
        ]
        for row in centroid.to_dict(orient="records")
    ]

    counts_df = (
        titik_cluster["USULAN_ID"]
        .value_counts()
        .sort_index()
        .rename_axis("USULAN_ID")
        .reset_index(name="JUMLAH_TITIK_DILAYANI")
    )
    point_counts = {
        int(row["USULAN_ID"]): int(row["JUMLAH_TITIK_DILAYANI"])
        for row in counts_df.to_dict(orient="records")
    }

    center_lat = float(titik_cluster[KOLOM_LATITUDE].mean()) if len(titik_cluster) else -6.8
    center_lng = float(titik_cluster[KOLOM_LONGITUDE].mean()) if len(titik_cluster) else 108.5

    html = """
<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Peta Posko Usulan K-Means</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
  <style>
    html, body, #map {
      height: 100%;
      margin: 0;
      font-family: Arial, sans-serif;
    }
    .panel {
      position: absolute;
      top: 12px;
      right: 12px;
      z-index: 1000;
      max-width: 300px;
      padding: 12px 14px;
      background: rgba(255, 255, 255, 0.94);
      border: 1px solid #d0d7de;
      border-radius: 6px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.16);
      font-size: 13px;
      line-height: 1.4;
    }
    .panel h1 {
      margin: 0 0 8px;
      font-size: 16px;
      line-height: 1.2;
    }
    .legend {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 4px 10px;
      margin-top: 8px;
      max-height: 260px;
      overflow: auto;
    }
    .legend-item {
      display: flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
    }
    .swatch {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      border: 1px solid rgba(0, 0, 0, 0.2);
      flex: 0 0 auto;
    }
  </style>
</head>
<body>
  <div id="map"></div>
  <div class="panel">
    <h1>Peta Posko Usulan</h1>
    <div><b>Jumlah posko usulan:</b> __JUMLAH_POSKO__</div>
    <div><b>Jumlah titik:</b> __JUMLAH_TITIK__</div>
    <div><b>Posko usulan:</b> tanda X hitam</div>
    <div style="margin-top:8px;"><b>Titik dilayani per posko:</b></div>
    <div id="legend" class="legend"></div>
  </div>

  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
    const points = __POINTS_JSON__;
    const centroids = __CENTROIDS_JSON__;
    const pointCounts = __POINT_COUNTS_JSON__;
    const colors = [
      "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
      "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
      "#005f73", "#bb3e03", "#0a9396", "#ae2012", "#6a4c93",
      "#1982c4", "#8ac926", "#ff595e", "#ffca3a", "#4267ac",
      "#7b2cbf", "#2d6a4f", "#c1121f", "#f77f00", "#6c757d",
      "#3a86ff", "#8338ec", "#fb5607", "#06d6a0", "#118ab2"
    ];

    const map = L.map("map", { preferCanvas: true }).setView([__CENTER_LAT__, __CENTER_LNG__], 11);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: "Map data: OpenStreetMap contributors"
    }).addTo(map);

    const canvasLayer = L.Layer.extend({
      onAdd: function(map) {
        this._map = map;
        this._canvas = L.DomUtil.create("canvas", "leaflet-zoom-animated");
        this._ctx = this._canvas.getContext("2d");
        map.getPanes().overlayPane.appendChild(this._canvas);
        map.on("move zoom resize", this._reset, this);
        this._reset();
      },
      onRemove: function(map) {
        map.getPanes().overlayPane.removeChild(this._canvas);
        map.off("move zoom resize", this._reset, this);
      },
      _reset: function() {
        const size = this._map.getSize();
        const topLeft = this._map.containerPointToLayerPoint([0, 0]);
        L.DomUtil.setPosition(this._canvas, topLeft);
        this._canvas.width = size.x;
        this._canvas.height = size.y;
        this._draw();
      },
      _draw: function() {
        const ctx = this._ctx;
        const size = this._map.getSize();
        ctx.clearRect(0, 0, size.x, size.y);

        for (const item of points) {
          const p = this._map.latLngToContainerPoint([item[0], item[1]]);
          if (p.x < -8 || p.y < -8 || p.x > size.x + 8 || p.y > size.y + 8) {
            continue;
          }
          ctx.beginPath();
          ctx.fillStyle = colors[item[2] % colors.length];
          ctx.globalAlpha = 0.68;
          ctx.arc(p.x, p.y, 3, 0, Math.PI * 2);
          ctx.fill();
        }
        ctx.globalAlpha = 1;
      }
    });

    map.addLayer(new canvasLayer());

    const bounds = [];
    for (const item of points) {
      bounds.push([item[0], item[1]]);
    }

    for (const item of centroids) {
      const icon = L.divIcon({
        className: "",
        html: `<div style="font-size:26px;font-weight:700;color:#111;text-shadow:0 0 3px #fff;">X</div>`,
        iconSize: [26, 26],
        iconAnchor: [13, 13]
      });
      const count = pointCounts[item[2]] || 0;
      L.marker([item[0], item[1]], { icon })
        .bindPopup(`${item[3]}<br>Titik dilayani: ${count.toLocaleString("id-ID")}<br>Lat: ${item[0].toFixed(6)}<br>Long: ${item[1].toFixed(6)}`)
        .addTo(map);
      bounds.push([item[0], item[1]]);
    }

    if (bounds.length > 0) {
      map.fitBounds(bounds, { padding: [24, 24] });
    }

    const clusterIds = [...new Set(points.map(item => item[2]))].sort((a, b) => a - b);
    const legend = document.getElementById("legend");
    for (const id of clusterIds) {
      const div = document.createElement("div");
      div.className = "legend-item";
      div.innerHTML = `<span class="swatch" style="background:${colors[id % colors.length]}"></span><span>USULAN_${id}: ${(pointCounts[id] || 0).toLocaleString("id-ID")}</span>`;
      legend.appendChild(div);
    }
  </script>
</body>
</html>
"""

    html = html.replace("__JUMLAH_POSKO__", f"{len(centroids):,}")
    html = html.replace("__JUMLAH_TITIK__", f"{len(points):,}")
    html = html.replace("__POINTS_JSON__", json.dumps(points, separators=(",", ":")))
    html = html.replace("__CENTROIDS_JSON__", json.dumps(centroids, separators=(",", ":")))
    html = html.replace("__POINT_COUNTS_JSON__", json.dumps(point_counts, separators=(",", ":")))
    html = html.replace("__CENTER_LAT__", str(center_lat))
    html = html.replace("__CENTER_LNG__", str(center_lng))

    path_output = Path(nama_file)
    path_output.write_text(html, encoding="utf-8")
    return path_output


def bentuk_posko_usulan(data_cluster):
    posko_usulan = (
        data_cluster
        .groupby("CLUSTER_USULAN", as_index=False)
        .agg(
            USULAN_LATITUDE=(KOLOM_LATITUDE, "mean"),
            USULAN_LONGITUDE=(KOLOM_LONGITUDE, "mean"),
            JUMLAH_GANGGUAN_AWAL=("CLUSTER_USULAN", "size"),
        )
        .rename(columns={"CLUSTER_USULAN": "USULAN_ID"})
        .sort_values("USULAN_ID")
        .reset_index(drop=True)
    )
    posko_usulan["NAMA_POSKO_USULAN"] = "USULAN_" + posko_usulan["USULAN_ID"].astype(str)
    return posko_usulan


def bentuk_posko_usulan_awal(data_cluster):
    posko_usulan_awal = (
        data_cluster
        .groupby("CLUSTER_KMEANS_AWAL", as_index=False)
        .agg(
            USULAN_AWAL_LATITUDE=(KOLOM_LATITUDE, "mean"),
            USULAN_AWAL_LONGITUDE=(KOLOM_LONGITUDE, "mean"),
            JUMLAH_GANGGUAN_AWAL=("CLUSTER_KMEANS_AWAL", "size"),
        )
        .rename(columns={"CLUSTER_KMEANS_AWAL": "USULAN_AWAL_ID"})
        .sort_values("USULAN_AWAL_ID")
        .reset_index(drop=True)
    )
    posko_usulan_awal["NAMA_POSKO_USULAN_AWAL"] = (
        "USULAN_AWAL_" + posko_usulan_awal["USULAN_AWAL_ID"].astype(str)
    )
    return posko_usulan_awal


def split_cluster_terbesar(data_cluster, target_jumlah_cluster):
    hasil = data_cluster.copy()
    riwayat = []
    visualisasi = []
    iterasi = 1

    while hasil["CLUSTER_USULAN"].nunique() < target_jumlah_cluster:
        jumlah_per_cluster = hasil["CLUSTER_USULAN"].value_counts().sort_values(ascending=False)
        cluster_sumber = None
        jumlah_awal = 0
        idx_cluster = None

        for kandidat_cluster, kandidat_jumlah in jumlah_per_cluster.items():
            kandidat_idx = hasil.index[hasil["CLUSTER_USULAN"] == kandidat_cluster]
            kandidat_xy = hasil.loc[kandidat_idx, [KOLOM_LATITUDE, KOLOM_LONGITUDE]]
            if int(kandidat_jumlah) >= 2 and kandidat_xy.drop_duplicates().shape[0] >= 2:
                cluster_sumber = int(kandidat_cluster)
                jumlah_awal = int(kandidat_jumlah)
                idx_cluster = kandidat_idx
                break

        if cluster_sumber is None or idx_cluster is None:
            print("         WARNING: split dihentikan karena tidak ada cluster yang bisa dibelah.", flush=True)
            break

        if jumlah_awal < 2:
            print("         WARNING: split dihentikan karena cluster terbesar hanya berisi 1 titik.", flush=True)
            break

        x_cluster = hasil.loc[idx_cluster, [KOLOM_LATITUDE, KOLOM_LONGITUDE]].copy()
        scaler_lokal = StandardScaler()
        x_cluster_scaled = scaler_lokal.fit_transform(x_cluster)

        kmeans_lokal = KMeans(n_clusters=2, random_state=RANDOM_STATE, n_init=N_INIT)
        label_lokal = kmeans_lokal.fit_predict(x_cluster_scaled)
        if len(np.unique(label_lokal)) < 2:
            print(f"         WARNING: cluster {cluster_sumber} tidak bisa dibelah menjadi 2 label.", flush=True)
            break

        data_split_visual = hasil.loc[idx_cluster, [KOLOM_LATITUDE, KOLOM_LONGITUDE]].copy()
        data_split_visual["CLUSTER_SPLIT_LOKAL"] = label_lokal
        file_visual_split = simpan_visualisasi_kmeans(
            data_split_visual,
            "CLUSTER_SPLIT_LOKAL",
            f"Split {iterasi} - Cluster {cluster_sumber} menjadi 2 subcluster",
            f"kmeans_split_{iterasi:02d}_cluster_{cluster_sumber}.png",
            catatan=f"Cluster sumber berisi {jumlah_awal:,} titik. Visualisasi ini hanya menampilkan titik cluster yang dibelah.",
        )

        label_baru = int(hasil["CLUSTER_USULAN"].max()) + 1
        idx_label_baru = idx_cluster[label_lokal == 1]
        if len(idx_label_baru) == 0 or len(idx_label_baru) == len(idx_cluster):
            print(f"         WARNING: cluster {cluster_sumber} menghasilkan split kosong.", flush=True)
            break
        hasil.loc[idx_label_baru, "CLUSTER_USULAN"] = label_baru

        jumlah_tetap = int((label_lokal == 0).sum())
        jumlah_baru = int((label_lokal == 1).sum())
        riwayat.append({
            "ITERASI_SPLIT": iterasi,
            "CLUSTER_DIBELAH": cluster_sumber,
            "CLUSTER_BARU": label_baru,
            "JUMLAH_SEBELUM_SPLIT": jumlah_awal,
            "JUMLAH_TETAP_DI_CLUSTER_LAMA": jumlah_tetap,
            "JUMLAH_CLUSTER_BARU": jumlah_baru,
            "JUMLAH_CLUSTER_SETELAH_SPLIT": hasil["CLUSTER_USULAN"].nunique(),
            "FILE_VISUALISASI_SPLIT": file_visual_split,
        })
        visualisasi.append({
            "TAHAP": f"SPLIT_{iterasi:02d}",
            "KETERANGAN": f"KMeans k=2 untuk membelah cluster {cluster_sumber}",
            "FILE_VISUALISASI": file_visual_split,
        })

        print(
            f"         Split {iterasi}: cluster {cluster_sumber} ({jumlah_awal:,} titik) "
            f"-> {jumlah_tetap:,} + {jumlah_baru:,} titik",
            flush=True,
        )
        iterasi += 1

    urutan_cluster = sorted(hasil["CLUSTER_USULAN"].unique())
    mapping_cluster = {cluster_lama: cluster_baru for cluster_baru, cluster_lama in enumerate(urutan_cluster)}
    hasil["CLUSTER_USULAN_SEBELUM_RENUMBER"] = hasil["CLUSTER_USULAN"]
    hasil["CLUSTER_USULAN"] = hasil["CLUSTER_USULAN"].map(mapping_cluster).astype(int)

    riwayat_df = pd.DataFrame(riwayat)
    if not riwayat_df.empty:
        riwayat_df["CLUSTER_DIBELAH_FINAL"] = riwayat_df["CLUSTER_DIBELAH"].map(mapping_cluster)
        riwayat_df["CLUSTER_BARU_FINAL"] = riwayat_df["CLUSTER_BARU"].map(mapping_cluster)

    return hasil, riwayat_df, visualisasi


def jalankan_kmeans_usulan(data, x_scaled, best_k, target_jumlah_cluster):
    kmeans = KMeans(n_clusters=best_k, random_state=RANDOM_STATE, n_init=N_INIT)
    labels = kmeans.fit_predict(x_scaled)

    data_cluster = data.copy()
    data_cluster["CLUSTER_KMEANS_AWAL"] = labels
    data_cluster["CLUSTER_USULAN"] = labels
    visualisasi = []
    posko_usulan_awal = bentuk_posko_usulan_awal(data_cluster)

    file_visual_awal = simpan_visualisasi_kmeans(
        data_cluster,
        "CLUSTER_KMEANS_AWAL",
        f"KMeans Awal Hasil Silhouette - k={best_k}",
        f"kmeans_awal_k_{best_k}.png",
        catatan="Cluster awal dari nilai k terbaik berdasarkan Silhouette Score.",
    )
    visualisasi.append({
        "TAHAP": "KMEANS_AWAL",
        "KETERANGAN": f"KMeans awal hasil silhouette dengan k={best_k}",
        "FILE_VISUALISASI": file_visual_awal,
    })

    if best_k < target_jumlah_cluster:
        jumlah_split = target_jumlah_cluster - best_k
        print(f"         K awal {best_k}, target {target_jumlah_cluster}; split cluster terbesar {jumlah_split} kali.", flush=True)
        data_cluster, riwayat_split, visualisasi_split = split_cluster_terbesar(data_cluster, target_jumlah_cluster)
        visualisasi.extend(visualisasi_split)
    else:
        print(f"         K awal sudah sama dengan target {target_jumlah_cluster}; tidak ada split tambahan.", flush=True)
        riwayat_split = pd.DataFrame(columns=[
            "ITERASI_SPLIT",
            "CLUSTER_DIBELAH",
            "CLUSTER_BARU",
            "JUMLAH_SEBELUM_SPLIT",
            "JUMLAH_TETAP_DI_CLUSTER_LAMA",
            "JUMLAH_CLUSTER_BARU",
            "JUMLAH_CLUSTER_SETELAH_SPLIT",
            "FILE_VISUALISASI_SPLIT",
            "CLUSTER_DIBELAH_FINAL",
            "CLUSTER_BARU_FINAL",
        ])

    posko_usulan = bentuk_posko_usulan(data_cluster)
    file_visual_final = simpan_visualisasi_kmeans(
        data_cluster,
        "CLUSTER_USULAN",
        f"KMeans Final Posko Usulan - k={len(posko_usulan)}",
        f"kmeans_final_k_{len(posko_usulan)}.png",
        catatan="Cluster final setelah split cluster terbesar sampai jumlah posko usulan sama dengan jumlah tim existing.",
    )
    visualisasi.append({
        "TAHAP": "KMEANS_FINAL",
        "KETERANGAN": f"Cluster final posko usulan dengan k={len(posko_usulan)}",
        "FILE_VISUALISASI": file_visual_final,
    })

    visualisasi_df = pd.DataFrame(visualisasi)
    return data_cluster, posko_usulan_awal, posko_usulan, riwayat_split, visualisasi_df


def hitung_jarak_existing(detail):
    return haversine_km(
        detail["POSKO_LATITUDE"].to_numpy(),
        detail["POSKO_LONGITUDE"].to_numpy(),
        detail[KOLOM_LATITUDE].to_numpy(),
        detail[KOLOM_LONGITUDE].to_numpy(),
    )


def alokasi_ke_posko_usulan(data_valid, posko_usulan):
    lat = data_valid[KOLOM_LATITUDE].to_numpy()
    lon = data_valid[KOLOM_LONGITUDE].to_numpy()
    pusat_lat = posko_usulan["USULAN_LATITUDE"].to_numpy()
    pusat_lon = posko_usulan["USULAN_LONGITUDE"].to_numpy()
    pusat_id = posko_usulan["USULAN_ID"].to_numpy()
    pusat_nama = posko_usulan["NAMA_POSKO_USULAN"].to_numpy()

    hasil_id = np.empty(len(data_valid), dtype=int)
    hasil_nama = np.empty(len(data_valid), dtype=object)
    hasil_lat = np.empty(len(data_valid), dtype=float)
    hasil_lon = np.empty(len(data_valid), dtype=float)
    hasil_jarak = np.empty(len(data_valid), dtype=float)

    total_chunk = int(np.ceil(len(data_valid) / CHUNK_SIZE))
    for nomor_chunk, start in enumerate(range(0, len(data_valid), CHUNK_SIZE), start=1):
        end = min(start + CHUNK_SIZE, len(data_valid))
        jarak_matrix = haversine_km(
            lat[start:end, None],
            lon[start:end, None],
            pusat_lat[None, :],
            pusat_lon[None, :],
        )
        idx_min = np.argmin(jarak_matrix, axis=1)
        row_pos = np.arange(end - start)

        hasil_id[start:end] = pusat_id[idx_min]
        hasil_nama[start:end] = pusat_nama[idx_min]
        hasil_lat[start:end] = pusat_lat[idx_min]
        hasil_lon[start:end] = pusat_lon[idx_min]
        hasil_jarak[start:end] = jarak_matrix[row_pos, idx_min]

        log_sub_progress(nomor_chunk, total_chunk, f"Alokasi chunk {nomor_chunk}/{total_chunk}", interval=1)

    hasil = data_valid.copy()
    hasil["USULAN_ID"] = hasil_id
    hasil["NAMA_POSKO_USULAN"] = hasil_nama
    hasil["USULAN_LATITUDE"] = hasil_lat
    hasil["USULAN_LONGITUDE"] = hasil_lon
    hasil["JARAK_POSKO_USULAN_KE_GANGGUAN_KM"] = hasil_jarak
    hasil["ESTIMASI_WAKTU_USULAN_MENIT"] = hasil_jarak / KECEPATAN_KM_PER_JAM * 60
    hasil["SELISIH_JARAK_KM"] = (
        hasil["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"]
        - hasil["JARAK_POSKO_USULAN_KE_GANGGUAN_KM"]
    )
    hasil["EFISIENSI_JARAK_PERSEN"] = np.where(
        hasil["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"] > 0,
        hasil["SELISIH_JARAK_KM"] / hasil["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"] * 100,
        np.nan,
    )
    hasil["STATUS_PERBANDINGAN"] = np.where(
        hasil["SELISIH_JARAK_KM"] > 0,
        "USULAN_LEBIH_DEKAT",
        np.where(hasil["SELISIH_JARAK_KM"] < 0, "EXISTING_LEBIH_DEKAT", "SAMA"),
    )
    return hasil


def tambah_estimasi_response_time(data_valid):
    hasil = data_valid.copy()
    hasil[KOLOM_RPT] = pd.to_numeric(hasil[KOLOM_RPT], errors="coerce")

    total_existing = hasil["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"].sum()
    total_usulan = hasil["JARAK_POSKO_USULAN_KE_GANGGUAN_KM"].sum()
    efisiensi_total = (total_existing - total_usulan) / total_existing if total_existing else np.nan

    hasil["EFISIENSI_TOTAL_JARAK_PERSEN_UNTUK_RPT"] = efisiensi_total * 100
    hasil["ESTIMASI_RPT_USULAN_MENIT"] = hasil[KOLOM_RPT] * (1 - efisiensi_total)
    hasil["SELISIH_RPT_ESTIMASI_MENIT"] = hasil[KOLOM_RPT] - hasil["ESTIMASI_RPT_USULAN_MENIT"]
    hasil["STATUS_RPT_EXISTING_GT_30"] = hasil[KOLOM_RPT] > 30
    hasil["STATUS_RPT_USULAN_GT_30"] = hasil["ESTIMASI_RPT_USULAN_MENIT"] > 30
    return hasil, efisiensi_total


def trace_posko_existing(detail_existing):
    awal = (
        detail_existing.groupby("UNIT_POSKO_EXISTING")
        .agg(
            JUMLAH_AWAL_SEBELUM_OUTLIER=(KOLOM_KODE_GANGGUAN, "count"),
            RATA_RATA_JARAK_AWAL_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "mean"),
            MEDIAN_JARAK_AWAL_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "median"),
            MAKS_JARAK_AWAL_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "max"),
        )
        .reset_index()
    )
    outlier = (
        detail_existing[detail_existing["IS_OUTLIER_EXISTING"]]
        .groupby("UNIT_POSKO_EXISTING")
        .agg(JUMLAH_DIBUANG_OUTLIER=(KOLOM_KODE_GANGGUAN, "count"))
        .reset_index()
    )
    akhir = (
        detail_existing[~detail_existing["IS_OUTLIER_EXISTING"]]
        .groupby("UNIT_POSKO_EXISTING")
        .agg(JUMLAH_AKHIR_VALID=(KOLOM_KODE_GANGGUAN, "count"))
        .reset_index()
    )

    trace = awal.merge(outlier, on="UNIT_POSKO_EXISTING", how="left")
    trace = trace.merge(akhir, on="UNIT_POSKO_EXISTING", how="left")
    trace["JUMLAH_DIBUANG_OUTLIER"] = trace["JUMLAH_DIBUANG_OUTLIER"].fillna(0).astype(int)
    trace["JUMLAH_AKHIR_VALID"] = trace["JUMLAH_AKHIR_VALID"].fillna(0).astype(int)
    trace["PERSEN_DIBUANG_OUTLIER"] = (
        trace["JUMLAH_DIBUANG_OUTLIER"] / trace["JUMLAH_AWAL_SEBELUM_OUTLIER"] * 100
    )
    return trace.sort_values("JUMLAH_DIBUANG_OUTLIER", ascending=False)


def ringkasan_global(data_valid, outlier_df, data_tidak_terpetakan):
    existing = data_valid["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"]
    usulan = data_valid["JARAK_POSKO_USULAN_KE_GANGGUAN_KM"]
    selisih = data_valid["SELISIH_JARAK_KM"]
    rpt_existing = data_valid[KOLOM_RPT].dropna()
    rpt_usulan = data_valid["ESTIMASI_RPT_USULAN_MENIT"].dropna()
    rpt_comparable = data_valid.dropna(subset=[KOLOM_RPT, "ESTIMASI_RPT_USULAN_MENIT"])

    return pd.DataFrame([
        {"METRIK": "Jumlah data awal valid ArcGIS", "NILAI": len(data_valid) + len(outlier_df) + len(data_tidak_terpetakan)},
        {"METRIK": "Jumlah data tidak terpetakan", "NILAI": len(data_tidak_terpetakan)},
        {"METRIK": "Jumlah outlier existing dibuang", "NILAI": len(outlier_df)},
        {"METRIK": "Jumlah data dibandingkan", "NILAI": len(data_valid)},
        {"METRIK": "Total jarak existing valid km", "NILAI": existing.sum()},
        {"METRIK": "Total jarak usulan valid km", "NILAI": usulan.sum()},
        {"METRIK": "Penghematan total km", "NILAI": selisih.sum()},
        {"METRIK": "Penghematan total persen", "NILAI": selisih.sum() / existing.sum() * 100 if existing.sum() else np.nan},
        {"METRIK": "Rata-rata jarak existing km", "NILAI": existing.mean()},
        {"METRIK": "Rata-rata jarak usulan km", "NILAI": usulan.mean()},
        {"METRIK": "Median jarak existing km", "NILAI": existing.median()},
        {"METRIK": "Median jarak usulan km", "NILAI": usulan.median()},
        {"METRIK": "Maks jarak existing valid km", "NILAI": existing.max()},
        {"METRIK": "Maks jarak usulan valid km", "NILAI": usulan.max()},
        {"METRIK": "Jumlah titik usulan lebih dekat", "NILAI": int((selisih > 0).sum())},
        {"METRIK": "Jumlah titik existing lebih dekat", "NILAI": int((selisih < 0).sum())},
        {"METRIK": "Rata-rata RPT existing menit", "NILAI": rpt_existing.mean()},
        {"METRIK": "Rata-rata estimasi RPT usulan menit", "NILAI": rpt_usulan.mean()},
        {"METRIK": "Median RPT existing menit", "NILAI": rpt_existing.median()},
        {"METRIK": "Median estimasi RPT usulan menit", "NILAI": rpt_usulan.median()},
        {"METRIK": "Total RPT existing menit", "NILAI": rpt_existing.sum()},
        {"METRIK": "Total estimasi RPT usulan menit", "NILAI": rpt_usulan.sum()},
        {"METRIK": "Estimasi efisiensi total RPT menit", "NILAI": rpt_existing.sum() - rpt_usulan.sum()},
        {"METRIK": "Estimasi efisiensi total RPT persen", "NILAI": (rpt_existing.sum() - rpt_usulan.sum()) / rpt_existing.sum() * 100 if rpt_existing.sum() else np.nan},
        {"METRIK": "Jumlah data valid RPT dibandingkan", "NILAI": len(rpt_comparable)},
        {"METRIK": "Jumlah gangguan RPT existing > 30 menit", "NILAI": int((rpt_comparable[KOLOM_RPT] > 30).sum())},
        {"METRIK": "Jumlah gangguan estimasi RPT usulan > 30 menit", "NILAI": int((rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"] > 30).sum())},
        {"METRIK": "Proporsi RPT existing > 30 menit persen", "NILAI": (rpt_comparable[KOLOM_RPT] > 30).mean() * 100 if len(rpt_comparable) else np.nan},
        {"METRIK": "Proporsi estimasi RPT usulan > 30 menit persen", "NILAI": (rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"] > 30).mean() * 100 if len(rpt_comparable) else np.nan},
    ])


def ringkasan_response_time(data_valid):
    rpt_comparable = data_valid.dropna(subset=[KOLOM_RPT, "ESTIMASI_RPT_USULAN_MENIT"]).copy()
    total = len(rpt_comparable)
    before_gt30 = int((rpt_comparable[KOLOM_RPT] > 30).sum())
    after_gt30 = int((rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"] > 30).sum())

    return pd.DataFrame([
        {
            "SKENARIO": "EXISTING",
            "JUMLAH_DATA_RPT": total,
            "TOTAL_RPT_MENIT": rpt_comparable[KOLOM_RPT].sum(),
            "RATA_RATA_RPT_MENIT": rpt_comparable[KOLOM_RPT].mean(),
            "MEDIAN_RPT_MENIT": rpt_comparable[KOLOM_RPT].median(),
            "MIN_RPT_MENIT": rpt_comparable[KOLOM_RPT].min(),
            "MAKS_RPT_MENIT": rpt_comparable[KOLOM_RPT].max(),
            "JUMLAH_RPT_GT_30_MENIT": before_gt30,
            "PROPORSI_RPT_GT_30_PERSEN": before_gt30 / total * 100 if total else np.nan,
        },
        {
            "SKENARIO": "USULAN_ESTIMASI",
            "JUMLAH_DATA_RPT": total,
            "TOTAL_RPT_MENIT": rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"].sum(),
            "RATA_RATA_RPT_MENIT": rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"].mean(),
            "MEDIAN_RPT_MENIT": rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"].median(),
            "MIN_RPT_MENIT": rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"].min(),
            "MAKS_RPT_MENIT": rpt_comparable["ESTIMASI_RPT_USULAN_MENIT"].max(),
            "JUMLAH_RPT_GT_30_MENIT": after_gt30,
            "PROPORSI_RPT_GT_30_PERSEN": after_gt30 / total * 100 if total else np.nan,
        },
    ])


def ringkasan_ketimpangan(data, kolom_beban):
    rata_rata = data[kolom_beban].mean()
    std = data[kolom_beban].std()
    cv = std / rata_rata * 100 if rata_rata else np.nan
    return {
        "Jumlah Posko": data.shape[0],
        "Total Beban": data[kolom_beban].sum(),
        "Rata-rata Beban per Posko": rata_rata,
        "Minimum Beban": data[kolom_beban].min(),
        "Maksimum Beban": data[kolom_beban].max(),
        "Range Beban": data[kolom_beban].max() - data[kolom_beban].min(),
        "Standar Deviasi": std,
        "Coefficient of Variation (%)": cv,
    }


def rekap_existing(data_valid):
    return (
        data_valid.groupby("UNIT_POSKO_EXISTING")
        .agg(
            JUMLAH_GANGGUAN=(KOLOM_KODE_GANGGUAN, "count"),
            TOTAL_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "sum"),
            RATA_RATA_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "mean"),
            MEDIAN_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "median"),
            MAKS_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "max"),
        )
        .reset_index()
        .sort_values("TOTAL_JARAK_KM", ascending=False)
    )


def rekap_usulan(data_valid):
    return (
        data_valid.groupby(["USULAN_ID", "NAMA_POSKO_USULAN"])
        .agg(
            JUMLAH_GANGGUAN=(KOLOM_KODE_GANGGUAN, "count"),
            TOTAL_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "sum"),
            RATA_RATA_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "mean"),
            MEDIAN_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "median"),
            MAKS_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "max"),
            USULAN_LATITUDE=("USULAN_LATITUDE", "first"),
            USULAN_LONGITUDE=("USULAN_LONGITUDE", "first"),
        )
        .reset_index()
        .sort_values("TOTAL_JARAK_KM", ascending=False)
    )


def rekap_usulan_sebelum_split(data_valid, posko_usulan_awal):
    detail_awal = data_valid.merge(
        posko_usulan_awal,
        left_on="CLUSTER_KMEANS_AWAL",
        right_on="USULAN_AWAL_ID",
        how="left",
        validate="many_to_one",
    )
    detail_awal["JARAK_POSKO_USULAN_AWAL_KE_GANGGUAN_KM"] = haversine_km(
        detail_awal["USULAN_AWAL_LATITUDE"].to_numpy(),
        detail_awal["USULAN_AWAL_LONGITUDE"].to_numpy(),
        detail_awal[KOLOM_LATITUDE].to_numpy(),
        detail_awal[KOLOM_LONGITUDE].to_numpy(),
    )
    return (
        detail_awal.groupby(["USULAN_AWAL_ID", "NAMA_POSKO_USULAN_AWAL"])
        .agg(
            JUMLAH_GANGGUAN=(KOLOM_KODE_GANGGUAN, "count"),
            TOTAL_JARAK_KM=("JARAK_POSKO_USULAN_AWAL_KE_GANGGUAN_KM", "sum"),
            RATA_RATA_JARAK_KM=("JARAK_POSKO_USULAN_AWAL_KE_GANGGUAN_KM", "mean"),
            MEDIAN_JARAK_KM=("JARAK_POSKO_USULAN_AWAL_KE_GANGGUAN_KM", "median"),
            MAKS_JARAK_KM=("JARAK_POSKO_USULAN_AWAL_KE_GANGGUAN_KM", "max"),
            USULAN_AWAL_LATITUDE=("USULAN_AWAL_LATITUDE", "first"),
            USULAN_AWAL_LONGITUDE=("USULAN_AWAL_LONGITUDE", "first"),
        )
        .reset_index()
        .sort_values("TOTAL_JARAK_KM", ascending=False)
    )


def rekap_response_time_per_posko(data_valid):
    rpt_comparable = data_valid.dropna(subset=[KOLOM_RPT, "ESTIMASI_RPT_USULAN_MENIT"]).copy()
    existing = (
        rpt_comparable.groupby("UNIT_POSKO_EXISTING")
        .agg(
            JUMLAH_DATA_RPT=(KOLOM_RPT, "count"),
            RATA_RATA_RPT_EXISTING_MENIT=(KOLOM_RPT, "mean"),
            RATA_RATA_RPT_USULAN_MENIT=("ESTIMASI_RPT_USULAN_MENIT", "mean"),
            TOTAL_RPT_EXISTING_MENIT=(KOLOM_RPT, "sum"),
            TOTAL_RPT_USULAN_MENIT=("ESTIMASI_RPT_USULAN_MENIT", "sum"),
            JUMLAH_RPT_EXISTING_GT_30=(KOLOM_RPT, lambda s: int((s > 30).sum())),
            JUMLAH_RPT_USULAN_GT_30=("ESTIMASI_RPT_USULAN_MENIT", lambda s: int((s > 30).sum())),
        )
        .reset_index()
    )
    existing["EFISIENSI_TOTAL_RPT_MENIT"] = (
        existing["TOTAL_RPT_EXISTING_MENIT"] - existing["TOTAL_RPT_USULAN_MENIT"]
    )
    existing["EFISIENSI_TOTAL_RPT_PERSEN"] = (
        existing["EFISIENSI_TOTAL_RPT_MENIT"] / existing["TOTAL_RPT_EXISTING_MENIT"] * 100
    )
    existing["PROPORSI_RPT_EXISTING_GT_30_PERSEN"] = (
        existing["JUMLAH_RPT_EXISTING_GT_30"] / existing["JUMLAH_DATA_RPT"] * 100
    )
    existing["PROPORSI_RPT_USULAN_GT_30_PERSEN"] = (
        existing["JUMLAH_RPT_USULAN_GT_30"] / existing["JUMLAH_DATA_RPT"] * 100
    )
    return existing.sort_values("EFISIENSI_TOTAL_RPT_MENIT", ascending=False)


def analisis_beban_kerja(data_valid):
    beban_existing = (
        data_valid.groupby("UNIT_POSKO_EXISTING")
        .agg(
            JUMLAH_GANGGUAN=(KOLOM_KODE_GANGGUAN, "count"),
            TOTAL_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "sum"),
            RATA_RATA_JARAK_KM=("JARAK_POSKO_EXISTING_KE_GANGGUAN_KM", "mean"),
            TOTAL_RPT_MENIT=(KOLOM_RPT, "sum"),
            RATA_RATA_RPT_MENIT=(KOLOM_RPT, "mean"),
            JUMLAH_RPT_GT_30=(KOLOM_RPT, lambda x: int((x > 30).sum())),
        )
        .reset_index()
        .rename(columns={"UNIT_POSKO_EXISTING": "POSKO_EXISTING"})
    )
    beban_existing["PROPORSI_RPT_GT_30_PERSEN"] = (
        beban_existing["JUMLAH_RPT_GT_30"] / beban_existing["JUMLAH_GANGGUAN"] * 100
    )
    beban_existing["JUMLAH_GANGGUAN_PER_UNIT_KERJA"] = beban_existing["JUMLAH_GANGGUAN"]
    beban_existing["TOTAL_JARAK_KM_PER_UNIT_KERJA"] = beban_existing["TOTAL_JARAK_KM"]
    beban_existing["TOTAL_RPT_MENIT_PER_UNIT_KERJA"] = beban_existing["TOTAL_RPT_MENIT"]
    beban_existing["JUMLAH_RPT_GT_30_PER_UNIT_KERJA"] = beban_existing["JUMLAH_RPT_GT_30"]

    beban_usulan = (
        data_valid.groupby("NAMA_POSKO_USULAN")
        .agg(
            JUMLAH_GANGGUAN=(KOLOM_KODE_GANGGUAN, "count"),
            TOTAL_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "sum"),
            RATA_RATA_JARAK_KM=("JARAK_POSKO_USULAN_KE_GANGGUAN_KM", "mean"),
            TOTAL_RPT_MENIT=("ESTIMASI_RPT_USULAN_MENIT", "sum"),
            RATA_RATA_RPT_MENIT=("ESTIMASI_RPT_USULAN_MENIT", "mean"),
            JUMLAH_RPT_GT_30=("ESTIMASI_RPT_USULAN_MENIT", lambda x: int((x > 30).sum())),
            USULAN_ID=("USULAN_ID", "first"),
            USULAN_LATITUDE=("USULAN_LATITUDE", "first"),
            USULAN_LONGITUDE=("USULAN_LONGITUDE", "first"),
        )
        .reset_index()
        .rename(columns={"NAMA_POSKO_USULAN": "POSKO_USULAN"})
        .sort_values("USULAN_ID")
    )
    beban_usulan["PROPORSI_RPT_GT_30_PERSEN"] = (
        beban_usulan["JUMLAH_RPT_GT_30"] / beban_usulan["JUMLAH_GANGGUAN"] * 100
    )
    beban_usulan["JUMLAH_GANGGUAN_PER_UNIT_KERJA"] = beban_usulan["JUMLAH_GANGGUAN"]
    beban_usulan["TOTAL_JARAK_KM_PER_UNIT_KERJA"] = beban_usulan["TOTAL_JARAK_KM"]
    beban_usulan["TOTAL_RPT_MENIT_PER_UNIT_KERJA"] = beban_usulan["TOTAL_RPT_MENIT"]
    beban_usulan["JUMLAH_RPT_GT_30_PER_UNIT_KERJA"] = beban_usulan["JUMLAH_RPT_GT_30"]

    rows = []
    for kolom_beban in [
        "JUMLAH_GANGGUAN_PER_UNIT_KERJA",
        "TOTAL_JARAK_KM_PER_UNIT_KERJA",
        "TOTAL_RPT_MENIT_PER_UNIT_KERJA",
        "JUMLAH_RPT_GT_30_PER_UNIT_KERJA",
    ]:
        rows.append({"SKENARIO": "EXISTING", "JENIS_BEBAN": kolom_beban, **ringkasan_ketimpangan(beban_existing, kolom_beban)})
        rows.append({"SKENARIO": "USULAN", "JENIS_BEBAN": kolom_beban, **ringkasan_ketimpangan(beban_usulan, kolom_beban)})

    return beban_existing, beban_usulan, pd.DataFrame(rows)


def buat_grafik_beban(beban_existing, beban_usulan):
    plt.figure(figsize=(12, 6))
    plt.bar(["Existing", "Usulan"], [
        beban_existing["TOTAL_JARAK_KM"].std() / beban_existing["TOTAL_JARAK_KM"].mean() * 100,
        beban_usulan["TOTAL_JARAK_KM"].std() / beban_usulan["TOTAL_JARAK_KM"].mean() * 100,
    ])
    plt.ylabel("CV Total Jarak (%)")
    plt.title("Perbandingan Ketimpangan Beban Total Jarak")
    plt.grid(True, axis="y", alpha=0.3)
    plt.tight_layout()
    plt.savefig(OUTPUT_GRAFIK_BEBAN, dpi=200)
    plt.close()


def _nama_input(file_obj, fallback):
    return getattr(file_obj, "name", None) or fallback


def run_analysis(
    gangguan_file,
    gangguan_sheet=SHEET_GANGGUAN_VALID,
    output_dir="output/web_runs/manual",
    jumlah_posko_existing=TARGET_JUMLAH_TIM_EXISTING,
    progress_callback=None,
):
    global FILE_GANGGUAN_VALID
    global FILE_MASTER_BERSIH
    global SHEET_GANGGUAN_VALID
    global SHEET_MASTER_BERSIH
    global TARGET_JUMLAH_TIM_EXISTING
    global K_MIN
    global K_MAX
    global OUTPUT_EXCEL
    global OUTPUT_GRAFIK_SILHOUETTE
    global OUTPUT_GRAFIK_BEBAN
    global OUTPUT_MAP_HTML
    global OUTPUT_DIR_VISUAL_KMEANS
    global _PROGRESS_CALLBACK

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    FILE_GANGGUAN_VALID = _nama_input(gangguan_file, "uploaded_gangguan.xlsx")
    FILE_MASTER_BERSIH = "Tidak digunakan - input sudah berisi kolom mapping"
    SHEET_GANGGUAN_VALID = gangguan_sheet
    SHEET_MASTER_BERSIH = ""
    TARGET_JUMLAH_TIM_EXISTING = int(jumlah_posko_existing)
    K_MIN = max(2, TARGET_JUMLAH_TIM_EXISTING - 10)
    K_MAX = TARGET_JUMLAH_TIM_EXISTING
    OUTPUT_EXCEL = str(output_dir / "hasil_analisis_terpadu_posko.xlsx")
    OUTPUT_GRAFIK_SILHOUETTE = str(output_dir / "hasil_analisis_terpadu_silhouette.png")
    OUTPUT_GRAFIK_BEBAN = str(output_dir / "hasil_analisis_terpadu_beban_kerja.png")
    OUTPUT_MAP_HTML = str(output_dir / "hasil_kmeans_map.html")
    OUTPUT_DIR_VISUAL_KMEANS = str(output_dir / "visualisasi_kmeans_terpadu")
    _PROGRESS_CALLBACK = progress_callback

    total_tahap = 12

    log_progress(1, total_tahap, "Membaca data gangguan valid ArcGIS...")
    gangguan = pd.read_excel(gangguan_file, sheet_name=gangguan_sheet)
    gangguan[KOLOM_LATITUDE] = pd.to_numeric(gangguan[KOLOM_LATITUDE], errors="coerce")
    gangguan[KOLOM_LONGITUDE] = pd.to_numeric(gangguan[KOLOM_LONGITUDE], errors="coerce")
    gangguan = gangguan.dropna(subset=[KOLOM_LATITUDE, KOLOM_LONGITUDE]).copy()
    if gangguan.empty:
        raise ValueError("Data gangguan tidak memiliki baris dengan koordinat valid.")
    if K_MAX >= len(gangguan):
        raise ValueError(
            f"Jumlah data koordinat valid ({len(gangguan):,}) terlalu sedikit untuk range silhouette "
            f"k={K_MIN} sampai k={K_MAX}. Kurangi jumlah posko existing atau gunakan data lebih banyak."
        )
    print(f"         Jumlah data valid koordinat: {len(gangguan):,}", flush=True)

    log_progress(2, total_tahap, "Memakai kolom mapping dari file upload...")
    data = gangguan.copy()
    data[f"{KOLOM_REGU}_MASTER"] = np.nan
    data[f"{KOLOM_POSKO}_MASTER"] = np.nan
    print("         Lookup master dilewati karena input dianggap sudah benar.", flush=True)

    log_progress(3, total_tahap, "Menghitung jumlah k ideal dengan Silhouette Score...")
    x = data[[KOLOM_LATITUDE, KOLOM_LONGITUDE]].copy()
    scaler = StandardScaler()
    x_scaled = scaler.fit_transform(x)
    silhouette_df, best_k, best_score = hitung_k_ideal(x_scaled)
    buat_grafik_silhouette(silhouette_df)
    print(f"         K terbaik: {best_k}", flush=True)
    print(f"         Silhouette Score terbaik: {best_score:.4f}", flush=True)

    log_progress(4, total_tahap, "Menjalankan K-Means dengan jumlah k terbaik...")
    data_cluster, posko_usulan_awal, posko_usulan, riwayat_split, visualisasi_kmeans = jalankan_kmeans_usulan(
        data,
        x_scaled,
        best_k,
        TARGET_JUMLAH_TIM_EXISTING,
    )
    print(f"         Jumlah posko usulan sebelum split: {len(posko_usulan_awal):,}", flush=True)
    print(f"         Jumlah posko usulan: {len(posko_usulan):,}", flush=True)
    print(f"         Visualisasi KMeans dibuat: {len(visualisasi_kmeans):,} file", flush=True)

    log_progress(5, total_tahap, "Mapping data gangguan ke posko existing...")
    posko_existing_df = pd.DataFrame(POSKO_EXISTING)
    data_cluster["UNIT_POSKO_EXISTING"] = data_cluster.apply(tentukan_unit_posko, axis=1)
    data_cluster = data_cluster.merge(posko_existing_df, left_on="UNIT_POSKO_EXISTING", right_on="UNIT_POSKO", how="left")
    lengkap = (
        data_cluster[KOLOM_KODE_GANGGUAN].notna()
        & data_cluster[KOLOM_LATITUDE].notna()
        & data_cluster[KOLOM_LONGITUDE].notna()
        & data_cluster["POSKO_LATITUDE"].notna()
        & data_cluster["POSKO_LONGITUDE"].notna()
    )
    data_tidak_terpetakan = data_cluster.loc[~lengkap].copy()
    detail_existing = data_cluster.loc[lengkap].copy()
    print(f"         Terpetakan existing: {len(detail_existing):,}", flush=True)
    print(f"         Tidak terpetakan: {len(data_tidak_terpetakan):,}", flush=True)

    log_progress(6, total_tahap, "Menghitung jarak posko existing ke gangguan...")
    detail_existing["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"] = hitung_jarak_existing(detail_existing)
    detail_existing["ESTIMASI_WAKTU_EXISTING_MENIT"] = (
        detail_existing["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"] / KECEPATAN_KM_PER_JAM * 60
    )

    log_progress(7, total_tahap, "Deteksi outlier jarak existing...")
    jarak = detail_existing["JARAK_POSKO_EXISTING_KE_GANGGUAN_KM"]
    q1 = jarak.quantile(0.25)
    q3 = jarak.quantile(0.75)
    iqr = q3 - q1
    batas_outlier_iqr = q3 + OUTLIER_IQR_MULTIPLIER * iqr
    batas_outlier = max(batas_outlier_iqr, BATAS_MIN_OUTLIER_EXISTING_KM)
    detail_existing["BATAS_OUTLIER_IQR_KM"] = batas_outlier_iqr
    detail_existing["BATAS_MIN_OUTLIER_EXISTING_KM"] = BATAS_MIN_OUTLIER_EXISTING_KM
    detail_existing["BATAS_OUTLIER_FINAL_KM"] = batas_outlier
    detail_existing["IS_OUTLIER_EXISTING"] = jarak > batas_outlier
    outlier_existing = detail_existing.loc[detail_existing["IS_OUTLIER_EXISTING"]].copy()
    data_valid = detail_existing.loc[~detail_existing["IS_OUTLIER_EXISTING"]].copy()
    trace_existing = trace_posko_existing(detail_existing)
    print(f"         Batas IQR: > {batas_outlier_iqr:.4f} km", flush=True)
    print(f"         Batas minimal wajar: > {BATAS_MIN_OUTLIER_EXISTING_KM:.4f} km", flush=True)
    print(f"         Batas outlier final: > {batas_outlier:.4f} km", flush=True)
    print(f"         Outlier dibuang: {len(outlier_existing):,}", flush=True)
    print(f"         Data valid untuk perbandingan: {len(data_valid):,}", flush=True)

    log_progress(8, total_tahap, "Mengalokasikan gangguan valid ke posko usulan terdekat...")
    perbandingan = alokasi_ke_posko_usulan(data_valid, posko_usulan)
    perbandingan, efisiensi_total_jarak = tambah_estimasi_response_time(perbandingan)
    output_map_html = buat_peta_cluster_html(perbandingan, posko_usulan, OUTPUT_MAP_HTML)
    print(f"         Efisiensi total jarak untuk estimasi RPT: {efisiensi_total_jarak * 100:.2f}%", flush=True)

    log_progress(9, total_tahap, "Membandingkan parameter jarak dan response time...")
    summary_global = ringkasan_global(perbandingan, outlier_existing, data_tidak_terpetakan)
    summary_response_time = ringkasan_response_time(perbandingan)
    rekap_response_time = rekap_response_time_per_posko(perbandingan)
    print("\nRingkasan global:")
    print(summary_global.to_string(index=False))

    log_progress(10, total_tahap, "Evaluasi beban kerja existing vs usulan...")
    rekap_existing_df = rekap_existing(perbandingan)
    rekap_usulan_awal_df = rekap_usulan_sebelum_split(data_valid, posko_usulan_awal)
    rekap_usulan_df = rekap_usulan(perbandingan)
    beban_existing, beban_usulan, summary_beban_kerja = analisis_beban_kerja(perbandingan)
    buat_grafik_beban(beban_existing, beban_usulan)
    print("\nPerbandingan ketimpangan beban kerja:")
    print(summary_beban_kerja.to_string(index=False))

    log_progress(11, total_tahap, "Menyiapkan parameter output...")
    parameter = pd.DataFrame([
        {"PARAMETER": "FILE_GANGGUAN_VALID", "NILAI": FILE_GANGGUAN_VALID},
        {"PARAMETER": "FILE_MASTER_BERSIH", "NILAI": FILE_MASTER_BERSIH},
        {"PARAMETER": "TARGET_JUMLAH_TIM_EXISTING", "NILAI": TARGET_JUMLAH_TIM_EXISTING},
        {"PARAMETER": "K_MIN", "NILAI": K_MIN},
        {"PARAMETER": "K_MAX", "NILAI": K_MAX},
        {"PARAMETER": "SILHOUETTE_SAMPLE_SIZE", "NILAI": SILHOUETTE_SAMPLE_SIZE},
        {"PARAMETER": "OUTPUT_DIR_VISUAL_KMEANS", "NILAI": OUTPUT_DIR_VISUAL_KMEANS},
        {"PARAMETER": "BEST_K_AWAL_SILHOUETTE", "NILAI": best_k},
        {"PARAMETER": "JUMLAH_POSKO_USULAN_SEBELUM_SPLIT", "NILAI": len(posko_usulan_awal)},
        {"PARAMETER": "JUMLAH_POSKO_USULAN_FINAL", "NILAI": len(posko_usulan)},
        {"PARAMETER": "JUMLAH_CLUSTER_DIBELAH", "NILAI": max(0, TARGET_JUMLAH_TIM_EXISTING - best_k)},
        {"PARAMETER": "BEST_SILHOUETTE_SCORE", "NILAI": best_score},
        {"PARAMETER": "METODE_OUTLIER", "NILAI": f"MAX(IQR > Q3 + {OUTLIER_IQR_MULTIPLIER} * IQR, batas minimal wajar)"},
        {"PARAMETER": "BATAS_MIN_OUTLIER_EXISTING_KM", "NILAI": BATAS_MIN_OUTLIER_EXISTING_KM},
        {"PARAMETER": "BATAS_OUTLIER_IQR_KM", "NILAI": batas_outlier_iqr},
        {"PARAMETER": "BATAS_OUTLIER_EXISTING_KM", "NILAI": batas_outlier},
        {"PARAMETER": "EFISIENSI_TOTAL_JARAK_UNTUK_RPT_PERSEN", "NILAI": efisiensi_total_jarak * 100},
        {"PARAMETER": "RUMUS_ESTIMASI_RPT_USULAN", "NILAI": "ESTIMASI_RPT_USULAN_MENIT = RPT * (1 - EFISIENSI_TOTAL_JARAK)"},
        {"PARAMETER": "ATURAN_BEBAN_KERJA_ULP", "NILAI": "ULP existing dipisah menjadi unit 51 dan 52 dengan koordinat yang sama."},
        {"PARAMETER": "CATATAN_JARAK", "NILAI": "Jarak dihitung dengan Haversine garis lurus, bukan jarak jaringan jalan."},
    ])

    log_progress(12, total_tahap, "Menyimpan hasil analisis terpadu ke Excel...")

    def tulis_excel(path_output):
        with pd.ExcelWriter(path_output) as writer:
            parameter.to_excel(writer, sheet_name="Parameter", index=False)
            silhouette_df.to_excel(writer, sheet_name="Silhouette Score", index=False)
            posko_existing_df.to_excel(writer, sheet_name="Posko Existing", index=False)
            posko_usulan_awal.to_excel(writer, sheet_name="Posko Usulan Sebelum Split", index=False)
            posko_usulan.to_excel(writer, sheet_name="Posko Usulan KMeans", index=False)
            riwayat_split.to_excel(writer, sheet_name="Riwayat Split Cluster", index=False)
            visualisasi_kmeans.to_excel(writer, sheet_name="Visualisasi KMeans", index=False)
            summary_global.to_excel(writer, sheet_name="Summary Global", index=False)
            summary_response_time.to_excel(writer, sheet_name="Summary Response Time", index=False)
            summary_beban_kerja.to_excel(writer, sheet_name="Perbandingan Beban Kerja", index=False)
            trace_existing.to_excel(writer, sheet_name="Trace Posko Existing", index=False)
            rekap_existing_df.to_excel(writer, sheet_name="Rekap Existing Valid", index=False)
            rekap_usulan_awal_df.to_excel(writer, sheet_name="Rekap Usulan Sebelum Split", index=False)
            rekap_usulan_df.to_excel(writer, sheet_name="Rekap Usulan", index=False)
            rekap_response_time.to_excel(writer, sheet_name="Rekap Response Time", index=False)
            beban_existing.to_excel(writer, sheet_name="Beban Kerja Existing", index=False)
            beban_usulan.to_excel(writer, sheet_name="Beban Kerja Usulan", index=False)
            perbandingan.to_excel(writer, sheet_name="Detail Perbandingan Valid", index=False)
            outlier_existing.to_excel(writer, sheet_name="Outlier Existing Dibuang", index=False)
            data_tidak_terpetakan.to_excel(writer, sheet_name="Data Tidak Terpetakan", index=False)

    output_excel_aktual = OUTPUT_EXCEL
    try:
        tulis_excel(output_excel_aktual)
    except PermissionError:
        output_path = Path(OUTPUT_EXCEL)
        output_excel_aktual = str(output_path.with_name(f"{output_path.stem}_{datetime.now():%Y%m%d_%H%M%S}{output_path.suffix}"))
        print(f"         WARNING: {OUTPUT_EXCEL} sedang terkunci. Simpan ke {output_excel_aktual}", flush=True)
        tulis_excel(output_excel_aktual)

    print(f"\n[100.00%] Selesai. File output: {output_excel_aktual}", flush=True)
    print(f"          Folder visualisasi KMeans: {OUTPUT_DIR_VISUAL_KMEANS}", flush=True)
    print(f"          Grafik silhouette: {OUTPUT_GRAFIK_SILHOUETTE}", flush=True)
    print(f"          Grafik beban: {OUTPUT_GRAFIK_BEBAN}", flush=True)
    print(f"          Peta: {output_map_html}", flush=True)
    _PROGRESS_CALLBACK = None
    return {
        "output_dir": output_dir,
        "output_excel": Path(output_excel_aktual),
        "grafik_silhouette": Path(OUTPUT_GRAFIK_SILHOUETTE),
        "grafik_beban": Path(OUTPUT_GRAFIK_BEBAN),
        "map_html": Path(output_map_html),
        "visualisasi_dir": Path(OUTPUT_DIR_VISUAL_KMEANS),
        "tables": {
            "Parameter": parameter,
            "Silhouette Score": silhouette_df,
            "Posko Existing": posko_existing_df,
            "Posko Usulan Sebelum Split": posko_usulan_awal,
            "Posko Usulan KMeans": posko_usulan,
            "Riwayat Split Cluster": riwayat_split,
            "Visualisasi KMeans": visualisasi_kmeans,
            "Summary Global": summary_global,
            "Summary Response Time": summary_response_time,
            "Perbandingan Beban Kerja": summary_beban_kerja,
            "Trace Posko Existing": trace_existing,
            "Rekap Existing Valid": rekap_existing_df,
            "Rekap Usulan Sebelum Split": rekap_usulan_awal_df,
            "Rekap Usulan": rekap_usulan_df,
            "Rekap Response Time": rekap_response_time,
            "Beban Kerja Existing": beban_existing,
            "Beban Kerja Usulan": beban_usulan,
            "Detail Perbandingan Valid": perbandingan,
            "Outlier Existing Dibuang": outlier_existing,
            "Data Tidak Terpetakan": data_tidak_terpetakan,
        },
    }


def main():
    return run_analysis(
        FILE_GANGGUAN_VALID,
        SHEET_GANGGUAN_VALID,
        Path("."),
        TARGET_JUMLAH_TIM_EXISTING,
    )


if __name__ == "__main__":
    main()
