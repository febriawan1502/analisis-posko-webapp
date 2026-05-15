from __future__ import annotations

import numbers
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from analysis_engine import run_analysis


APP_DIR = Path(__file__).resolve().parent
OUTPUT_ROOT = APP_DIR / "output" / "web_runs"
SAMPLE_DATA_PATH = APP_DIR / "contoh_data_100_baris.xlsx"
DEFAULT_GANGGUAN_SHEET = "data_gangguan_valid_14052026"

GANGGUAN_HEADERS = [
    "NO_LAPORAN",
    "IDPEL",
    "LAT_PLGN",
    "LONG_PLGN",
    "TGL_LAPOR",
    "BULAN",
    "JAM_LAPOR",
    "RPT",
    "NAMA_REGU",
    "NAMA_POSKO",
]


def build_template(headers: list[str], sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(columns=headers).to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


def get_sample_data() -> bytes:
    if SAMPLE_DATA_PATH.exists():
        return SAMPLE_DATA_PATH.read_bytes()
    return build_template(GANGGUAN_HEADERS, DEFAULT_GANGGUAN_SHEET)


def as_named_buffer(data: bytes, name: str) -> BytesIO:
    buffer = BytesIO(data)
    buffer.name = name
    return buffer


def get_sheet_names(data: bytes, filename: str) -> list[str]:
    try:
        return pd.ExcelFile(BytesIO(data)).sheet_names
    except Exception as exc:
        raise ValueError(f"Gagal membaca sheet dari {filename}: {exc}") from exc


def validate_workbook(
    data: bytes,
    filename: str,
    sheet_name: str,
    required_columns: set[str],
    coordinate_check: bool = False,
) -> None:
    try:
        workbook = pd.ExcelFile(BytesIO(data))
    except Exception as exc:
        raise ValueError(f"{filename} bukan file Excel yang bisa dibaca: {exc}") from exc

    if sheet_name not in workbook.sheet_names:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan di {filename}. Sheet tersedia: {available}")

    try:
        header_df = pd.read_excel(BytesIO(data), sheet_name=sheet_name, nrows=0)
    except Exception as exc:
        raise ValueError(f"Gagal membaca header {filename}: {exc}") from exc

    missing = sorted(required_columns - set(header_df.columns))
    if missing:
        raise ValueError(f"{filename} kurang kolom wajib: {', '.join(missing)}")

    if coordinate_check:
        coord_df = pd.read_excel(BytesIO(data), sheet_name=sheet_name, usecols=["LAT_PLGN", "LONG_PLGN"])
        lat = pd.to_numeric(coord_df["LAT_PLGN"], errors="coerce")
        lon = pd.to_numeric(coord_df["LONG_PLGN"], errors="coerce")
        if coord_df.empty or not (lat.notna() & lon.notna()).any():
            raise ValueError(f"{filename} tidak memiliki koordinat numerik valid di LAT_PLGN/LONG_PLGN.")


def zip_images(result: dict) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for key in ["grafik_silhouette", "grafik_beban"]:
            path = Path(result[key])
            if path.exists():
                archive.write(path, arcname=path.name)

        visual_dir = Path(result["visualisasi_dir"])
        if visual_dir.exists():
            for image_path in sorted(visual_dir.glob("*.png")):
                archive.write(image_path, arcname=f"visualisasi_kmeans/{image_path.name}")
    return buffer.getvalue()


def show_table(title: str, table: pd.DataFrame, height: int = 420) -> None:
    st.subheader(title)
    st.caption(f"{len(table):,} baris, {len(table.columns):,} kolom")
    st.dataframe(table, use_container_width=True, height=height)


def metric_from_summary(summary: pd.DataFrame, metric_name: str):
    row = summary.loc[summary["METRIK"] == metric_name, "NILAI"]
    if row.empty:
        return "-"
    value = row.iloc[0]
    if isinstance(value, numbers.Real) and not isinstance(value, numbers.Integral):
        return f"{value:,.2f}"
    return f"{value:,}" if isinstance(value, numbers.Integral) else value


def render_dashboard(result: dict) -> None:
    tables = result["tables"]

    st.success("Analisis selesai.")
    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button(
            "Download Excel hasil",
            data=Path(result["output_excel"]).read_bytes(),
            file_name=Path(result["output_excel"]).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_b:
        st.download_button(
            "Download semua gambar",
            data=zip_images(result),
            file_name="gambar_analisis_terpadu_posko.zip",
            mime="application/zip",
            use_container_width=True,
        )
    map_path = Path(result["map_html"])
    if map_path.exists():
        st.download_button(
            "Download HTML map",
            data=map_path.read_bytes(),
            file_name=map_path.name,
            mime="text/html",
            use_container_width=True,
        )

    tabs = st.tabs([
        "Summary",
        "Response Time",
        "Beban Kerja",
        "Posko",
        "Detail Perbandingan",
        "Outlier",
        "Data Tidak Terpetakan",
        "Map",
        "Grafik",
        "Visualisasi KMeans",
    ])

    with tabs[0]:
        summary = tables["Summary Global"]
        metric_cols = st.columns(4)
        metric_cols[0].metric("Data dibandingkan", metric_from_summary(summary, "Jumlah data dibandingkan"))
        metric_cols[1].metric("Penghematan km", metric_from_summary(summary, "Penghematan total km"))
        metric_cols[2].metric("Penghematan %", metric_from_summary(summary, "Penghematan total persen"))
        metric_cols[3].metric("Tidak terpetakan", metric_from_summary(summary, "Jumlah data tidak terpetakan"))
        show_table("Summary Global", summary)
        show_table("Parameter", tables["Parameter"], height=300)

    with tabs[1]:
        show_table("Summary Response Time", tables["Summary Response Time"], height=250)
        show_table("Rekap Response Time", tables["Rekap Response Time"])

    with tabs[2]:
        show_table("Perbandingan Beban Kerja", tables["Perbandingan Beban Kerja"], height=320)
        show_table("Beban Kerja Existing", tables["Beban Kerja Existing"])
        show_table("Beban Kerja Usulan", tables["Beban Kerja Usulan"])

    with tabs[3]:
        show_table("Posko Existing", tables["Posko Existing"], height=320)
        show_table("Posko Usulan Sebelum Split", tables["Posko Usulan Sebelum Split"], height=320)
        show_table("Posko Usulan KMeans", tables["Posko Usulan KMeans"], height=320)
        show_table("Riwayat Split Cluster", tables["Riwayat Split Cluster"], height=260)
        show_table("Rekap Existing Valid", tables["Rekap Existing Valid"], height=320)
        show_table("Rekap Usulan Sebelum Split", tables["Rekap Usulan Sebelum Split"], height=320)
        show_table("Rekap Usulan", tables["Rekap Usulan"], height=320)

    with tabs[4]:
        show_table("Detail Perbandingan Valid", tables["Detail Perbandingan Valid"], height=560)

    with tabs[5]:
        show_table("Outlier Existing Dibuang", tables["Outlier Existing Dibuang"], height=420)
        show_table("Trace Posko Existing", tables["Trace Posko Existing"], height=320)

    with tabs[6]:
        show_table("Data Tidak Terpetakan", tables["Data Tidak Terpetakan"], height=560)

    with tabs[7]:
        map_path = Path(result["map_html"])
        if map_path.exists():
            components.html(map_path.read_text(encoding="utf-8"), height=720, scrolling=False)
        else:
            st.info("Map belum tersedia.")

    with tabs[8]:
        for label, key in [
            ("Silhouette Score", "grafik_silhouette"),
            ("Beban Kerja", "grafik_beban"),
        ]:
            path = Path(result[key])
            if path.exists():
                st.subheader(label)
                st.image(str(path), use_container_width=True)
        show_table("Silhouette Score", tables["Silhouette Score"], height=300)

    with tabs[9]:
        show_table("Daftar Visualisasi KMeans", tables["Visualisasi KMeans"], height=260)
        visual_dir = Path(result["visualisasi_dir"])
        images = sorted(visual_dir.glob("*.png")) if visual_dir.exists() else []
        if not images:
            st.info("Belum ada gambar visualisasi KMeans.")
        for image_path in images:
            st.subheader(image_path.stem)
            st.image(str(image_path), use_container_width=True)


def main() -> None:
    st.set_page_config(page_title="Analisis Terpadu Posko", layout="wide")
    st.title("Analisis Terpadu Posko")

    with st.sidebar:
        st.header("Template")
        st.download_button(
            "Template data upload",
            data=build_template(GANGGUAN_HEADERS, DEFAULT_GANGGUAN_SHEET),
            file_name="template_data_upload.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Contoh data 100 baris",
            data=get_sample_data(),
            file_name="contoh_data_100_baris.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    upload_col_a, upload_col_b = st.columns(2)
    gangguan_sheet = None
    with upload_col_a:
        gangguan_upload = st.file_uploader("Upload Excel data", type=["xlsx"])
        if gangguan_upload is not None:
            try:
                sheet_names = get_sheet_names(gangguan_upload.getvalue(), gangguan_upload.name)
            except ValueError as exc:
                st.error(str(exc))
                sheet_names = []

            if sheet_names:
                default_index = (
                    sheet_names.index(DEFAULT_GANGGUAN_SHEET)
                    if DEFAULT_GANGGUAN_SHEET in sheet_names
                    else 0
                )
                gangguan_sheet = st.selectbox("Pilih sheet data", sheet_names, index=default_index)
    with upload_col_b:
        jumlah_posko_existing = st.number_input(
            "Jumlah posko existing",
            min_value=2,
            max_value=500,
            value=28,
            step=1,
            help="Range silhouette otomatis: jumlah posko - 10 sampai jumlah posko.",
        )
        k_min = max(2, int(jumlah_posko_existing) - 10)
        k_max = int(jumlah_posko_existing)
        st.caption(f"Range silhouette: k={k_min} sampai k={k_max}")

    run_disabled = gangguan_upload is None or gangguan_sheet is None
    if st.button("Jalankan analisis", type="primary", disabled=run_disabled):
        gangguan_bytes = gangguan_upload.getvalue()

        try:
            validate_workbook(
                gangguan_bytes,
                gangguan_upload.name,
                gangguan_sheet,
                set(GANGGUAN_HEADERS),
                coordinate_check=True,
            )
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = OUTPUT_ROOT / run_id
        progress_bar = st.progress(0)
        status = st.empty()

        def update_progress(stage: int, total: int, message: str) -> None:
            progress_bar.progress(stage / total)
            status.write(f"{stage}/{total} - {message}")

        try:
            result = run_analysis(
                as_named_buffer(gangguan_bytes, gangguan_upload.name),
                gangguan_sheet,
                output_dir,
                int(jumlah_posko_existing),
                update_progress,
            )
        except Exception as exc:
            st.error(f"Analisis gagal: {exc}")
            st.stop()

        st.session_state["analysis_result"] = result

    if "analysis_result" in st.session_state:
        render_dashboard(st.session_state["analysis_result"])


if __name__ == "__main__":
    main()
