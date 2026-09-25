# -*- coding: utf-8 -*-
"""Standalone web UI: pivot Topos vs BC and export mismatch report."""

from __future__ import annotations

import importlib
import shutil
import sys
import tempfile
from pathlib import Path

import streamlit as st

_APP_DIR = Path(__file__).resolve().parent
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

import engine

importlib.reload(engine)

XLSX = ["xlsx"]

st.set_page_config(
    page_title="TO Pivot Compare",
    page_icon="📊",
    layout="centered",
)

st.title("TO Pivot Compare (Topos vs BC)")
st.caption(
    f"Engine **{engine.ENGINE_VERSION}** — Chi xuat dong pivot lech; "
    "Ten san pham, Topos SL/Unit, BC SL/Unit, Status, Ma bill Topos lech."
)

st.info(
    "**Khong can cai Microsoft Excel** tren may ban, tablet hay dien thoai. "
    "Chi can trinh duyet web (Chrome, Safari, Edge…): upload file `.xlsx`, xem ket qua ngay tren trang, "
    "roi tai file ve neu can. Xu ly file dung thu vien **openpyxl** (Python), khong goi Excel tren may."
)

st.markdown(
    """
**Mot file nguon du** — sheet **Topos** va **BC** (va sheet khac nhu *Ten nha hang*) cung nam trong mot `.xlsx`
(vi du `TO 2009.xlsx`). Ung dung doc truc tiep tu file do; file ket qua giu nguyen cac sheet co san.

**Quy tac doi chieu**
- **Topos:** `TransferFromCode`, `TransferToCode`, `Ngay`/`Date`, `Ma san pham`, tong `So luong`, `Don vi`.
- **BC:** cung bo Transfer + `PostingDate`, `ItemNo` (map sang Ma SP qua `Item Reference No.` hoac ten san pham).
- **Ma bill lech:** gom theo tung dong pivot lech (nhieu bill cach nhau boi `;`).
"""
)

data_file = st.file_uploader(
    "Chon file Excel nguon (.xlsx)",
    type=XLSX,
    key="data",
    help="File chua sheet Topos va BC (vd. TO 2009.xlsx). Chi can chon mot file.",
)


def _write_upload(uploaded, folder: Path) -> Path:
    dest = folder / Path(uploaded.name).name
    dest.write_bytes(uploaded.getvalue())
    return dest


if st.button("Chay doi chieu", type="primary"):
    if data_file is None:
        st.warning("Hay chon file nguon.")
    else:
        work = Path(tempfile.mkdtemp(prefix="to-pivot-"))
        try:
            source = _write_upload(data_file, work)
            output = work / (source.stem + engine.OUTPUT_SUFFIX)
            logs: list[str] = []

            def log(msg: str) -> None:
                logs.append(str(msg))

            with st.spinner("Dang pivot va doi chieu…"):
                result = engine.compare_workbook(source, output, log)
            payload = output.read_bytes()
            summary = (
                f"Xong (engine {result.get('EngineVersion', '?')}): "
                f"{result['MismatchRows']} dong lech "
                f"({result.get('MismatchWithBill', 0)} dong co ma bill Topos)."
            )
            st.success(summary)
            st.caption(
                "Header sheet ket qua: "
                + " | ".join(result.get("ResultHeaders") or [])
            )
            st.download_button(
                "Tai file ket qua (.xlsx)",
                data=payload,
                file_name=output.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            prev_prod = result.get("PreviewProducts") or []
            if prev_prod:
                st.subheader("Xem nhanh tren trinh duyet")
                st.caption(
                    f"Hien toi da {engine.WEB_PREVIEW_ROWS} dong. "
                    "Day du trong sheet ket qua file tai ve."
                )
                st.dataframe(prev_prod, use_container_width=True, hide_index=True)
            with st.expander("Nhat ky xu ly"):
                st.code("\n".join(logs) or "(trong)", language="text")
        except engine.CompareError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"Loi: {exc}")
        finally:
            shutil.rmtree(work, ignore_errors=True)
