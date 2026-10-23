import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import zipfile
import re

st.set_page_config(page_title="Bulk QR Sticker Generator", layout="wide")
st.title("📱 Bulk QR Sticker Generator (1 QR = 1 Sticker)")
st.markdown("Upload **Excel** or **TXT** → each row/line becomes **one sticker** (QR + clean label).")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("⚙️ Sticker Settings")
    box_size = st.slider("QR Box Size", 5, 20, 10)
    border = st.slider("QR Border", 1, 8, 2)
    label_size = st.slider("Label Font Size", 15, 60, 25)
    sticker_padding = st.slider("Sticker Padding", 5, 40, 20)
    stickers_per_sheet = st.number_input(
        "Stickers per A4 sheet (PDF only)", 1, 40, 8
    )

# ---------------- Helpers ----------------
def clean_value(v):
    """
    Convert raw values into clean labels:
      '2424.0'   -> '2424'
      '2424.50'  -> '2424.5'
      '  2424  ' -> '2424'
      '1e3'      -> '1000'
      'ABC123'   -> 'ABC123'
    """
    if v is None:
        return ""
    s = str(v).strip()

    # Try to treat as a number
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))            # 2424.0 -> "2424"
        else:
            # trim trailing zeros but keep meaningful decimals
            return ("%g" % f)             # 2424.50 -> "2424.5"
    except (ValueError, TypeError):
        return s                          # non-numeric -> keep as-is


def make_qr(text, box_size, border):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )
    qr.add_data(str(text))
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


def make_sticker(text, box_size, border, label_size, padding):
    """ONE sticker = QR + clean label underneath + light border."""
    qr_img = make_qr(text, box_size, border)

    try:
        font = ImageFont.truetype("arial.ttf", label_size)
    except Exception:
        font = ImageFont.load_default()

    tmp = Image.new("RGB", (1, 1))
    d = ImageDraw.Draw(tmp)
    bbox = d.textbbox((0, 0), str(text), font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    gap = 8
    sticker_w = max(qr_img.width, text_w) + padding * 2
    sticker_h = qr_img.height + gap + text_h + padding * 2

    canvas = Image.new("RGB", (sticker_w, sticker_h), "white")
    draw = ImageDraw.Draw(canvas)

    qr_x = (sticker_w - qr_img.width) // 2
    canvas.paste(qr_img, (qr_x, padding))

    text_x = (sticker_w - text_w) // 2
    text_y = padding + qr_img.height + gap
    draw.text((text_x, text_y), str(text), fill="black", font=font)

    draw.rectangle(
        [(0, 0), (sticker_w - 1, sticker_h - 1)],
        outline="#CCCCCC",
        width=1,
    )
    return canvas


def safe_name(v, idx):
    s = "".join(c for c in str(v) if c.isalnum() or c in "-_") or "qr"
    return f"{idx:04d}_{s}.png"


def stickers_to_pdf_bytes(stickers, per_sheet=8):
    if not stickers:
        return None

    A4_W, A4_H = 1240, 1754   # A4 @ 150 DPI
    margin = 30

    stickers_per_row = max(
        1, int((A4_W - 2 * margin) / (max(s.width for s in stickers) + 10))
    )

    pages = []
    for i in range(0, len(stickers), per_sheet):
        batch = stickers[i:i + per_sheet]
        page = Image.new("RGB", (A4_W, A4_H), "white")

        for j, sticker in enumerate(batch):
            r, c = divmod(j, stickers_per_row)
            x = margin + c * (sticker.width + 10)
            y = margin + r * (sticker.height + 10)
            page.paste(sticker, (x, y))

        pages.append(page)

    buf = BytesIO()
    first, rest = pages[0], pages[1:]
    if rest:
        first.save(buf, "PDF", resolution=150.0, save_all=True, append_images=rest)
    else:
        first.save(buf, "PDF", resolution=150.0)
    return buf.getvalue()


# ---------------- Upload ----------------
uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx / .xls) or TXT (.txt)",
    type=["xlsx", "xls", "txt"],
)

if uploaded_file:
    file_ext = uploaded_file.name.lower().rsplit(".", 1)[-1]

    # -------- Parse file --------
    if file_ext in ("xlsx", "xls"):
        df = pd.read_excel(uploaded_file)
        st.subheader("📄 Preview of Uploaded Excel")
        st.dataframe(df.head(20), use_container_width=True)

        columns = df.columns.tolist()
        value_col = st.selectbox(
            "Select column that contains the QR value / number", columns
        )

        # raw values (will be cleaned at generation time)
        raw_values = df[value_col].dropna().tolist()

    else:  # txt
        content = uploaded_file.read().decode("utf-8", errors="ignore")
        lines = [ln.strip() for ln in content.splitlines()]
        # drop blank lines and pure whitespace
        raw_values = [ln for ln in lines if ln]

        st.subheader("📄 Preview of Uploaded TXT")
        st.code(
            "\n".join(raw_values[:20])
            + (f"\n... ({len(raw_values)} total lines)" if len(raw_values) > 20 else ""),
            language="text",
        )
        value_col = "TXT Lines"

    # -------- Generate --------
    if st.button("🚀 Generate Stickers", type="primary"):
        if not raw_values:
            st.error("No values found in the uploaded file.")
            st.stop()

        # Clean values: 2424.0 -> "2424"
        values = [clean_value(v) for v in raw_values]
        values = [v for v in values if v]  # drop empties after cleaning

        if not values:
            st.error("All values were empty after cleaning.")
            st.stop()

        st.success(f"Generating {len(values)} stickers...")

        stickers = []
        progress = st.progress(0)
        for i, v in enumerate(values):
            stickers.append(
                make_sticker(v, box_size, border, label_size, sticker_padding)
            )
            if i % max(1, len(values) // 20) == 0 or i == len(values) - 1:
                progress.progress((i + 1) / len(values))
        progress.empty()

        # -------- Preview --------
        st.subheader("🖼️ Sticker Preview (first 8)")
        preview_cols = st.columns(4)
        for i, s in enumerate(stickers[:8]):
            with preview_cols[i % 4]:
                st.image(s, caption=f"{values[i]}")

        # -------- Downloads --------
        st.subheader("⬇️ Download")
        col1, col2 = st.columns(2)

        # 1) Individual stickers ZIP
        indiv_zip = BytesIO()
        with zipfile.ZipFile(indiv_zip, "w") as zf:
            for i, (v, sticker) in enumerate(zip(values, stickers), start=1):
                ib = BytesIO()
                sticker.save(ib, format="PNG")
                zf.writestr(safe_name(v, i), ib.getvalue())
        indiv_zip.seek(0)

        with col1:
            st.download_button(
                "📱 Download Individual Stickers (ZIP)",
                data=indiv_zip,
                file_name="qr_stickers.zip",
                mime="application/zip",
            )

        # 2) PDF
        pdf_bytes = stickers_to_pdf_bytes(stickers, per_sheet=stickers_per_sheet)
        if pdf_bytes:
            with col2:
                st.download_button(
                    "📄 Download Stickers PDF (A4)",
                    data=pdf_bytes,
                    file_name="qr_stickers.pdf",
                    mime="application/pdf",
                )

        # 3) CSV mapping
        mapping = pd.DataFrame({
            "value": values,
            "sticker_file": [safe_name(v, i) for i, v in enumerate(values, start=1)],
        })
        st.download_button(
            "📄 Download CSV Mapping",
            data=mapping.to_csv(index=False),
            file_name="sticker_mapping.csv",
            mime="text/csv",
        )

else:
    st.info(
        "👆 Upload an Excel (.xlsx/.xls) or TXT (.txt) file to begin.\n\n"
        "• Excel: pick the column with QR values\n"
        "• TXT: one value per line\n\n"
        "Numbers like `2424.0` will be cleaned to `2424` automatically."
    )