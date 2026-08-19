import io
import os
import math
import datetime as dt

import pandas as pd
import streamlit as st

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# PNG
import matplotlib.pyplot as plt


# =========================================
# CONFIG
# =========================================
PREPARERS = {
    "Emre Orhan": {"email": "emre.orhan@kodsan.com.tr", "telefon": "0 543 659 36 73"},
    "Fikret Özdemir": {"email": "fikret.ozdemir@kodsan.com.tr", "telefon": "0 542 349 33 99"},
    "Mert Özen": {"email": "mert.ozen@kodsan.com.tr", "telefon": "0 542 235 33 77"},
}

NOTES = [
    "<b>FİYAT :</b> Fiyatlar EUR bazında olup KDV hariçtir.",
    "<b>DÖVİZ KURU :</b> Fatura tarihindeki TCMB Euro DSK (Döviz Satış Kuru) esas alınarak fatura TÜRK LİRASI üzerinden kesilecek ve ödeme TÜRK LİRASI üzerinden yapılacaktır; gün içerisinde serbest piyasa kuru T.C.M.B’nin belirlemiş olduğu efektif satış kurunu %1 oranında aşmış ise T.C.M.B kuru yerine serbest piyasa kuru geçerli olacaktır.",
    "<b>FATURA :</b> Kodsan Bayileri üzerinden faturalandırılacaktır.",
    "<b>TESLİM ŞEKLİ / GARANTİ :</b> Ankara merkez depomuz teslim. Teklif konusu ürün talimatnamelerine uygun olarak monte edilmiş ve amacı dahilinde kullanılması kaydı ve şartı ile fatura tarihinden itibaren 2 yıl imalat hatalarına karşı garanti kapsamındadır. Kullanıcı ve/veya sistemden kaynaklanan hasarlar garanti kapsamı dışındadır. Müşteri teklife konu olan ürünlerin, yapacağı uygulamada teknik olarak yeterli olduğunu teyit etmektedir.",
    "<b>SİPARİŞ İPTALİ :</b> Müşteri teslimata hazır ürünü teslim almaktan kaçınamaz, iade edemez.",
    "<b>OPSİYON TEMERRÜT :</b> Teklifimiz taşıdığı tarihten itibaren 5 gün süre ile geçerlidir.",
    "<b>ANLAŞMAZLIK :</b> Satışa konu olan ürünlerin bedelinin vade tarihinden sonraki 8 gün içinde ödenmemesi halinde başka ihtara gerek kalmaksızın temerrüde düşmüş kabul edilir ve alıcı bu tarih itibariyle yasal temerrüt faizi ödemekle yükümlüdür. Her türlü anlaşmazlık durumunda ANKARA mahkemeleri ve icra daireleri yetkilidir.",
    "<b>DEVREYE ALMA HİZMETI :</b> Türkiye sınırları içindeki cihazların devreye alma işlemlerinin Kodsan Yetkili Servisleri tarafından yapılması zorunludur. Devreye alma işlemi yapılmayan cihazlar garanti şartlarından yararlanamaz. Devreye alma işlemi ÜCRETSİZ olup 444 50 39 nolu numarayı arayınız.",
]


# =========================================
# HELPERS
# =========================================
def today_tr() -> dt.date:
    return dt.date.today()


def eur_fmt_dec(x: float, decimals: int = 2) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    s = f"{x:,.{decimals}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def calc_discounted(list_price: float, discount_pct: float) -> float:
    return list_price * (1.0 - (discount_pct / 100.0))


def ensure_fonts_registered():
    reg_path = os.path.join("fonts", "DejaVuSans.ttf")
    bold_path = os.path.join("fonts", "DejaVuSans-Bold.ttf")

    if not os.path.exists(reg_path) or not os.path.exists(bold_path):
        raise FileNotFoundError(
            "Font dosyalari eksik. Repo'ya fonts/DejaVuSans.ttf ve fonts/DejaVuSans-Bold.ttf ekleyin."
        )

    try:
        pdfmetrics.getFont("DejaVuSans")
    except Exception:
        pdfmetrics.registerFont(TTFont("DejaVuSans", reg_path))

    try:
        pdfmetrics.getFont("DejaVuSans-Bold")
    except Exception:
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold_path))


# =========================================
# PDF Watermark (diagonal, full-page feel)
# =========================================
def _watermark(canvas, doc, text: str = "KODSAN"):
    w, h = A4
    canvas.saveState()

    font_name = "DejaVuSans-Bold"
    angle = 45

    alpha_ok = True
    try:
        canvas.setFillAlpha(0.04) 
    except Exception:
        alpha_ok = False

    if alpha_ok:
        canvas.setFillColor(colors.HexColor("#BFBFBF"))
    else:
        canvas.setFillColor(colors.HexColor("#EFEFEF"))

    diag = (w * w + h * h) ** 0.5
    target = diag * 0.72

    font_size = 190
    while font_size > 80:
        tw = canvas.stringWidth(text, font_name, font_size)
        if tw <= target:
            break
        font_size -= 2

    canvas.setFont(font_name, font_size)
    canvas.translate(w / 2.0, h / 2.0)
    canvas.rotate(angle)

    tw = canvas.stringWidth(text, font_name, font_size)
    canvas.drawString(-tw / 2.0, -font_size * 0.15, text)
    canvas.restoreState()


def build_pdf_bytes(meta: dict, cart_df: pd.DataFrame, total: float) -> bytes:
    ensure_fonts_registered()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="KODSAN TEKLİF",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "title_style",
        parent=styles["Heading2"],
        fontName="DejaVuSans-Bold",
        fontSize=14,
        leading=16,
        spaceAfter=4,
    )

    normal = ParagraphStyle(
        "normal",
        parent=styles["Normal"],
        fontName="DejaVuSans",
        fontSize=9,
        leading=11,
    )

    cell_model = ParagraphStyle(
        "cell_model",
        parent=styles["Normal"],
        fontName="DejaVuSans-Bold",
        fontSize=7.4,
        leading=9,
    )

    cell_desc = ParagraphStyle(
        "cell_desc",
        parent=styles["Normal"],
        fontName="DejaVuSans",
        fontSize=7.2,
        leading=9,
    )

    cell_num = ParagraphStyle(
        "cell_num",
        parent=styles["Normal"],
        fontName="DejaVuSans",
        fontSize=7.4,
        leading=9,
        alignment=2, 
    )

    small = ParagraphStyle(
        "small",
        parent=styles["BodyText"],
        fontName="DejaVuSans",
        fontSize=8.2,
        leading=10,
    )

    story = []
    story.append(Paragraph("KODSAN TEKLİF", title_style))
    story.append(Spacer(1, 3 * mm))

    info_data = [
        ["Tarih", meta["tarih"]],
        ["Geçerlilik", meta["gecerlilik"]],
        ["Firma İsmi", meta["firma"]],
        ["Yetkili İsmi", meta["yetkili"]],
        ["Proje İsmi", meta["proje"]],
        ["Teklifi Hazırlayan", meta["hazirlayan"]],
        ["E-mail", meta["email"]],
        ["Telefon", meta["telefon"]],
    ]
    info_tbl = Table(info_data, colWidths=[34 * mm, 150 * mm])
    info_tbl.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "DejaVuSans"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(info_tbl)
    story.append(Spacer(1, 5 * mm))

    header = ["Model", "Açıklama", "Adet", "Birim (EUR)", "Tutar (EUR)"]
    rows = [header]

    for _, r in cart_df.iterrows():
        # Yapıştırılan metindeki satır atlamalarını (Enter) PDF formatına (br) çevirme
        desc_html = str(r["AÇIKLAMA"]).replace("\n", "<br/>")
        
        rows.append(
            [
                Paragraph(str(r["MODEL"]), cell_model),
                Paragraph(desc_html, cell_desc), 
                Paragraph(str(int(r["ADET"])), cell_num),
                Paragraph(eur_fmt_dec(float(r["BİRİM (EUR)"]), 2), cell_num),
                Paragraph(eur_fmt_dec(float(r["TOPLAM (EUR)"]), 2), cell_num),
            ]
        )

    prod_tbl = Table(
        rows,
        colWidths=[36 * mm, 88 * mm, 10 * mm, 25 * mm, 25 * mm],
        repeatRows=1,
    )
    prod_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8.2),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (2, 1), (4, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )

    story.append(prod_tbl)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(f"<b>Toplam:</b> {eur_fmt_dec(total, 2)} EUR + KDV", normal))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("<b>NOTLAR</b>", normal))
    story.append(Spacer(1, 1.5 * mm))
    for n in NOTES:
        story.append(Paragraph(n, small))
        story.append(Spacer(1, 1.2 * mm))

    doc.build(
        story,
        onFirstPage=lambda c, d: _watermark(c, d, "KODSAN"),
        onLaterPages=lambda c, d: _watermark(c, d, "KODSAN"),
    )

    pdf = buf.getvalue()
    buf.close()
    return pdf


# =========================================
# PNG output
# =========================================
def build_table_png_bytes(cart_df: pd.DataFrame, meta: dict, total: float) -> bytes:
    view = cart_df.copy()
    view["BİRİM (EUR)"] = view["BİRİM (EUR)"].map(lambda v: eur_fmt_dec(float(v), 2))
    view["TOPLAM (EUR)"] = view["TOPLAM (EUR)"].map(lambda v: eur_fmt_dec(float(v), 2))
    view["ADET"] = view["ADET"].astype(int).astype(str)

    title = f"{meta['firma']} | {meta['proje']} | Toplam: {eur_fmt_dec(total, 2)} EUR + KDV"

    fig_h = 1.2 + 0.35 * max(1, len(view))
    fig, ax = plt.subplots(figsize=(12, fig_h))
    ax.axis("off")
    ax.set_title(title, fontsize=11, pad=10)

    tbl = ax.table(
        cellText=view[["MODEL", "AÇIKLAMA", "ADET", "BİRİM (EUR)", "TOPLAM (EUR)"]].values,
        colLabels=["MODEL", "AÇIKLAMA", "ADET", "BİRİM (EUR)", "TOPLAM (EUR)"],
        cellLoc="left",
        colLoc="left",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.5)
    tbl.scale(1, 1.2)

    buf = io.BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# =========================================
# Streamlit UI
# =========================================
st.set_page_config(
    page_title="Teklif Oluşturucu (Manuel Giriş)",
    page_icon="📄",
    layout="wide"
)

if "cart" not in st.session_state:
    st.session_state.cart = []

with st.sidebar:
    st.header("Teklif Bilgileri")

    tarih = today_tr()
    gecerlilik = tarih + dt.timedelta(days=5)

    firma = st.text_input("FİRMA İSMİ", value="")
    yetkili = st.text_input("YETKİLİ İSMİ", value="")
    proje = st.text_input("PROJE İSMİ", value="")
    iskonto = st.number_input("İSKONTO ORANI (%)", min_value=0.0, max_value=100.0, value=30.0, step=0.5)

    hazirlayan = st.selectbox("TEKLİFİ HAZIRLAYAN", list(PREPARERS.keys()), index=0)
    email = PREPARERS[hazirlayan]["email"]
    telefon = PREPARERS[hazirlayan]["telefon"]

    st.text_input("EMAİL", value=email, disabled=True)
    st.text_input("TELEFON", value=telefon, disabled=True)

    st.divider()
    st.caption(f"Tarih: {tarih.strftime('%d.%m.%Y')}")
    st.caption(f"Geçerlilik: {gecerlilik.strftime('%d.%m.%Y')} (5 gün)")

    st.divider()
    if st.button("Sepeti sıfırla", use_container_width=True):
        st.session_state.cart = []
        st.rerun()

colA, colB = st.columns([1.1, 1.2], gap="large")

with colA:
    st.subheader("Ürün Ekle (Manuel)")

    # Varsayılan metin şablonu
    ornek_model = "KHE-P01-5-C10-CS-00 PLAKALI EŞANJÖR"
    ornek_aciklama = """Kapasite : 796,62 kW
Primer Devre : 85,00°C / 65,00°C - 48,37 kPa
Sekonder Devre : 60,00°C / 80,00°C - 48,69 kPa
Plaka ve Conta Malzemesi : 0,5 mm - AISI 316L - EPDM
Gövde Malzemesi : Boyalı Karbon Çelik
İşletme ve Test Basıncı : 10 / 15 Bar
Bağlantı Malzemesi ve Çapı : 2" Dıştan Dişli CS"""

    model_input = st.text_input("Ürün Kodu / Başlık", value=ornek_model)
    aciklama_input = st.text_area("İçerik / Teknik Özellikler", value=ornek_aciklama, height=220)

    c1, c2 = st.columns(2)
    with c1:
        qty = st.number_input("Adet", min_value=1, value=1, step=1)
    with c2:
        list_price = st.number_input("Birim Liste Fiyatı (EUR)", min_value=0.0, value=1000.0, step=10.0, format="%.2f")

    unit = calc_discounted(list_price, float(iskonto))

    st.markdown("**Seçilen ürün özeti**")
    st.write(f"**Liste fiyatı:** {eur_fmt_dec(list_price, 2)} EUR")
    st.write(f"**İskontolu birim fiyat:** {eur_fmt_dec(unit, 2)} EUR + KDV")

    if st.button("Sepete ekle", type="primary", use_container_width=True):
        if not model_input.strip() and not aciklama_input.strip():
            st.error("Lütfen ürün kodu veya açıklama giriniz.")
        else:
            found = False
            for r in st.session_state.cart:
                # Birebir aynı model adı, açıklaması ve liste fiyatı varsa adedi güncelle
                if r["MODEL"] == model_input and r["AÇIKLAMA"] == aciklama_input and r["LİSTE FİYATI"] == list_price:
                    r["ADET"] = int(r["ADET"]) + int(qty)
                    found = True
                    break
            
            if not found:
                st.session_state.cart.append(
                    {
                        "MODEL": model_input,
                        "AÇIKLAMA": aciklama_input,
                        "LİSTE FİYATI": list_price,
                        "ADET": int(qty),
                    }
                )
            st.rerun()

with colB:
    st.subheader("Teklif Kalemleri")

    if len(st.session_state.cart) == 0:
        st.info("Sepet boş. Soldan manuel ürün ekleyebilirsiniz.")
    else:
        cart_df = pd.DataFrame(st.session_state.cart)
        cart_df["BİRİM (EUR)"] = cart_df["LİSTE FİYATI"].apply(lambda p: calc_discounted(float(p), float(iskonto)))
        cart_df["TOPLAM (EUR)"] = cart_df["BİRİM (EUR)"] * cart_df["ADET"].astype(int)
        total = float(cart_df["TOPLAM (EUR)"].sum())

        edit_df = cart_df[["MODEL", "AÇIKLAMA", "ADET", "BİRİM (EUR)", "TOPLAM (EUR)"]].copy()
        edit_df["SİL"] = False

        edited = st.data_editor(
            edit_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "ADET": st.column_config.NumberColumn("ADET", min_value=1, step=1),
                "BİRİM (EUR)": st.column_config.NumberColumn("BİRİM (EUR)", format="%.2f"),
                "TOPLAM (EUR)": st.column_config.NumberColumn("TOPLAM (EUR)", format="%.2f"),
                "SİL": st.column_config.CheckboxColumn("SİL"),
            },
            disabled=["MODEL", "AÇIKLAMA", "BİRİM (EUR)", "TOPLAM (EUR)"],
            key="cart_editor",
        )

        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("Değişiklikleri uygula", use_container_width=True):
                keep = []
                for idx, r in edited.iterrows():
                    if bool(r.get("SİL", False)):
                        continue
                    keep.append(
                        {
                            "MODEL": r["MODEL"],
                            "AÇIKLAMA": r["AÇIKLAMA"],
                            "LİSTE FİYATI": float(cart_df.loc[idx, "LİSTE FİYATI"]),
                            "ADET": int(r["ADET"]),
                        }
                    )
                st.session_state.cart = keep
                st.rerun()

        with c2:
            st.metric("Kümülatif Toplam", f"{eur_fmt_dec(total, 2)} EUR + KDV")

        st.divider()

        meta = {
            "tarih": tarih.strftime("%d.%m.%Y"),
            "gecerlilik": gecerlilik.strftime("%d.%m.%Y"),
            "firma": firma.strip() or "-",
            "yetkili": yetkili.strip() or "-",
            "proje": proje.strip() or "-",
            "hazirlayan": hazirlayan,
            "email": email,
            "telefon": telefon,
        }

        pdf_bytes = build_pdf_bytes(meta, cart_df, total)
        st.download_button(
            label="PDF indir (KODSAN TEKLİF)",
            data=pdf_bytes,
            file_name=f"Kodsan_Teklif_{meta['firma'].replace(' ', '_')}_{meta['tarih'].replace('.', '-')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

        png_bytes = build_table_png_bytes(cart_df, meta, total)
        st.download_button(
            label="PNG indir (ekran görüntüsü gibi)",
            data=png_bytes,
            file_name=f"Teklif_{meta['firma'].replace(' ', '_')}_{meta['tarih'].replace('.', '-')}.png",
            mime="image/png",
            use_container_width=True,
        )

st.caption("Fiyatlar EUR bazında; KDV hariç gösterilir. İskonto, liste fiyatına yüzde olarak uygulanır.")
