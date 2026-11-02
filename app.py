import streamlit as st
import pandas as pd
import plotly.express as px
import os
import smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from openpyxl.utils import get_column_letter
import io

# Sayfa Genişlik Ayarı ve Temiz Arayüz
st.set_page_config(page_title="⚡ Elektrik Depo Takip", layout="wide")

# --- ÖZEL KUTUCUK VE GÖLGE TASARIMI (CSS) ---
st.markdown("""
<style>
    .metric-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
        border: 1px solid #e2e8f0;
        text-align: center;
        margin-bottom: 15px;
    }
    .metric-title {
        font-size: 13px;
        color: #64748b;
        font-weight: 600;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 24px;
        color: #1e293b;
        font-weight: 700;
    }
    .id-card {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        padding: 15px;
        border-radius: 10px;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = "ELEKTRİK_DEPO_TAKİP.xlsx"
IMAGE_FOLDER = "urun_resimleri"
LOG_FILE = "DEPO_HAREKET_KAYITLARI.xlsx"

ADMIN_USER = "admin"
ADMIN_PASSWORD = "depo123"

# --- E-POSTA AYARLARI ---
GÖNDERİCİ_GMAIL = "kendi_gmailiniz@gmail.com"
GMAIL_UYGULAMA_SİFRESİ = "google_aldiginiz_16_haneli_kod" 
ALICI_EMAIL = "hedef_eposta@gmail.com"

def log_action(urun_kodu, urun_adi, islem_turu, miktar, detay=""):
    yeni_kayit = {
        "Tarih / Saat": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Ürün Kodu": urun_kodu,
        "Ürün Adı": urun_adi,
        "İşlem Türü": islem_turu,
        "Miktar": miktar,
        "Detay / Konum": detay
    }
    if os.path.exists(LOG_FILE):
        log_df = pd.read_excel(LOG_FILE)
        log_df = pd.concat([log_df, pd.DataFrame([yeni_kayit])], ignore_index=True)
    else:
        log_df = pd.DataFrame([yeni_kayit])
    with pd.ExcelWriter(LOG_FILE, engine='openpyxl') as writer:
        log_df.to_excel(writer, index=False, sheet_name='Hareketler')
        worksheet = writer.sheets['Hareketler']
        for i in range(1, 7):
            worksheet.column_dimensions[get_column_letter(i)].width = 25

def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if len(df.columns) == 8:
            df["Birim Fiyat"] = 0.0
            df["Tedarikçi Bilgisi"] = "Belirtilmedi"
        elif len(df.columns) == 9:
            df["Tedarikçi Bilgisi"] = "Belirtilmedi"
        df = df.iloc[:, :10]
        df.columns = ["Ürün Kategorisi", "Marka", "Ürün Adı", "Ürün Kodu", "Miktar", "Emniyet Faktörü", "Çekmece Numarası", "Görsel", "Birim Fiyat", "Tedarikçi Bilgisi"]
        df = df.dropna(subset=["Ürün Kodu"])
        df["Ürün Kodu"] = df["Ürün Kodu"].astype(str).str.strip()
        df["Miktar"] = pd.to_numeric(df["Miktar"], errors='coerce').fillna(0).astype(int)
        df["Emniyet Faktörü"] = pd.to_numeric(df["Emniyet Faktörü"], errors='coerce').fillna(0).astype(int)
        df["Birim Fiyat"] = pd.to_numeric(df["Birim Fiyat"], errors='coerce').fillna(0).astype(float)
        df["Çekmece Numarası"] = df["Çekmece Numarası"].fillna("Belirtilmedi").astype(str).str.strip()
        df["Görsel"] = df["Görsel"].fillna("").astype(str).str.strip()
        df["Tedarikçi Bilgisi"] = df["Tedarikçi Bilgisi"].fillna("Belirtilmedi").astype(str).str.strip()
        df["Toplam Tutar (TL)"] = df["Miktar"] * df["Birim Fiyat"]
        df["Seçim Etiketi"] = df["Ürün Kodu"] + " | " + df["Ürün Adı"].astype(str)
        return df
    else:
        st.error(f"'{EXCEL_FILE}' dosyası bulunamadı!")
        return None

def save_data(df):
    if "Seçim Etiketi" in df.columns: df = df.drop(columns=["Seçim Etiketi"])
    if "Toplam Tutar (TL)" in df.columns: df = df.drop(columns=["Toplam Tutar (TL)"])
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Stok_Listesi')
        worksheet = writer.sheets['Stok_Listesi']
        for i, col in enumerate(df.columns, 1):
            worksheet.column_dimensions[get_column_letter(i)].width = 25

# HATA GİDERİLEN MATEMATİKSEL TAHMİN FONKSİYONU
def calculate_predictions(stock_df):
    stock_df["Aylık Ortalama Tüketim"] = 0.0
    stock_df["Tahmini Stok Ömrü (Gün)"] = "Veri Yetersiz"
    stock_df["Tahmini Stok Ömrü (Gün)"] = stock_df["Tahmini Stok Ömrü (Gün)"].astype(object)
    
    if os.path.exists(LOG_FILE):
        log_df = pd.read_excel(LOG_FILE)
        log_df["Tarih / Saat"] = pd.to_datetime(log_df["Tarih / Saat"])
        
        son_30_gun = datetime.now() - timedelta(days=30)
        cikislar_df = log_df[(log_df["İşlem Türü"] == "Stok Çıkış (-)") & (log_df["Tarih / Saat"] >= son_30_gun)]
        
        if not pd.DataFrame(cikislar_df).empty:
            tuketim_serisi = cikislar_df.groupby("Ürün Kodu")["Miktar"].sum()
            
            for kod in stock_df["Ürün Kodu"].unique():
                if kod in tuketim_serisi.index:
                    toplam_tuketim = tuketim_serisi[kod]
                    stock_df.loc[stock_df["Ürün Kodu"] == kod, "Aylık Ortalama Tüketim"] = float(toplam_tuketim)
                    
                    gunluk_hiz = toplam_tuketim / 30.0
                    mevcut_stok = stock_df.loc[stock_df["Ürün Kodu"] == kod, "Miktar"].values
                    
                    if gunluk_hiz > 0 and len(mevcut_stok) > 0:
                        # DİZİ HATA ÇÖZÜMÜ: mevcut_stok[0] yapılarak scalar değere indirgendi
                        kalan_gun = int(mevcut_stok[0] / gunluk_hiz)
                        idx_match = stock_df[stock_df["Ürün Kodu"] == kod].index
                        stock_df.loc[idx_match, "Tahmini Stok Ömrü (Gün)"] = str(kalan_gun)
                    else:
                        idx_match = stock_df[stock_df["Ürün Kodu"] == kod].index
                        stock_df.loc[idx_match, "Tahmini Stok Ömrü (Gün)"] = "Tüketim Yok"
    return stock_df

df = load_data()
if df is not None:
    df = calculate_predictions(df)
    kritik_df = df[df["Miktar"] < df["Emniyet Faktörü"]]
    st.title("⚡ Elektrik Depo Takip & Yönetim Dashboard")
    st.markdown("---")

    # --- ÜST FİNANSAL KPI KARTLARI ---
    gecerli_fiyatlar = df[df["Birim Fiyat"] > 0]
    en_pahali = gecerli_fiyatlar.loc[gecerli_fiyatlar["Birim Fiyat"].idxmax()] if not gecerli_fiyatlar.empty else None
    en_ucuz = gecerli_fiyatlar.loc[gecerli_fiyatlar["Birim Fiyat"].idxmin()] if not gecerli_fiyatlar.empty else None

    row1_kpi1, row1_kpi2, row1_kpi3 = st.columns(3)
    with row1_kpi1: st.markdown(f'<div class="metric-card"><div class="metric-title">📦 TOPLAM ÜRÜN ÇEŞİDİ</div><div class="metric-value">{len(df)} Adet</div></div>', unsafe_allow_html=True)
    with row1_kpi2: st.markdown(f'<div class="metric-card"><div class="metric-title">🔢 TOPLAM DEPO STOK MİKTARI</div><div class="metric-value">{int(df["Miktar"].sum())} Adet</div></div>', unsafe_allow_html=True)
    with row1_kpi3: st.markdown(f'<div class="metric-card"><div class="metric-title">🚨 KRİTİK STOKTAKİ ÜRÜNLER</div><div class="metric-value" style="color: #ef4444;">{len(kritik_df)} Çeşit</div></div>', unsafe_allow_html=True)

    row2_kpi1, row2_kpi2, row2_kpi3 = st.columns(3)
    with row2_kpi1: st.markdown(f'<div class="metric-card"><div class="metric-title">💵 TOPLAM DEPO DEĞERİ</div><div class="metric-value" style="color: #10b981;">{df["Toplam Tutar (TL)"].sum():,.2f} TL</div></div>', unsafe_allow_html=True)
    with row2_kpi2:
        p_metin = f"{en_pahali['Ürün Adı']} ({en_pahali['Birim Fiyat']:,} TL)" if en_pahali is not None else "Fiyat Belirtilmedi"
        st.markdown(f'<div class="metric-card"><div class="metric-title">📈 EN PAHALI MALZEME</div><div class="metric-value" style="font-size: 15px; padding-top: 8px;">{p_metin}</div></div>', unsafe_allow_html=True)
    with row2_kpi3:
        u_metin = f"{en_ucuz['Ürün Adı']} ({en_ucuz['Birim Fiyat']:,} TL)" if en_ucuz is not None else "Fiyat Belirtilmedi"
        st.markdown(f'<div class="metric-card"><div class="metric-title">📉 EN UCUZ MALZEME</div><div class="metric-value" style="font-size: 15px; padding-top: 8px;">{u_metin}</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    # --- 2. ANA EKRAN İŞLEM VE DETAY ALANI ---
    st.subheader("🛠️ Depo Hareketleri ve İşlemleri")
    islem_sol, islem_sag = st.columns(2)
    
    with islem_sol:
        st.markdown("### 🔍 Ürün Seçimi & Kimlik Kartı")
        secilen_etiket = st.selectbox("Panelden İşlem Yapılacak Ürünü Seçin", df["Seçim Etiketi"].unique(), label_visibility="collapsed")
        secilen_kod = secilen_etiket.split(" | ")[0].strip()
        urun_satiri = df[df["Ürün Kodu"] == secilen_kod].iloc[0]
        
        gorsel_adi = urun_satiri["Görsel"]
        gorsel_bulundu = False; gorsel_yolu = ""
        if gorsel_adi:
            test_yolu = os.path.join(IMAGE_FOLDER, gorsel_adi)
            if os.path.exists(test_yolu): gorsel_yolu = test_yolu; gorsel_bulundu = True
            else:
                for dosya in os.listdir(IMAGE_FOLDER):
                    if dosya.lower() == gorsel_adi.lower(): gorsel_yolu = os.path.join(IMAGE_FOLDER, dosya); gorsel_bulundu = True; break

        if gorsel_bulundu: st.image(gorsel_yolu, width=350)
        else: st.warning("📦 Bu ürünün görseli yok.")
            
        st.markdown(f"""
        <div class="id-card">
            <h4>🆔 ÜRÜN KART BİLGİLERİ</h4>
            <hr style="margin: 8px 0; border-color: #cbd5e1;">
            <b>📋 Kategori:</b> {urun_satiri['Ürün Kategorisi']}<br>
            <b>🏬 Marka:</b> {urun_satiri['Marka']}<br>
            <b>📦 Ürün Adı:</b> {urun_satiri['Ürün Adı']}<br>
            <b>🔑 Ürün Kodu:</b> {urun_satiri['Ürün Kodu']}<br>
            <b>⚠️ Emniyet Faktörü:</b> {urun_satiri['Emniyet Faktörü']} Adet<br>
            <b>💰 Birim Fiyat:</b> {urun_satiri['Birim Fiyat']:,} TL<br>
            <b>🏢 Tedarikçi Bilgisi:</b> {urun_satiri['Tedarikçi Bilgisi']}<br>
            <hr style="margin: 8px 0; border-color: #ff7900; border-style: dashed;">
            <h5 style="color: #ff7900; margin: 4px 0;">🧠 AKILLI GELECEK TAHMİNLERİ (SON 30 GÜN)</h5>
            <b>📊 Aylık Tüketim Hızı:</b> {urun_satiri['Aylık Ortalama Tüketim']} Adet / Ay<br>
            <b>⏳ Tahmini Stok Ömrü:</b> <span style='color: {"#ef4444" if str(urun_satiri["Tahmini Stok Ömrü (Gün)"]).isdigit() and int(urun_satiri["Tahmini Stok Ömrü (Gün)"]) <= 30 else "#10b981"}; font-weight: bold;'>{urun_satiri['Tahmini Stok Ömrü (Gün)']} { "Gün" if str(urun_satiri["Tahmini Stok Ömrü (Gün)"]).isdigit() else "" }</span><br>
            <hr style="margin: 8px 0; border-color: #cbd5e1;">
            <b>📍 Depo Çekmece No:</b> {urun_satiri['Çekmece Numarası']}<br>
            <b>🔢 Mevcut Stok:</b> {urun_satiri['Miktar']} Adet
        </div>
        """, unsafe_allow_html=True)

    with islem_sag:
        st.markdown("### 2️⃣ İşlem Türü Seçin")
        tab_ekle, tab_cikar, tab_konum, tab_duzenle, tab_yeni_urun = st.tabs(["📥 STOK EKLE (+)", "📤 STOK ÇIKAR (-)", "📍 ÇEKMECE DEĞİŞTİR", "📝 ÜRÜN KARTINI DÜZENLE", "🔐 SIFIRDAN YENİ MALZEME KAYDET"])

        with tab_ekle:
            eklenecek_miktar = st.number_input("Eklenecek Miktar", min_value=1, value=1, step=1, key="ekle_input")
            if st.button("🟩 STOKU ARTTIR VE KAYDET", use_container_width=True):
                idx = df[df["Ürün Kodu"] == secilen_kod].index
                df.loc[idx, "Miktar"] += eklenecek_miktar
                save_data(df)
                log_action(secilen_kod, urun_satiri['Ürün Adı'], "Stok Giriş (+)", eklenecek_miktar, f"Eski Stok: {urun_satiri['Miktar']} -> Yeni Stok: {urun_satiri['Miktar'] + eklenecek_miktar}")
                st.toast("📥 Stok başarıyla artırıldı!", icon="✅")
                st.rerun()

        with tab_cikar:
            cikarilacak_miktar = st.number_input("Çıkarılacak Miktar", min_value=1, value=1, step=1, key="cikar_input")
            if st.button("🟥 STOKTAN DÜŞ VE KAYDET", use_container_width=True):
                idx = df[df["Ürün Kodu"] == secilen_kod].index
                yeni_miktar = urun_satiri["Miktar"] - cikarilacak_miktar
                if yeni_miktar >= 0:
                    df.loc[idx, "Miktar"] -= cikarilacak_miktar
                    save_data(df)
                    log_action(secilen_kod, urun_satiri['Ürün Adı'], "Stok Çıkış (-)", cikarilacak_miktar, f"Eski Stok: {urun_satiri['Miktar']} -> Yeni Stok: {yeni_miktar}")
                    st.toast("📤 Stok başarıyla düşürüldü!", icon="✅")
                    st.rerun()
                else: st.toast("⚠️ Hata: Depoda bu kadar ürün yok!", icon="❌")
        with tab_konum:
            k_user = st.text_input("Kullanıcı Adı", key="k_user_input")
            k_pass = st.text_input("Şifre", type="password", key="k_pass_input")
            if k_user == ADMIN_USER and k_pass == ADMIN_PASSWORD:
                yeni_cekmece_no = st.text_input("Yeni Çekmece Numarası", value=urun_satiri['Çekmece Numarası'])
                if st.button("💾 YENİ KONUMU EXCEL'E KAYDET", use_container_width=True):
                    idx = df[df["Ürün Kodu"] == secilen_kod].index
                    df.loc[idx, "Çekmece Numarası"] = yeni_cekmece_no.strip()
                    save_data(df)
                    log_action(secilen_kod, urun_satiri['Ürün Adı'], "Konum Değişimi", 0, f"Eski Konum: {urun_satiri['Çekmece Numarası']} -> Yeni Konum: {yeni_cekmece_no}")
                    st.toast("📍 Çekmece konumu güncellendi!", icon="✅")
                    st.rerun()
            elif k_user or k_pass: st.error("🔒 Hatalı Yetkili Girişi!")

        with tab_duzenle:
            d_user = st.text_input("Kullanıcı Adı", key="d_user_input")
            d_pass = st.text_input("Şifre", type="password", key="d_pass_input")
            if d_user == ADMIN_USER and d_pass == ADMIN_PASSWORD:
                edit_col1, edit_col2 = st.columns(2)
                with edit_col1:
                    e_kat = st.text_input("Ürün Kategorisi", value=urun_satiri["Ürün Kategorisi"])
                    e_marka = st.text_input("Marka", value=urun_satiri["Marka"])
                    e_adi = st.text_input("Ürün Adı", value=urun_satiri["Ürün Adı"])
                with edit_col2:
                    e_emniyet = st.number_input("Emniyet Faktörü", min_value=0, value=int(urun_satiri["Emniyet Faktörü"]))
                    e_fiyat = st.number_input("Birim Fiyat (TL)", min_value=0.0, value=float(urun_satiri["Birim Fiyat"]))
                    e_tedarikci = st.text_input("Tedarikçi Bilgisi", value=urun_satiri["Tedarikçi Bilgisi"])
                if st.button("💾 ÜRÜN KARTINI EXCEL'DE GÜNCELLE", use_container_width=True):
                    idx = df[df["Ürün Kodu"] == secilen_kod].index
                    df.loc[idx, ["Ürün Kategorisi", "Marka", "Ürün Adı", "Emniyet Faktörü", "Birim Fiyat", "Tedarikçi Bilgisi"]] = [e_kat.strip(), e_marka.strip(), e_adi.strip(), int(e_emniyet), float(e_fiyat), e_tedarikci.strip()]
                    save_data(df)
                    log_action(secilen_kod, e_adi, "Ürün Kartı Düzenleme", 0, "Teknik kart detayları güncellendi")
                    st.toast("🎴 Ürün kartı bilgileri güncellendi!", icon="✅")
                    st.rerun()
            elif d_user or d_pass: st.error("🔒 Hatalı Yetkili Girişi!")

        with tab_yeni_urun:
            auth_col1, auth_col2 = st.columns(2)
            with auth_col1: input_user = st.text_input("Kullanıcı Adı", key="auth_user")
            with auth_col2: input_password = st.text_input("Şifre", type="password", key="auth_pass")
            if input_user == ADMIN_USER and input_password == ADMIN_PASSWORD:
                y_col1, y_col2 = st.columns(2)
                with y_col1:
                    y_kat = st.text_input("Ürün Kategorisi"); y_marka = st.text_input("Marka"); y_adi = st.text_input("Ürün Adı"); y_kod = st.text_input("Ürün Kodu")
                with y_col2:
                    y_miktar = st.number_input("Başlangıç Miktar", min_value=0, value=0, step=1); y_emniyet = st.number_input("Emniyet Faktörü", min_value=0, value=5, step=1); y_fiyat = st.number_input("Birim Fiyat (TL)", min_value=0.0, value=0.0, step=1.0); y_tedarik = st.text_input("Tedarikçi Bilgisi"); y_cekmece = st.text_input("Çekmece No"); y_gorsel_dosya = st.file_uploader("Görsel Seçin", type=["jpg", "jpeg", "png"])
                if st.button("💾 YENİ ÜRÜNÜ EXCEL'E EN ALTTA KAYDET", use_container_width=True):
                    if not y_kod or not y_adi: st.error("Hata: Kod ve Ad boş bırakılamaz!")
                    elif y_kod in df["Ürün Kodu"].values: st.error("Hata: Bu kod sistemde var!")
                    else:
                        kaydedilecek_gorsel_adi = ""
                        if y_gorsel_dosya is not None:
                            uzanti = y_gorsel_dosya.name.split(".")[-1]; kaydedilecek_gorsel_adi = f"{y_kod}.{uzanti}"
                            with open(os.path.join(IMAGE_FOLDER, kaydedilecek_gorsel_adi), "wb") as f: f.write(y_gorsel_dosya.getbuffer())
                        yeni_satir = {"Ürün Kategorisi": y_kat, "Marka": y_marka, "Ürün Adı": y_adi, "Ürün Kodu": y_kod, "Miktar": int(y_miktar), "Emniyet Faktörü": int(y_emniyet), "Çekmece Numarası": y_cekmece if y_cekmece else "Belirtilmedi", "Görsel": kaydedilecek_gorsel_adi, "Birim Fiyat": float(y_fiyat), "Tedarikçi Bilgisi": y_tedarik if y_tedarik else "Belirtilmedi"}
                        temp_df = df.drop(columns=["Seçim Etiketi", "Toplam Tutar (TL)", "Aylık Ortalama Tüketim", "Tahmini Stok Ömrü (Gün)"])
                        temp_df = pd.concat([temp_df, pd.DataFrame([yeni_satir])], ignore_index=True)
                        save_data(temp_df)
                        log_action(y_kod, y_adi, "Yeni Ürün Tanımlama", int(y_miktar), f"Kategori: {y_kat}")
                        st.toast(f"✨ {y_adi} en alta kaydedildi!", icon="🚀")
                        st.rerun()

    st.markdown("---")
    # --- DEPO HAREKET KAYITLARI LOG TABLOSU ---
    st.subheader("📜 Depo Geçmişi ve Hareket Raporları")
    if os.path.exists(LOG_FILE):
        history_df = pd.read_excel(LOG_FILE)
        history_df["Tarih / Saat"] = pd.to_datetime(history_df["Tarih / Saat"])
        hist_col1, hist_col2, hist_col3 = st.columns(3)
        with hist_col1: tar_araligi = st.date_input("Tarih Aralığı Seçin", value=[history_df["Tarih / Saat"].min().date(), history_df["Tarih / Saat"].max().date()])
        with hist_col2: tur_filtre = st.multiselect("İşlem Türü Filtresi", history_df["İşlem Türü"].unique())
        with hist_col3: arama_metni = st.text_input("Ürün Adı veya Kodu Ara")
        filtered_hist = history_df.copy()
        
        if isinstance(tar_araligi, list) or isinstance(tar_araligi, tuple):
            if len(tar_araligi) == 2:
                filtered_hist = filtered_hist[(filtered_hist["Tarih / Saat"].dt.date >= tar_araligi[0]) & (filtered_hist["Tarih / Saat"].dt.date <= tar_araligi[1])]
            elif len(tar_araligi) == 1:
                filtered_hist = filtered_hist[filtered_hist["Tarih / Saat"].dt.date == tar_araligi[0]]
                
        if tur_filtre: filtered_hist = filtered_hist[filtered_hist["İşlem Türü"].isin(tur_filtre)]
        if arama_metni: filtered_hist = filtered_hist[(filtered_hist["Ürün Adı"].str.contains(arama_metni, case=False)) | (filtered_hist["Ürün Kodu"].str.contains(arama_metni, case=False))]
        st.dataframe(filtered_hist, use_container_width=True)

    st.markdown("---")
    # --- 3. ANA EKRAN: FİLTRELER VE GRAFİKLER ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🔍 Stok Listesi ve Filtreleme")
        kat_filtresi = st.multiselect("Kategori Filtresi", df["Ürün Kategorisi"].unique())
        marka_filtresi = st.multiselect("Marka Filtresi", df["Marka"].unique())
        filtered_df = df.copy()
        if kat_filtresi: filtered_df = filtered_df[filtered_df["Ürün Kategorisi"].isin(kat_filtresi)]
        if marka_filtresi: filtered_df = filtered_df[filtered_df["Marka"].isin(marka_filtresi)]
        st.dataframe(filtered_df.drop(columns=["Seçim Etiketi"]), use_container_width=True)
        
    with col2:
        st.subheader("📊 EN ÇOK TÜKETİLEN ÜRÜNLER (TÜKETİM HIZ ANALİZİ)")
        tuketim_grafik_df = df[df["Aylık Ortalama Tüketim"] > 0].sort_values(by="Aylık Ortalama Tüketim", ascending=False)
        if not tuketim_grafik_df.empty:
            fig = px.bar(tuketim_grafik_df, x="Ürün Adı", y="Aylık Ortalama Tüketim", color="Ürün Kategorisi", title="Son 30 Günlük Malzeme Tüketim Adetleri")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Grafik için son 30 güne ait herhangi bir 'Stok Çıkış (-)' kaydı bulunamadı. Malzeme çıkışı yapıldıkça bu grafik otomatik beslenecektir.")

    if len(kritik_df) > 0:
        st.markdown("---")
        st.subheader("⚠️ Emniyet Faktörünün Altına Düşen Acil Sipariş Listesi")
        excel_gonder_df = kritik_df[["Ürün Kategorisi", "Marka", "Ürün Adı", "Ürün Kodu", "Miktar", "Emniyet Faktörü", "Çekmece Numarası", "Tedarikçi Bilgisi"]].copy()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as excel_writer:
            excel_gonder_df.to_excel(excel_writer, index=False, sheet_name='Kritik_Siparis_Listesi')
            ws = excel_writer.sheets['Kritik_Siparis_Listesi']
            for col_idx in range(1, len(excel_gonder_df.columns) + 1): ws.column_dimensions[get_column_letter(col_idx)].width = 25
        buffer.seek(0)
        st.download_button(label="📥 KRİTİK STOK LİSTESİNİ EXCEL OLARAK İNDİR (FİYATSIZ)", data=buffer, file_name=f"Kritik_Stok_Siparis_Listesi_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.table(kritik_df[["Ürün Kodu", "Ürün Adı", "Çekmece Numarası", "Miktar", "Emniyet Faktörü", "Tedarikçi Bilgisi", "Tahmini Stok Ömrü (Gün)"]])
