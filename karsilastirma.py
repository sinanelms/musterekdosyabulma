import glob
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer, Image
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch, cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
from itertools import combinations
import locale
import threading
import math
import sys # Font hata ayıklaması için sys
import traceback # Detaylı hata izi için

# --- Yapılandırma ve Sabitler ---

# Türkçe yerel ayarları (locale) belirlemeye çalışıyoruz, eğer başarısız olursa varsayılan ayar kullanılır
try:
    # Önce UTF-8 deneyelim
    locale.setlocale(locale.LC_ALL, 'tr_TR.UTF-8')
except locale.Error:
    try:
        # Eğer UTF-8 desteklenmiyorsa, sadece tr_TR deneyelim
        locale.setlocale(locale.LC_ALL, 'tr_TR')
    except locale.Error:
        print("Uyarı: 'tr_TR.UTF-8' veya 'tr_TR' yerel ayarı bulunamadı. Varsayılan yerel ayar kullanılıyor.")
        # Sistem varsayılanını kullanmak için aşağıdaki satır etkinleştirilebilir:
        # locale.setlocale(locale.LC_ALL, '')


# Font Kayıt İşlemi
FONT_NAME = 'DejaVuSans'
# Font dosyalarının bulunduğu varsayılan konumları kontrol et
# Programın çalıştığı dizin veya fontlar alt dizini olabilir
# Daha robust bir yol: programın kendi .py dosyasının olduğu dizini kullan
program_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
FONT_PATH = os.path.join(program_dir, "DejaVuSans.ttf")
FONT_BOLD_PATH = os.path.join(program_dir, "DejaVuSans-Bold.ttf")

# Eğer program dizininde yoksa, belki mevcut çalışma dizinindedir?
if not os.path.exists(FONT_PATH):
    FONT_PATH = "DejaVuSans.ttf"
if not os.path.exists(FONT_BOLD_PATH):
    FONT_BOLD_PATH = "DejaVuSans-Bold.ttf"


registered_font_name = 'Helvetica' # Varsayılan olarak Helvetica başla
registered_bold_font_name = 'Helvetica-Bold'

try:
    # Normal fontu kaydet
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
        registered_font_name = FONT_NAME
        print(f"Font '{FONT_NAME}' '{FONT_PATH}' yolundan kaydedildi.")
    else:
        print(f"Uyarı: Font dosyası bulunamadı: '{FONT_PATH}'. Varsayılan Helvetica kullanılıyor.")

    # Kalın fontu kaydet
    if os.path.exists(FONT_BOLD_PATH):
        pdfmetrics.registerFont(TTFont(FONT_NAME + '-Bold', FONT_BOLD_PATH))
        # Sadece normal font kaydedildiyse kalın font adına normal font adını ata (Fallback)
        registered_bold_font_name = FONT_NAME + '-Bold' if registered_font_name == FONT_NAME else 'Helvetica-Bold'
        print(f"Kalın Font '{FONT_NAME}-Bold' '{FONT_BOLD_PATH}' yolundan kaydedildi.")

    else:
         print(f"Uyarı: Kalın font dosyası bulunamadı: '{FONT_BOLD_PATH}'. Varsayılan Helvetica-Bold kullanılıyor.")
         registered_bold_font_name = 'Helvetica-Bold' # Varsayılan kalın fontu kullan

except Exception as e:
    print(f"Font kaydı sırasında kritik hata: {e}")
    print("Varsayılan Helvetica ve Helvetica-Bold fontları kullanılacak.")
    registered_font_name = 'Helvetica'
    registered_bold_font_name = 'Helvetica-Bold'


# Excel'den okunacak ve birleştirme için kullanılacak sütunlar
BASE_COLUMNS = ["Birim Adı", "Dosya No", "Dosya Durumu", "Dosya Türü"]

# Raporda kullanılacak geçerli dosya türleri
VALID_DOSYA_TURU = ["Soruşturma Dosyası", "Ceza Dava Dosyası", "CBS İhbar Dosyası"]

# Kısaltma için metin değişim kuralları
REPLACEMENTS = {
    "Birim Adı": {"Cumhuriyet Başsavcılığı": "CBS"},
    "Dosya Türü": {"CBS Sorusturma Dosyası": "Soruşturma Dosyası"} # Bu kural artık gerekli olmayabilir ama dursun
}

# Varsayılan sütun adı değiştirme haritası (GUI üzerinden değiştirilebilir)
DEFAULT_COLUMN_RENAME_MAP = {"Dosya Durumu": "Derdest"}

# Varsayılan Kenar Boşlukları (cm cinsinden)
DEFAULT_MARGIN_CM = 1.5

# --- Stil Fonksiyonları ---

def get_base_styles():
    """Temel ReportLab stillerini alır ve varsayılan fontu ayarlar."""
    styles = getSampleStyleSheet()
    # Kaydedilen font adını kullan
    styles['Title'].fontName = registered_font_name
    styles['Heading1'].fontName = registered_font_name
    styles['Heading2'].fontName = registered_font_name
    styles['Normal'].fontName = registered_font_name
    styles['Italic'].fontName = registered_font_name
    styles['BodyText'].fontName = registered_font_name
    styles['BodyText'].leading = 14  # Satır aralığını artır
    styles['Normal'].leading = 14
    # Başlık stilini biraz küçült
    styles['h1'].fontSize = 16
    styles['h1'].leading = 20
    styles['h3'].fontSize = 10
    styles['h3'].leading = 12

    return styles

def create_table_style(num_rows):
    """Modern görünümlü bir tablo stili oluşturur."""
    # Kaydedilen kalın font adını kullan
    bold_font = registered_bold_font_name

    style = TableStyle([
        # Başlık Stili
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), bold_font),  # Başlık için kalın font
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),

        # Genel Gövde Stili
        ('FONTNAME', (0, 1), (-1, -1), registered_font_name), # Veri için normal font
        ('FONTSIZE', (0, 1), (-1, -1), 9), # Veri font boyutunu biraz küçült
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.darkslategray),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4), # Satır padding'i azalt
        ('TOPPADDING', (0, 1), (-1, -1), 4),

        # Izgara ve Alternatif Satır Renkleri
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.aliceblue, colors.whitesmoke])
    ])
    return style

def calculate_column_widths(dataframe, page_width, min_col_width=1*cm, max_col_width=8*cm, base_char_width=2.5):
    """
    İçeriğe ve sayfa genişliğine göre dinamik sütun genişlikleri hesaplar.
    Args:
        dataframe: Sütun genişlikleri hesaplanacak pandas DataFrame.
        page_width: Sayfada kullanılabilir genişlik (örneğin, doc.width).
        min_col_width: Minimum sütun genişliği.
        max_col_width: Maksimum sütun genişliği.
        base_char_width: Karakter başına tahmini genişlik (font ve boyuta göre ayarlanabilir).
    Returns:
        Sütun genişliklerinin bir listesi (nokta cinsinden).
    """
    widths = []
    total_max_len = 0
    max_lengths = []

    # Her sütun için maksimum uzunluğu hesapla (başlık + veri)
    for column in dataframe.columns:
        header_len = len(str(column))
        try:
            # Veri tipleri farklı olabileceğinden string'e çevirip uzunluğu al
            max_data_len = dataframe[column].astype(str).map(len).max()
            if pd.isna(max_data_len): max_data_len = 0
        except Exception:
            max_data_len = 0
        # Başlık ve veri uzunluğunun maksimumunu al, biraz boşluk ekle (+2)
        current_max = max(header_len, int(max_data_len)) + 2
        max_lengths.append(current_max)
        total_max_len += current_max

    # Maksimum uzunluk oranına göre genişlikleri hesapla, min/max sınırlarını koru
    available_width = page_width
    # Tahmini toplam genişlik (piksel/nokta cinsinden)
    estimated_total_content_width = total_max_len * base_char_width

    # Oranlama faktörü: Mevcut genişliği tahmini içerik genişliğine oranla
    # Eğer tahmini içerik genişliği sıfırsa veya negatifse (olmamalı ama önlem), 1 kullan
    # ÖNCEKİ HATA DÜZELTİLDİ: estimated_total_content_content_width -> estimated_total_content_width
    scale_factor = available_width / estimated_total_content_width if estimated_total_content_width > 0 else 1

    for max_len in max_lengths:
        # Oranlanmış genişliği hesapla
        calculated_width = max_len * base_char_width * scale_factor
        # Minimum ve maksimum genişlik sınırlarını uygula
        final_width = max(min_col_width, min(calculated_width, max_col_width))
        widths.append(final_width)

    # Toplam hesaplanan genişliği kontrol et ve gerekirse ayarlama yap
    # Bu adım, yuvarlamalar veya min/max sınırları nedeniyle toplamın page_width'ten sapmasını düzeltir.
    total_calculated_width = sum(widths)
    if total_calculated_width > 0 and page_width > 0:
        # Tam genişliği kullanmaya çalış
        target_width = page_width * 1.0

        # Eğer toplam hesaplanan genişlik hedef genişlikten farklıysa ayar yap
        if abs(total_calculated_width - target_width) > 1.0: # Küçük sapmaları görmezden gel
             adjustment_factor = target_width / total_calculated_width
             widths = [w * adjustment_factor for w in widths]
             # Ayarlamadan sonra yine min/max sınırlarını kontrol et (önemli!)
             widths = [max(min_col_width, min(w, max_col_width)) for w in widths]

    # print(f"Sayfa Genişliği: {page_width:.2f}, Hesaplanan Toplam Genişlik: {sum(widths):.2f}")

    return widths


# --- Arka Plan/Filigran Fonksiyonu ---

def draw_background(canvas, doc, background_type, background_value):
    """Sayfada filigran metni veya arka plan resmi çizer."""
    if not background_type or background_type == "None":
        return

    canvas.saveState()
    # canvas.setFont('Helvetica', 1) # Burası gereksiz gibi, aşağıdaki setFont kullanılacak

    if background_type == "Watermark Text" and background_value:
        # Filigran için fontu büyük ayarla
        canvas.setFont(registered_font_name, 60)
        canvas.setFillGray(0.85) # Çok açık gri
        # Sayfayı ortala ve döndür
        page_width, page_height = doc.pagesize
        canvas.translate(page_width/2.0, page_height/2.0)
        canvas.rotate(45)
        # Türkçe karakterler için encode etmek gerekebilir, ama TTFont kullandıysak
        # genellikle gerekmez. Yine de problem olursa burada düzenleme yapılabilir.
        try:
             canvas.drawCentredString(0, 0, background_value) # Döndürülmüş ve ortalanmış sayfada (0,0) sayfanın merkezidir
        except Exception as e:
             print(f"Filigran metni çizilirken hata: {e}. ASCII olmayan karakterler olabilir mi?")
             # Hata durumunda en azından bir placeholder çizelim
             canvas.setFont("Helvetica", 30)
             canvas.drawCentredString(0, 0, "Metin Hatası")


    elif background_type == "Background Image" and background_value and os.path.exists(background_value):
        try:
            img_width, img_height = doc.pagesize
            # Resmin kenar boşluklarını hesaba katarak çizim alanını belirle
            drawable_width = img_width - doc.leftMargin - doc.rightMargin
            drawable_height = img_height - doc.topMargin - doc.bottomMargin

            # Resmi çiz
            canvas.drawImage(
                background_value,
                doc.leftMargin, doc.bottomMargin, # Resmin sol alt köşesi
                width=drawable_width, # Yatayda kullanılabilir alan
                height=drawable_height, # Dikeyde kullanılabilir alan
                preserveAspectRatio=True,
                anchor='c' # Merkeze hizala
            )
        except Exception as e:
            print(f"Arka plan resmi çizilirken hata: '{background_value}': {e}")
            canvas.setFillColor(colors.red)
            canvas.setFont("Helvetica", 12)
            # Hata mesajını sayfanın ortasına çiz
            canvas.drawCentredString(doc.pagesize[0]/2, doc.pagesize[1]/2, f"Arka plan resmi yüklenemedi: {os.path.basename(background_value)}")

    canvas.restoreState()

# --- Temel Mantık Fonksiyonları ---

def process_files(file1_path, file2_path, columns_to_use, column_rename_map, log_callback):
    """İki Excel dosyasını okur, birleştirir, temizler, filtreler, sıralar ve sütun adlarını değiştirir."""
    try:
        log_callback(f"{os.path.basename(file1_path)} okunuyor...")
        # Sadece gerekli sütunları oku
        df1 = pd.read_excel(file1_path, usecols=lambda c: c in columns_to_use)
        log_callback(f"{os.path.basename(file2_path)} okunuyor...")
        df2 = pd.read_excel(file2_path, usecols=lambda c: c in columns_to_use)

        # Kontrol: Gerekli tüm sütunlar okundu mu?
        missing_cols_df1 = [col for col in columns_to_use if col not in df1.columns]
        if missing_cols_df1:
            raise KeyError(f"{os.path.basename(file1_path)} dosyasında eksik sütunlar: {', '.join(missing_cols_df1)}")

        missing_cols_df2 = [col for col in columns_to_use if col not in df2.columns]
        if missing_cols_df2:
            raise KeyError(f"{os.path.basename(file2_path)} dosyasında eksik sütunlar: {', '.join(missing_cols_df2)}")


    except FileNotFoundError as e:
        log_callback(f"Hata: Dosya bulunamadı - {e}")
        return None
    except KeyError as e:
        log_callback(f"Hata: Sütun bulunamadı - {e}")
        return None
    except pd.errors.EmptyDataError:
         log_callback(f"Hata: {os.path.basename(file1_path)} veya {os.path.basename(file2_path)} dosyası boş veya okunabilir veri içermiyor.")
         return None
    except Exception as e:
        log_callback(f"Excel dosyaları okunurken hata ({os.path.basename(file1_path)}, {os.path.basename(file2_path)}): {e}")
        log_callback(traceback.format_exc()) # Hatanın detayını logla
        return None

    log_callback("Dosyalar birleştiriliyor...")
    # Verileri birleştir (tüm columns_to_use sütunları üzerinden aynı kayıtları bul)
    # inner join sadece her iki dosyada da ortak olan satırları alır
    # Birleştirme öncesi sütun isimlerinin tam eşleştiğinden emin ol
    try:
         # Sütunların string olduğundan emin ol ve boşlukları kaldır
         df1.columns = df1.columns.astype(str).str.strip()
         df2.columns = df2.columns.astype(str).str.strip()
         merged_df = pd.merge(df1, df2, on=columns_to_use, how='inner')
    except KeyError as e:
         log_callback(f"Hata: Birleştirme sütunları ({', '.join(BASE_COLUMNS)}) dosyalarda eşleşmiyor veya bulunamıyor. Hata: {e}")
         return None
    except Exception as e:
         log_callback(f"Birleştirme sırasında beklenmedik hata: {e}")
         log_callback(traceback.format_exc())
         return None


    # Birleştirmeden sonra aynı satırlar olabilir, tekrar edenleri kaldır
    merged_df = merged_df.drop_duplicates(subset=BASE_COLUMNS).reset_index(drop=True)


    if merged_df.empty:
        log_callback(f"Bilgi: {os.path.basename(file1_path)} ve {os.path.basename(file2_path)} arasında ortak kayıt bulunamadı. Kontrol Edilen Sütunlar: {columns_to_use}")
        return pd.DataFrame() # Boş DataFrame döndür

    log_callback(f"Ortak kayıt sayısı (birleştirme sonrası): {len(merged_df)}")

    # Metin değişimlerini uygula (büyük/küçük harf duyarlılığı olmadan)
    log_callback("Metin değişimleri uygulanıyor...")
    for col, replacements in REPLACEMENTS.items():
        if col in merged_df.columns:
            # NaN değerleri string'e çevirirken 'nan' olmasını önle
            # .loc ile atama yaparak SettingWithCopyWarning'den kaçın
            merged_df.loc[:, col] = merged_df[col].apply(lambda x: str(x) if pd.notna(x) else '').str.strip()
            for old, new in replacements.items():
                 # Regex=False kullanmak özel karakter sorunlarını azaltır
                 # inplace=True kullanmak yerine atama yapıyoruz, pandas uyumluluğu için daha iyi
                 merged_df.loc[:, col] = merged_df[col].str.replace(old, new, case=False, regex=False)


    # Boşlukları temizle (yukarıda değişimler sırasında yapılıyor ama ek kontrol zarar vermez)
    for col in merged_df.columns:
         if merged_df[col].dtype == 'object': # Eğer sütun tipi object (genellikle string) ise
              merged_df.loc[:, col] = merged_df[col].astype(str).str.strip()


    # Dosya Türü'ne göre filtrele
    if "Dosya Türü" in merged_df.columns:
        log_callback(f"'Dosya Türü'ne göre filtreleme uygulanıyor: {VALID_DOSYA_TURU}")
        # Filtreleme sonrası SettingWithCopyWarning almamak için .copy() kullan
        filtered_df = merged_df[merged_df["Dosya Türü"].isin(VALID_DOSYA_TURU)].copy()
    else:
        log_callback("Uyarı: 'Dosya Türü' sütunu bulunamadı, filtreleme atlanıyor.")
        filtered_df = merged_df.copy() # Kopyasını al

    if filtered_df.empty:
        log_callback(f"Bilgi: Birleştirme ve filtreleme sonrası geçerli kayıt bulunamadı.")
        return pd.DataFrame() # Boş DataFrame döndür

    log_callback(f"Filtreleme sonrası ortak kayıt sayısı: {len(filtered_df)}")


    # Sıralama
    log_callback("Veriler sıralanıyor...")
    if 'Dosya No' in filtered_df.columns:
        try:
            # 'Dosya No' sütununu stringe çevir ve böl
            # NaN değerleri boş string olarak ele al
            # .loc ile atama yaparak SettingWithCopyWarning'den kaçın
            split_data = filtered_df['Dosya No'].astype(str).str.split('/', expand=True)

            # Yıl kısmını al (genellikle ilk parça), hataları NaN yap
            filtered_df.loc[:, 'Yıl'] = pd.to_numeric(split_data[0], errors='coerce')

            # Numara kısmını al (genellikle ikinci parça)
            if len(split_data.columns) > 1:
                # Numara kısmındaki '-' gibi karakterleri temizle ve sayıya çevir
                no_part = split_data[1].astype(str).str.replace(r'[^\d]', '', regex=True) # Sadece rakamları bırak
                filtered_df.loc[:, 'No'] = pd.to_numeric(no_part, errors='coerce')
            else:
                filtered_df.loc[:, 'No'] = None # Numara kısmı yoksa None

            # Sıralama için kullanılacak sütunlar
            sort_columns = []
            if 'Birim Adı' in filtered_df.columns:
                # Birim Adı'nda NaN değerleri string olarak ele al ve sırala
                 filtered_df.loc[:, 'Birim Adı_str'] = filtered_df['Birim Adı'].astype(str)
                 sort_columns.append('Birim Adı_str')

            # Yıl ve No sütunlarını sıralama sütunlarına ekle
            sort_columns.extend(['Yıl', 'No'])

            # Sıralama işlemini uygula
            # NaN değerler sıralamada sonda yer alsın
            # inplace=True yerine yeni DataFrame'e atama yap
            filtered_df = filtered_df.sort_values(by=sort_columns, na_position='last').reset_index(drop=True)

            # Sıralama için eklenen yardımcı sütunları kaldır
            filtered_df = filtered_df.drop(['Yıl', 'No'], axis=1, errors='ignore')
            if 'Birim Adı_str' in filtered_df.columns:
                 filtered_df = filtered_df.drop('Birim Adı_str', axis=1)


        except Exception as e:
             log_callback(f"Uyarı: 'Dosya No'ya göre detaylı sıralama başarısız oldu: {e}. Alternatif sıralama uygulanıyor.")
             log_callback(traceback.format_exc())
             if 'Birim Adı' in filtered_df.columns:
                 # Birim Adı'na göre basit sıralama
                 # inplace=True yerine yeni DataFrame'e atama yap
                 filtered_df.loc[:, 'Birim Adı_str'] = filtered_df['Birim Adı'].astype(str)
                 filtered_df = filtered_df.sort_values(by=['Birim Adı_str'], na_position='last').drop('Birim Adı_str', axis=1).reset_index(drop=True)
             else:
                 log_callback("Uyarı: Sıralama yapılamadı.")


    else:
        log_callback("Uyarı: 'Dosya No' sütunu bulunamadı, detaylı sıralama atlanıyor.")
        if 'Birim Adı' in filtered_df.columns:
             # Birim Adı'na göre basit sıralama
             # inplace=True yerine yeni DataFrame'e atama yap
             filtered_df.loc[:, 'Birim Adı_str'] = filtered_df['Birim Adı'].astype(str)
             filtered_df = filtered_df.sort_values(by=['Birim Adı_str'], na_position='last').drop('Birim Adı_str', axis=1).reset_index(drop=True)


    # Sütun adlarını değiştir ve Sıra No ekle
    log_callback("Sütun adları değiştiriliyor ve Sıra No ekleniyor.")
    if column_rename_map:
        # Yalnızca DataFrame'de bulunan sütunları yeniden adlandır
        valid_rename_map = {k: v for k, v in column_rename_map.items() if k in filtered_df.columns}
        if valid_rename_map: # Eğer geçerli yeniden adlandırma varsa uygula
            # inplace=True yerine atama yap
            filtered_df = filtered_df.rename(columns=valid_rename_map)

    final_df = filtered_df.reset_index(drop=True)
    final_df.insert(0, 'Sıra No', range(1, len(final_df) + 1))

    log_callback("Veri işleme tamamlandı.")
    return final_df

def build_pdf_report(output_pdf_path, dataframe, file1_name, file2_name, page_orientation, background_info, margins_cm, log_callback):
    """İşlenmiş verilerle PDF dokümanını oluşturur."""
    styles = get_base_styles()

    # Sayfa yönüne göre sayfa boyutunu ayarla
    page_size = landscape(A4) if page_orientation == "Landscape" else A4

    # Santimetre cinsinden gelen boşlukları ReportLab'in nokta birimine çevir
    left_margin_pt = margins_cm["left"] * cm
    right_margin_pt = margins_cm["right"] * cm
    top_margin_pt = margins_cm["top"] * cm
    bottom_margin_pt = margins_cm["bottom"] * cm

    # Kenar boşlukları sayfa boyutundan büyük olmamalı
    page_width_pt, page_height_pt = page_size
    if left_margin_pt + right_margin_pt >= page_width_pt or top_margin_pt + bottom_margin_pt >= page_height_pt:
        log_callback(f"Hata: Hesaplanan kenar boşlukları sayfa boyutundan büyük veya eşit! Sol+Sağ: {left_margin_pt+right_margin_pt:.2f} vs {page_width_pt:.2f}, Üst+Alt: {top_margin_pt+bottom_margin_pt:.2f} vs {page_height_pt:.2f}")
        # Varsayılan güvenli boşluklara dön veya hata ver
        # Şimdilik hata verip işlemi durdurmak daha güvenli
        root.after(0, lambda: messagebox.showerror("Kenar Boşluğu Hatası", "Belirtilen kenar boşlukları sayfa boyutuna sığmıyor. Lütfen daha küçük değerler girin."))
        return False # PDF oluşturulamadı

    # Doküman şablonunu oluştur
    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=page_size,
        leftMargin=left_margin_pt, rightMargin=right_margin_pt,
        topMargin=top_margin_pt, bottomMargin=bottom_margin_pt,
        title=f"Karşılaştırma - {os.path.basename(file1_name)} vs {os.path.basename(file2_name)}",
        author="Comparison Tool"
    )

    elements = []

    # Başlık
    title_text = f"{datetime.now().strftime('%d/%m/%Y')} Tarihi İtibarıyla Müşterek Dosyalar"
    elements.append(Paragraph(title_text, styles['h1']))
    elements.append(Spacer(1, 0.2*cm))

    # # Alt Başlık/Kaynak Dosyalar
    # subtitle_text = f"(Kaynak Dosyalar: {os.path.basename(file1_name)}.xlsx ve {os.path.basename(file2_name)}.xlsx)"
    # elements.append(Paragraph(subtitle_text, styles['h3']))
    # elements.append(Spacer(1, 0.5*cm))

    # Tablo
    # DataFrame'i list of lists formatına çevir (başlık satırı dahil)
    # NaN değerleri veya None'ları boş string'e çevir
    data_list = [dataframe.columns.to_list()] + [[str(cell) if pd.notna(cell) else "" for cell in row] for row in dataframe.values.tolist()]

    # Tablonun sığabileceği kullanılabilir genişlik
    available_width = doc.width # Bu, pagesize[0] - leftMargin - rightMargin'e eşittir

    # Sütun genişliklerini hesapla
    try:
        column_widths = calculate_column_widths(dataframe, available_width)
    except Exception as e:
        log_callback(f"Hata: Sütun genişlikleri hesaplanırken hata oluştu: {e}")
        log_callback(traceback.format_exc())
        # Hata durumunda varsayılan genişlikler kullanmayı deneyebiliriz veya hata ver
        # Şimdilik hata verip işlemi durdurmak daha güvenli
        root.after(0, lambda: messagebox.showerror("PDF Oluşturma Hatası", f"Sütun genişlikleri hesaplanırken hata oluştu:\n{e}\nPDF oluşturulamadı."))
        return False


    # Tabloyu oluştur
    table = Table(data_list, colWidths=column_widths, repeatRows=1) # repeatRows=1 başlığın her sayfada tekrarlanmasını sağlar
    table.setStyle(create_table_style(len(data_list)))
    elements.append(table)

    # Not Bölümü (Tablodan sonra yeni sayfaya geçilirse not altta kalır, bu genelde istenen durumdur)
    elements.append(Spacer(1, 0.5*cm))
    note_text = f"<b>Not:</b> Bu tablo, <u>{os.path.basename(file1_name)}.xlsx</u> ve <u>{os.path.basename(file2_name)}.xlsx</u> dosyalarında bulunan ortak kayıtları göstermektedir. Karşılaştırma {', '.join(BASE_COLUMNS)} sütunlarına göre yapılmıştır."
    elements.append(Paragraph(note_text, styles['Normal']))

    # PDF Oluştur
    log_callback(f"PDF oluşturuluyor: {os.path.basename(output_pdf_path)}")
    try:
        background_func = None
        if background_info and background_info["type"] != "None":
            # Arka plan fonksiyonunu tanımla, geçerli doc boşlukları ve sayfa boyutu burada kullanılabilir
            def page_background(canvas, doc):
                draw_background(canvas, doc, background_info["type"], background_info["value"])
            background_func = page_background

        if background_func:
            doc.build(elements, onFirstPage=background_func, onLaterPages=background_func)
        else:
            doc.build(elements)

        log_callback(f"BAŞARILI: PDF oluşturuldu -> {output_pdf_path}")

        return True

    except Exception as e:
        log_callback(f"KRİTİK HATA: PDF oluşturulamadı ({os.path.basename(output_pdf_path)}): {e}")
        log_callback(traceback.format_exc()) # Hatanın detayını logla
        # Hata mesajını GUI'de de göster
        root.after(0, lambda: messagebox.showerror("PDF Hatası", f"PDF oluşturulamadı:\n{os.path.basename(output_pdf_path)}\nHata: {e}"))
        return False

# --- Ana Uygulama Sınıfı (GUI) ---

class ComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Dosya Karşılaştırma ve PDF Oluşturma Aracı")
        # Pencere boyutunu biraz küçült çünkü bir satır eksildi
        self.root.geometry("750x700")

        # Stil yapılandırması
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')

        # Değişkenler
        # self.input_folder artık kullanılmıyor
        self.output_folder = tk.StringVar(value=os.getcwd())
        self.page_orientation = tk.StringVar(value="Landscape")
        self.background_type = tk.StringVar(value="None")
        self.background_value = tk.StringVar()
        self.column_rename_map = DEFAULT_COLUMN_RENAME_MAP.copy()

        # Kenar boşluğu değişkenleri (cm cinsinden)
        self.left_margin = tk.DoubleVar(value=DEFAULT_MARGIN_CM)
        self.right_margin = tk.DoubleVar(value=DEFAULT_MARGIN_CM)
        self.top_margin = tk.DoubleVar(value=DEFAULT_MARGIN_CM)
        self.bottom_margin = tk.DoubleVar(value=DEFAULT_MARGIN_CM)


        # Arayüz Çerçeveleri
        # Ayarlar çerçevesini en üste alıyoruz
        options_frame = ttk.LabelFrame(self.root, text="Ayarlar", padding="10")
        options_frame.pack(fill=tk.X, expand=False, padx=10, pady=5) # Ayarlar sabit boyutta kalsın, genişlesin

        # Klasör seçimi ve başlat butonu çerçevesi
        # Input folder seçimi kaldırıldı
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X, expand=False)

        # Durum ve bilgi çerçevesi
        status_frame = ttk.Frame(self.root, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5) # Durum alanı genişlesin

        # --- options_frame içeriği ---
        # Sayfa Yönü
        orient_frame = ttk.Frame(options_frame)
        orient_frame.pack(fill=tk.X, pady=5)
        ttk.Label(orient_frame, text="Sayfa Yönü:").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(orient_frame, text="Yatay (Landscape)", variable=self.page_orientation, value="Landscape").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(orient_frame, text="Dikey (Portrait)", variable=self.page_orientation, value="Portrait").pack(side=tk.LEFT, padx=5)

        # Kenar Boşlukları (Yeni Çerçeve)
        margin_frame = ttk.LabelFrame(options_frame, text="Kenar Boşlukları (cm)", padding="10")
        margin_frame.pack(fill=tk.X, pady=5)

        # Grid kullanarak margin inputlarını düzenleme
        margin_frame.columnconfigure(1, weight=1) # Giriş alanlarının genişlemesini sağla
        margin_frame.columnconfigure(3, weight=1)

        ttk.Label(margin_frame, text="Sol:").grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
        ttk.Entry(margin_frame, textvariable=self.left_margin, width=10).grid(row=0, column=1, sticky=tk.EW, padx=2, pady=2)

        ttk.Label(margin_frame, text="Sağ:").grid(row=0, column=2, sticky=tk.W, padx=2, pady=2)
        ttk.Entry(margin_frame, textvariable=self.right_margin, width=10).grid(row=0, column=3, sticky=tk.EW, padx=2, pady=2)

        ttk.Label(margin_frame, text="Üst:").grid(row=1, column=0, sticky=tk.W, padx=2, pady=2)
        ttk.Entry(margin_frame, textvariable=self.top_margin, width=10).grid(row=1, column=1, sticky=tk.EW, padx=2, pady=2)

        ttk.Label(margin_frame, text="Alt:").grid(row=1, column=2, sticky=tk.W, padx=2, pady=2)
        ttk.Entry(margin_frame, textvariable=self.bottom_margin, width=10).grid(row=1, column=3, sticky=tk.EW, padx=2, pady=2)


        # Arka Plan/Filigran
        bg_frame = ttk.Frame(options_frame)
        bg_frame.pack(fill=tk.X, pady=5)
        ttk.Label(bg_frame, text="Arka Plan:").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(bg_frame, text="Yok", variable=self.background_type, value="None", command=self.update_bg_input_state).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(bg_frame, text="Filigran Yazı", variable=self.background_type, value="Watermark Text", command=self.update_bg_input_state).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(bg_frame, text="Resim Dosyası", variable=self.background_type, value="Background Image", command=self.update_bg_input_state).pack(side=tk.LEFT, padx=2)
        bg_frame.columnconfigure(4, weight=1) # Giriş alanının genişlemesini sağla

        self.bg_entry = ttk.Entry(bg_frame, textvariable=self.background_value, width=30, state=tk.DISABLED)
        self.bg_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True) # Fill ve expand eklendi
        self.bg_button = ttk.Button(bg_frame, text="Seç...", command=self.select_bg_image, state=tk.DISABLED)
        self.bg_button.pack(side=tk.LEFT, padx=5)

        # Sütun Adı Değiştirme
        rename_frame = ttk.Frame(options_frame)
        rename_frame.pack(fill=tk.X, pady=5)
        ttk.Label(rename_frame, text="Sütun Adı Değiştirme:").pack(side=tk.LEFT, padx=5)
        ttk.Button(rename_frame, text="Düzenle", command=self.edit_renames).pack(side=tk.LEFT, padx=5) # Buton metni kısaltıldı
        self.rename_label = ttk.Label(rename_frame, text=f"Aktif: {self.column_rename_map}", foreground="blue") # Rengi mavi yap
        self.rename_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)


        # --- control_frame içeriği ---
        # Giriş Klasörü seçimi kaldırıldı

        ttk.Label(control_frame, text="PDF'lerin Kaydedileceği Klasör:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5) # Satır 0 oldu
        ttk.Entry(control_frame, textvariable=self.output_folder, width=50).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5) # Satır 0 oldu
        ttk.Button(control_frame, text="Gözat...", command=self.select_output_folder).grid(row=0, column=2, padx=5, pady=5) # Satır 0 oldu

        # Eylem Düğmesi - Şimdi Dosya Seçme Diyaloğunu Açacak
        self.run_button = ttk.Button(control_frame, text="Excel Dosyalarını Seç ve Karşılaştırmayı Başlat", command=self.select_files_and_start)
        self.run_button.grid(row=1, column=0, columnspan=3, pady=15) # Satır 1 oldu


        # --- status_frame içeriği ---
        # Bilgi etiketleri (Yeni)
        info_frame = ttk.Frame(status_frame)
        info_frame.pack(fill=tk.X, expand=False, pady=(0, 5)) # Log alanının üstünde, biraz boşlukla

        self.current_pair_label = ttk.Label(info_frame, text="Mevcut Çift: Bekleniyor...", foreground="gray")
        self.current_pair_label.pack(side=tk.LEFT, padx=5)

        self.common_rows_label = ttk.Label(info_frame, text="Ortak Kayıt Sayısı: N/A", foreground="gray")
        self.common_rows_label.pack(side=tk.RIGHT, padx=5)


        # Durum Alanı (Log)
        self.status_text = tk.Text(status_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0", fg="#333333") # Renk ekle
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.config(yscrollcommand=scrollbar.set)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Başlangıçta arka plan giriş alanının durumunu güncelle
        self.update_bg_input_state()


    def log_status(self, message):
        """Durum metin alanına mesaj ekler."""
        # Tkinter GUI update'lerinin main thread'de yapılması gerekir.
        # Eğer bu metot farklı bir thread'den çağrılıyorsa, root.after kullanmalıyız.
        # run_comparison thread'den çağrıldığı için root.after kullanıyoruz.
        self.root.after(0, self._append_status_text, message)

    def _append_status_text(self, message):
          """Durum metin alanına güvenli bir şekilde mesaj ekler (main thread)."""
          self.status_text.config(state=tk.NORMAL)
          timestamp = datetime.now().strftime("%H:%M:%S")
          self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
          self.status_text.see(tk.END)
          self.status_text.config(state=tk.DISABLED)
          self.root.update_idletasks() # Hemen güncellenmesini sağla


    def update_info_labels(self, pair_text, row_count):
          """GUI bilgi etiketlerini günceller (main thread)."""
          self.current_pair_label.config(text=f"Mevcut Çift: {pair_text}", foreground="black" if pair_text != "Bekleniyor..." else "gray")
          self.common_rows_label.config(text=f"Ortak Kayıt Sayısı: {row_count}", foreground="black" if row_count != "N/A" else "gray")


    def select_folder(self, variable):
        """Klasör seçim diyaloğunu açar."""
        # initialdir: Eğer kayıtlı klasör varsa onu kullan, yoksa mevcut çalışma dizinini kullan
        initial_dir = variable.get() if os.path.isdir(variable.get()) else os.getcwd()
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.root)
        if folder_path:
            variable.set(folder_path)

    # select_input_folder kaldırıldı

    def select_output_folder(self):
        """PDF'lerin kaydedileceği klasörü belirler."""
        self.select_folder(self.output_folder)

    def update_bg_input_state(self):
        """Arka plan türü seçimine göre giriş alanını etkinleştirir/devre dışı bırakır."""
        bg_type = self.background_type.get()
        if bg_type == "None":
            self.bg_entry.config(state=tk.DISABLED)
            self.bg_button.config(state=tk.DISABLED)
            self.background_value.set("") # Seçimi sıfırla
        elif bg_type == "Watermark Text":
            self.bg_entry.config(state=tk.NORMAL)
            self.bg_button.config(state=tk.DISABLED)
            # Eğer önceden resim seçilmişse metin alanını temizleme, kalsın.
        elif bg_type == "Background Image":
            self.bg_entry.config(state=tk.NORMAL)
            self.bg_button.config(state=tk.NORMAL)
            # Eğer önceden metin girilmişse resim alanını temizleme, kalsın.


    def select_bg_image(self):
        """Arka plan resmi seçmek için dosya diyaloğunu açar."""
        if self.background_type.get() == "Background Image":
            # initialdir: Eğer kayıtlı resim yolu varsa onun dizinini kullan, yoksa mevcut çalışma dizinini kullan
            initial_dir = os.path.dirname(self.background_value.get()) if self.background_value.get() and os.path.exists(os.path.dirname(self.background_value.get())) else os.getcwd()
            file_path = filedialog.askopenfilename(
                title="Arka Plan Resmini Seç",
                initialdir=initial_dir,
                filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All Files", "*.*")],
                 parent=self.root # Diyalogun ana pencereye bağlı olmasını sağla
            )
            if file_path:
                self.background_value.set(file_path)

    def edit_renames(self):
        """Sütun adlarını düzenlemek için basit bir diyalog açar."""
        # Mevcut yeniden adlandırmaları string formatına getir
        current_map_str = '; '.join([f"{k}->{v}" for k, v in self.column_rename_map.items()])

        new_map_str = simpledialog.askstring(
            "Sütun Adı Değiştir",
            "Sütun adlarını 'EskiAd1->YeniAd1; EskiAd2->YeniAd2' formatında girin:\n"
            "(Örnek: 'Dosya Durumu->Derdest; Birim Adı->CBS Adı')",
            initialvalue=current_map_str,
            parent=self.root # Diyalogun ana pencereye bağlı olmasını sağla
        )

        if new_map_str is not None: # Eğer kullanıcı İptal'e basmadıysa
            try:
                # Kullanıcının girdiği string'i parse et
                temp_map = {}
                if new_map_str.strip(): # Eğer boş değilse parse et
                    pairs = new_map_str.split(';')
                    for pair in pairs:
                        if '->' in pair:
                            old, new = pair.split('->', 1) # Sadece ilk '->'ya göre böl
                            temp_map[old.strip()] = new.strip()
                        # else: Tek '->' içermeyenleri yoksay veya hata verilebilir

                # Geçerli yeniden adlandırma haritasını güncelle
                self.column_rename_map = temp_map
                self.rename_label.config(text=f"Aktif: {self.column_rename_map}")
            except Exception as e:
                messagebox.showerror("Hata", f"Sütun adı değiştirme formatı hatalı. Lütfen 'EskiAd->YeniAd; ...' formatını kullanın.\nHata: {e}")
                # Hata durumunda haritayı son geçerli haline geri döndürmek istenebilir, ama şimdilik simpledialog'un
                # işlevselliği yeterli kabul edildi.

    def select_files_and_start(self):
        """Dosya seçme diyaloğunu açar ve seçilen dosyalarla işlemi başlatır."""
        # Dosya seçme diyaloğu için başlangıç dizini olarak mevcut çalışma dizinini kullan
        initial_dir = os.getcwd()

        # Dosya seçme diyaloğunu aç
        file_paths = filedialog.askopenfilenames(
            title="Karşılaştırılacak Excel Dosyalarını Seçin",
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx")],
            parent=self.root # Diyalogun ana pencereye bağlı olmasını sağla
        )

        if not file_paths: # Kullanıcı hiçbir dosya seçmediyse veya iptal ettiyse
            self.log_status("Dosya seçimi iptal edildi.")
            return

        # Tuple'ı listeye çevir
        file_paths_list = list(file_paths)

        if len(file_paths_list) < 2:
            messagebox.showwarning("Yetersiz Dosya", "Karşılaştırma yapmak için en az iki Excel (.xlsx) dosyası seçmelisiniz.")
            self.log_status("Uyarı: Karşılaştırma için yetersiz dosya seçildi.")
            return

        # Kenar boşluklarını al ve doğrula
        try:
            margins_cm = {
                "left": self.left_margin.get(),
                "right": self.right_margin.get(),
                "top": self.top_margin.get(),
                "bottom": self.bottom_margin.get()
            }
            # Negatif veya sıfır boşluk kontrolü
            if any(m < 0 for m in margins_cm.values()):
                messagebox.showerror("Hata", "Kenar boşlukları negatif olamaz.")
                self.log_status("Hata: Kenar boşlukları negatif olamaz.")
                return
             # Sayfa boyutu kontrolü build_pdf_report fonksiyonuna taşındı

        except tk.TclError:
            messagebox.showerror("Hata", "Kenar boşluğu değerleri sayı olmalıdır.")
            self.log_status("Hata: Kenar boşluğu değerleri sayı olmalıdır.")
            return

        out_folder = self.output_folder.get()
        if not os.path.isdir(out_folder):
            try:
                os.makedirs(out_folder)
                self.log_status(f"Çıkış klasörü oluşturuldu: {out_folder}")
            except Exception as e:
                messagebox.showerror("Hata", f"Çıkış klasörü oluşturulamadı:\n{out_folder}\n{e}")
                self.log_status(f"Hata: Çıkış klasörü oluşturulamadı: {e}")
                return


        # GUI elementlerini devre dışı bırak
        self.run_button.config(state=tk.DISABLED, text="İşlem Başlatılıyor...")
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete('1.0', tk.END) # Önceki logları temizle
        self.status_text.config(state=tk.DISABLED)
        self.update_info_labels("Hazırlanıyor...", "N/A") # Bilgi etiketlerini sıfırla
        self.log_status(f"Seçilen {len(file_paths_list)} dosya ile işlem başlatılıyor...")

        bg_info = {
            "type": self.background_type.get(),
            "value": self.background_value.get() if self.background_type.get() != "None" else None
        }

        # Arka plan resmi seçildiyse dosyanın varlığını kontrol et
        if bg_info["type"] == "Background Image" and bg_info["value"] and not os.path.exists(bg_info["value"]):
             messagebox.showerror("Hata", f"Arka plan resim dosyası bulunamadı:\n{bg_info['value']}")
             self.log_status(f"Hata: Arka plan resim dosyası bulunamadı: {bg_info['value']}")
             self.enable_run_button() # Düğmeyi tekrar etkinleştir
             return


        # İşlemi ayrı bir iş parçacığında başlat
        thread = threading.Thread(
            target=self.run_comparison,
            args=(file_paths_list, out_folder, self.page_orientation.get(), bg_info, self.column_rename_map.copy(), margins_cm),
            daemon=True # Ana uygulama kapanınca thread de kapansın
        )
        thread.start()


    def run_comparison(self, file_paths, output_dir, page_orientation, background_info, column_map, margins_cm):
        """Seçilen dosya çiftlerini karşılaştırır ve PDF oluşturur."""
        try:
            # Seçilen dosyalardan tüm çiftleri oluştur
            pairings = list(combinations(file_paths, 2))
            self.log_status(f"Toplam {len(pairings)} dosya çifti karşılaştırılacak.")

            success_count = 0
            fail_count = 0
            skipped_count = 0 # Ortak kayıt bulunamadığı için atlananlar

            for i, (file1_path, file2_path) in enumerate(pairings):
                file1_name = os.path.splitext(os.path.basename(file1_path))[0]
                file2_name = os.path.splitext(os.path.basename(file2_path))[0]
                pdf_name = f"{file1_name}_vs_{file2_name}_Comparison.pdf"
                output_pdf_path = os.path.join(output_dir, pdf_name)

                current_pair_text = f"{os.path.basename(file1_path)} vs {os.path.basename(file2_path)} ({i+1}/{len(pairings)})"
                self.log_status(f"--- Karşılaştırma ({i+1}/{len(pairings)}): {os.path.basename(file1_path)} vs {os.path.basename(file2_path)} ---")
                self.root.after(0, self.update_info_labels, current_pair_text, "Hesaplanıyor...") # GUI'yi güncelle

                processed_data = process_files(file1_path, file2_path, BASE_COLUMNS, column_map, self.log_status) # log_callback'i pass et

                if processed_data is not None and not processed_data.empty:
                    common_rows_count = len(processed_data)
                    self.root.after(0, self.update_info_labels, current_pair_text, common_rows_count) # Ortak kayıt sayısını GUI'ye yaz
                    self.log_status(f"Ortak kayıt bulundu: {common_rows_count}. PDF oluşturuluyor: {pdf_name}")

                    pdf_success = build_pdf_report(
                        output_pdf_path,
                        processed_data,
                        file1_name,
                        file2_name,
                        page_orientation,
                        background_info,
                        margins_cm,
                        self.log_status # log_callback'i pass et
                    )
                    if pdf_success:
                        success_count += 1
                    else:
                        fail_count += 1
                        # Hata build_pdf_report içinde loglanıyor
                elif processed_data is None:
                    # process_files bir hata nedeniyle None döndürdüyse
                    fail_count += 1
                    # Hata process_files içinde loglanıyor
                    self.root.after(0, self.update_info_labels, current_pair_text, "Hata Oluştu") # GUI'yi güncelle
                else: # processed_data boş DataFrame ise (ortak kayıt bulunamadı veya filtre sonrası boş)
                    self.log_status(f"Bilgi: {os.path.basename(file1_path)} ve {os.path.basename(file2_path)} arasında ortak kayıt bulunamadı veya geçerli kayıt kalmadı. PDF oluşturulmuyor.")
                    self.root.after(0, self.update_info_labels, current_pair_text, "0 (Atlandı)") # GUI'yi güncelle
                    skipped_count += 1


            self.log_status("--- İşlem Tamamlandı ---")
            self.log_status(f"Başarılı Oluşturulan PDF: {success_count}")
            self.log_status(f"Ortak Kayıt Bulunamayan/Atlanan Çift: {skipped_count}")
            self.log_status(f"Hata Oluşan Çift: {fail_count}")


            final_message = (
                f"İşlem tamamlandı.\n\n"
                f"Başarıyla oluşturulan PDF sayısı: {success_count}\n"
                f"Ortak kayıt bulunamayan/atlanan çift sayısı: {skipped_count}\n"
                f"Hata oluşan dosya çifti sayısı: {fail_count}"
            )
            self.root.after(0, lambda: messagebox.showinfo("İşlem Tamamlandı", final_message))


        except Exception as e:
            self.log_status(f"BEKLENMEDİK KRİTİK HATA OLUŞTU: {e}")
            self.log_status(traceback.format_exc()) # Hatanın detayını logla
            self.root.after(0, lambda: messagebox.showerror("Kritik Hata", f"İşlem sırasında beklenmedik bir hata oluştu:\n{e}"))
        finally:
            # İşlem bitince düğmeyi tekrar etkinleştir ve GUI'yi temizle
            self.root.after(0, self.enable_run_button)
            self.root.after(0, self.update_info_labels, "Bekleniyor...", "N/A") # Bilgi etiketlerini sıfırla


    def enable_run_button(self):
        """Ana iş parçacığından çalıştırma düğmesini güvenli bir şekilde yeniden etkinleştirir."""
        self.run_button.config(state=tk.NORMAL, text="Excel Dosyalarını Seç ve Karşılaştırmayı Başlat")


# --- Ana Çalıştırma ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ComparisonApp(root)
    root.mainloop()
