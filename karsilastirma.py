import pandas as pd
import glob
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# Excel dosyalarını yükleme
excel_files = glob.glob('*.xlsx')

# İlk iki Excel dosyasını okuma
first_df = pd.read_excel(excel_files[0])
second_df = pd.read_excel(excel_files[1])

# İlgili kolonları seçme
columns = ["Birim Adı", "Dosya No", "Dosya Durumu", "Dosya Türü"]

# Dosyaları okuma ve ilgili kolonları seçme
try:
    data_first = first_df[columns]
    data_second = second_df[columns]
except KeyError as e:
    print(f"Column not found: {e}")
    data_first, data_second = None, None

# "Dosya No" sütunu üzerinden her iki dataframe'i birleştirme
matched_records = pd.merge(data_first, data_second, on=columns, how='inner')
result = matched_records.dropna(how='all').reset_index(drop=True)

# "Birim Adı" sütunundaki "Cumhuriyet Başsavcılığı" ifadelerini "CBS" ile değiştirme
result["Birim Adı"] = result["Birim Adı"].str.replace("Cumhuriyet Başsavcılığı", "CBS")

# Yeni sütunu ekle ve default değerleri "bilinmiyor" olarak ata
result['Karar Türü'] = 'bilinmiyor'

# İlk dataframe'de "Karar Türü" kontrol etme ve atama
for index, row in result.iterrows():
    matching_row_first = first_df[(first_df['Dosya No'] == row['Dosya No']) & (first_df['Dosya Türü'] == row['Dosya Türü'])]
    if not matching_row_first.empty and not pd.isna(matching_row_first['Karar Türü'].values[0]):
        result.at[index, 'Karar Türü'] = matching_row_first['Karar Türü'].values[0]

# İkinci dataframe'de "Karar Türü" kontrol etme ve atama
for index, row in result.iterrows():
    matching_row_second = second_df[(second_df['Dosya No'] == row['Dosya No']) & (second_df['Dosya Türü'] == row['Dosya Türü'])]
    if not matching_row_second.empty and not pd.isna(matching_row_second['Karar Türü'].values[0]):
        result.at[index, 'Karar Türü'] = matching_row_second['Karar Türü'].values[0]

# "Karar Türü" sütunundaki değerleri düzenleme
def clean_decision_type(value):
    if isinstance(value, str) and value.startswith('[') and value.endswith(']'):
        value = value[1:-1]  # Remove brackets
        value = ", ".join(set(value.split(", ")))  # Split, remove duplicates, and join
        return value.strip()
    else:
        return "bilinmiyor"

result["Karar Türü"] = result["Karar Türü"].apply(clean_decision_type)



#### PDF ÇEVİRMEK için


# .xlsx uzantılarını kaldırıp dosya isimlerini birleştiriyoruz
dosya_adi = "_".join([dosya.replace('.xlsx', '') for dosya in excel_dosyalari])
deger = sonuc[sonuc["Dosya Türü"].isin(["CBS Sorusturma Dosyası", "Ceza Dava Dosyası"])]


# Fontu kaydedelim.
font_name = "Roboto-Light.ttf"  # Yüklediğiniz fontun adını bu şekilde değiştirebilirsiniz.
pdfmetrics.registerFont(TTFont('Roboto-Light', font_name))

# PDF belgesini oluşturma
# file_path = "deneysel.pdf"

# doc = SimpleDocTemplate(file_path, pagesize=landscape(letter))
doc = SimpleDocTemplate(dosya_adi +".pdf", pagesize=landscape(letter)) #yatay
# DataFrame'i bir liste olarak alın ve sütun adlarını ekleyin
data_list = [deger.columns.to_list()] + deger.values.tolist()

# Tabloyu oluşturma
table = Table(data_list, repeatRows=1)
style = TableStyle([
    ('BACKGROUND', (0,0), (-1,0), colors.grey),
    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),

    ('FONTNAME', (0,0), (-1,-1), 'Roboto-Light'),
    ('FONTSIZE', (0,0), (-1,0), 14),

    ('BACKGROUND', (0,1), (-1,-1), colors.beige),
    ('GRID', (0,0), (-1,-1), 1, colors.black),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),

    # Sütun başlıklarını yukarı ve aşağı doğru ayarlama
    ('LINEABOVE', (0,0), (-1,0), 1, colors.black, None, (2,2)),
    ('LINEBELOW', (0,0), (-1,0), 1, colors.black, None, (2,2)),
    ('LEADING', (0, 0), (-1, 0), 20)
])

# Tek ve çift satırlar için arkaplan rengini ayarlama
for i, _ in enumerate(data_list[1:], start=1):
    if i % 2 == 0:
        bg_color = colors.lightgrey
    else:
        bg_color = colors.beige
    style.add('BACKGROUND', (0,i), (-1,i), bg_color)

table.setStyle(style)


# Stil oluşturma
styles = getSampleStyleSheet()
title_style = styles['Title']
title_style.fontName = 'Roboto-Light'

# Bugünün tarihini al
bugun = datetime.today()

# Tarihi belirli bir formatta yazdır
tarih_string = bugun.strftime("%d/%m/%Y")
# Başlığı oluşturma
title = Paragraph(f"{tarih_string} İtibariyle Müşterek CBS Dosyaları", title_style)
# PDF'e tabloyu ekleyin ve dosyayı kaydedin
doc.build([title, table])
