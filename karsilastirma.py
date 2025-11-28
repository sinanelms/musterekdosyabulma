import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime
import io
import traceback
import numpy as np
import os
import tempfile
import subprocess
import platform

# --- REPORTLAB IMPORTLARI ---
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, portrait, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# --- PANDAS AYARLARI ---
pd.set_option('future.no_silent_downcasting', True)

# --- FONT AYARLARI ---
FONT_NAME = 'DejaVuSans'
FONT_BOLD_NAME = 'DejaVuSans-Bold'

try:
    if os.path.exists("DejaVuSans.ttf"):
        pdfmetrics.registerFont(TTFont(FONT_NAME, "DejaVuSans.ttf"))
        font_regular = FONT_NAME
    else:
        font_regular = "Helvetica"

    if os.path.exists("DejaVuSans-Bold.ttf"):
        pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, "DejaVuSans-Bold.ttf"))
        font_bold = FONT_BOLD_NAME
    else:
        font_bold = "Helvetica-Bold"
except Exception:
    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"

# --- SABİTLER ---

# GÖRSELDEKİ SABİT SÜTUN İSİMLERİ (YENİ EKLENDİ)
FIXED_HEADERS = [
    "Birim Adı", "Dosya Durumu", "Dosya Türü", "Dosya No", "Sıfatı", "Vekilleri",
    "Dava Türleri", "Dava Konusu", "İlamat Numaraları", "Suçu", "Suç Tarihi",
    "Karar Türü", "Kesinleşme Tarihi", "Kesinleşme Türü", "Açıklama"
]

BASE_COLUMNS = ["Birim Adı", "Dosya No", "Dosya Durumu", "Dosya Türü"]
MERGE_FIX_COLUMNS = ["Birim Adı", "Dosya Durumu", "Dosya Türü", "Dosya No", "Sıfatı", "Vekilleri"]
VALID_DOSYA_TURU = ["Soruşturma Dosyası", "Ceza Dava Dosyası", "CBS İhbar Dosyası"]

REPLACEMENTS = {
    "Birim Adı": {"Cumhuriyet Başsavcılığı": "CBS"},
    "Dosya Türü": {"CBS Sorusturma Dosyası": "Soruşturma Dosyası"}
}

# --- GÖRSEL PDF EDİTÖRÜ VE ÖNİZLEME PENCERESİ ---

class PDFLayoutEditor:
    def __init__(self, parent, dataframe, callback_save):
        self.top = tk.Toplevel(parent)
        self.top.title("PDF Düzenleme ve Önizleme")
        self.top.geometry("1100x700")
        self.df = dataframe
        self.callback_save = callback_save 
        
        self.orientation_var = tk.StringVar(value="Landscape")
        self.margin_var = tk.DoubleVar(value=1.0)
        self.col_weights = {} 
        
        self.paned = ttk.PanedWindow(self.top, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.settings_frame = ttk.LabelFrame(self.paned, text="Ayarlar", padding=10)
        self.paned.add(self.settings_frame, weight=1)
        
        self.preview_frame = ttk.LabelFrame(self.paned, text="Şematik Önizleme (Sayfa Yerleşimi)", padding=10)
        self.paned.add(self.preview_frame, weight=3)
        
        self.setup_settings_ui()
        self.setup_preview_ui()
        self.calculate_initial_weights()
        self.draw_preview()

    def setup_settings_ui(self):
        ttk.Label(self.settings_frame, text="Sayfa Yönü:", font="bold").pack(anchor="w", pady=(0, 5))
        ttk.Radiobutton(self.settings_frame, text="Yatay (Landscape)", variable=self.orientation_var, value="Landscape", command=self.draw_preview).pack(anchor="w")
        ttk.Radiobutton(self.settings_frame, text="Dikey (Portrait)", variable=self.orientation_var, value="Portrait", command=self.draw_preview).pack(anchor="w")
        
        ttk.Label(self.settings_frame, text="Kenar Boşluğu (cm):", font="bold").pack(anchor="w", pady=(15, 5))
        scale_margin = ttk.Scale(self.settings_frame, from_=0.5, to=3.0, variable=self.margin_var, command=lambda x: self.draw_preview())
        scale_margin.pack(fill=tk.X)
        
        ttk.Label(self.settings_frame, text="Sütun Genişlik Ayarları:", font="bold").pack(anchor="w", pady=(20, 5))
        ttk.Label(self.settings_frame, text="(Sütunların kaplayacağı alanı ayarlayın)", font=("Arial", 8)).pack(anchor="w")

        canvas_scroll = tk.Canvas(self.settings_frame, height=300)
        scrollbar = ttk.Scrollbar(self.settings_frame, orient="vertical", command=canvas_scroll.yview)
        self.sliders_frame = ttk.Frame(canvas_scroll)
        
        self.sliders_frame.bind("<Configure>", lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all")))
        canvas_scroll.create_window((0, 0), window=self.sliders_frame, anchor="nw")
        canvas_scroll.configure(yscrollcommand=scrollbar.set)
        
        canvas_scroll.pack(side="top", fill="both", expand=True, pady=5)
        scrollbar.pack(side="right", fill="y")
        
        btn_frame = ttk.Frame(self.settings_frame)
        btn_frame.pack(side="bottom", fill="x", pady=10)
        
        ttk.Button(btn_frame, text="👁️ Gerçek PDF Önizle", command=self.generate_temp_preview).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="💾 PDF Olarak Kaydet", command=self.save_final).pack(fill=tk.X, pady=(10, 2))

    def setup_preview_ui(self):
        self.canvas = tk.Canvas(self.preview_frame, bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", lambda event: self.draw_preview())

    def calculate_initial_weights(self):
        self.sliders = {}
        for col in self.df.columns:
            max_len = len(str(col))
            data_len = self.df[col].astype(str).map(len).head(50).max()
            if pd.isna(data_len): data_len = 0
            weight = max(max_len, data_len, 5)
            self.col_weights[col] = tk.DoubleVar(value=weight)
            f = ttk.Frame(self.sliders_frame)
            f.pack(fill=tk.X, pady=2)
            ttk.Label(f, text=col[:20], width=15, anchor="w").pack(side=tk.LEFT)
            s = ttk.Scale(f, from_=1, to=100, variable=self.col_weights[col], command=lambda x: self.draw_preview())
            s.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def draw_preview(self):
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w < 50: return
        
        if self.orientation_var.get() == "Landscape":
            ratio = 29.7 / 21.0
        else:
            ratio = 21.0 / 29.7
            
        paper_h = h - 40
        paper_w = paper_h * ratio
        
        if paper_w > w - 40:
            paper_w = w - 40
            paper_h = paper_w / ratio
            
        x_start = (w - paper_w) / 2
        y_start = (h - paper_h) / 2
        
        self.canvas.create_rectangle(x_start, y_start, x_start + paper_w, y_start + paper_h, fill="white", outline="black", width=2)
        
        margin_cm = self.margin_var.get()
        page_width_cm = 29.7 if self.orientation_var.get() == "Landscape" else 21.0
        px_per_cm = paper_w / page_width_cm
        margin_px = margin_cm * px_per_cm
        
        draw_x = x_start + margin_px
        draw_y = y_start + margin_px
        draw_w = paper_w - (2 * margin_px)
        draw_h = paper_h - (2 * margin_px)
        
        self.canvas.create_rectangle(draw_x, draw_y, draw_x + draw_w, draw_y + draw_h, outline="red", dash=(2, 4))
        
        total_weight = sum(v.get() for v in self.col_weights.values())
        if total_weight == 0: total_weight = 1
        
        current_x = draw_x
        colors_cycle = ["#e6f3ff", "#fff0e6", "#e6ffe6", "#fffde6"]
        
        for i, col in enumerate(self.df.columns):
            weight = self.col_weights[col].get()
            col_px = (weight / total_weight) * draw_w
            self.canvas.create_rectangle(current_x, draw_y, current_x + col_px, draw_y + draw_h, fill=colors_cycle[i%4], outline="gray")
            if col_px > 20:
                self.canvas.create_text(current_x + col_px/2, draw_y + 15, text=col[:10], font=("Arial", 7), angle=90)
            current_x += col_px

    def get_column_widths_cm(self, page_width_cm):
        total_weight = sum(v.get() for v in self.col_weights.values())
        if total_weight == 0: total_weight = 1
        widths = []
        for col in self.df.columns:
            w = (self.col_weights[col].get() / total_weight) * page_width_cm
            widths.append(w * cm)
        return widths

    def create_pdf_data(self, output_path):
        margin = self.margin_var.get()
        orientation = self.orientation_var.get()
        page_size = landscape(A4) if orientation == "Landscape" else A4
        page_w_pt, page_h_pt = page_size
        margin_pt = margin * cm
        printable_width_cm = (page_w_pt / cm) - (2 * margin)
        col_widths = self.get_column_widths_cm(printable_width_cm)
        
        doc = SimpleDocTemplate(output_path, pagesize=page_size, leftMargin=margin_pt, rightMargin=margin_pt, topMargin=margin_pt, bottomMargin=margin_pt)
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontName=font_bold, alignment=1, spaceAfter=10)
        elements.append(Paragraph(f"Karşılaştırma Raporu - {datetime.now().strftime('%d.%m.%Y')}", title_style))
        elements.append(Spacer(1, 0.5 * cm))
        
        cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=font_regular, fontSize=8, leading=10, alignment=TA_LEFT)
        header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontName=font_bold, fontSize=9, textColor=colors.whitesmoke, alignment=TA_CENTER)
        
        data = []
        headers = [Paragraph(col, header_style) for col in self.df.columns]
        data.append(headers)
        
        for row in self.df.values:
            row_data = []
            for item in row:
                text = str(item) if pd.notna(item) else ""
                row_data.append(Paragraph(text, cell_style))
            data.append(row_data)
            
        table = Table(data, colWidths=col_widths, repeatRows=1)
        tbl_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.aliceblue, colors.whitesmoke]),
            ('LEFTPADDING', (0,0), (-1,-1), 3), ('RIGHTPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ])
        table.setStyle(tbl_style)
        elements.append(table)
        elements.append(Spacer(1, 0.5 * cm))
        elements.append(Paragraph(f"Toplam Kayıt Sayısı: {len(self.df)}", styles['Normal']))
        
        try:
            doc.build(elements)
            return True, ""
        except Exception as e:
            return False, str(e)

    def generate_temp_preview(self):
        try:
            fd, temp_path = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            success, msg = self.create_pdf_data(temp_path)
            if success:
                if platform.system() == 'Windows': os.startfile(temp_path)
                elif platform.system() == 'Darwin': subprocess.call(('open', temp_path))
                else: subprocess.call(('xdg-open', temp_path))
            else: messagebox.showerror("Hata", f"Önizleme oluşturulamadı: {msg}")
        except Exception as e: messagebox.showerror("Hata", f"Önizleme hatası: {e}")

    def save_final(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Rapor_{datetime.now().strftime('%Y%m%d')}.pdf", filetypes=[("PDF Dosyası", "*.pdf")], title="PDF Kaydet")
        if file_path:
            success, msg = self.create_pdf_data(file_path)
            if success:
                messagebox.showinfo("Başarılı", "PDF dosyası kaydedildi.")
                try: os.startfile(file_path)
                except: pass
                self.top.destroy()
            else: messagebox.showerror("Hata", f"Kaydedilemedi: {msg}")

# --- VERİ İŞLEME FONKSİYONLARI ---

def parse_clipboard_data(clipboard_text, log_callback):
    """
    Panodaki veriyi okur. Sütun isimleri FIXED_HEADERS'dan alınır.
    header=None yapılarak ilk satırın veri olması sağlanır.
    """
    try:
        if not clipboard_text or clipboard_text.strip() == "":
            log_callback("Hata: Yapıştırılan veri boş.", "ERROR")
            return None
        
        # header=None: Verinin içinde başlık satırı yok kabul et
        # names=FIXED_HEADERS: Başlıkları biz zorla atıyoruz
        df = pd.read_csv(io.StringIO(clipboard_text), sep='\t', engine='python', dtype=str, header=None, names=FIXED_HEADERS)
        
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        df = df.replace(r'^\s*$', np.nan, regex=True).infer_objects(copy=False)
        
        cols_to_fill = [col for col in MERGE_FIX_COLUMNS if col in df.columns]
        if cols_to_fill:
            df[cols_to_fill] = df[cols_to_fill].ffill()
        
        df = df.fillna("")
        log_callback(f"Veri parça olarak işlendi: {len(df)} satır.", "INFO")
        return df
        
    except Exception as e:
        log_callback(f"Veri işlenirken hata: {e}", "ERROR")
        log_callback(f"Detay: {traceback.format_exc()}", "DEBUG")
        return None

def process_comparison(df1, df2, columns_to_use, log_callback):
    try:
        missing_cols_df1 = [col for col in columns_to_use if col not in df1.columns]
        if missing_cols_df1:
            log_callback(f"İlk veri setinde eksik sütunlar: {', '.join(missing_cols_df1)}", "ERROR")
            return None
            
        missing_cols_df2 = [col for col in columns_to_use if col not in df2.columns]
        if missing_cols_df2:
            log_callback(f"İkinci veri setinde eksik sütunlar: {', '.join(missing_cols_df2)}", "ERROR")
            return None
        
        log_callback("Veriler birleştiriliyor...", "INFO")
        
        for col in columns_to_use:
            df1[col] = df1[col].astype(str).str.strip()
            df2[col] = df2[col].astype(str).str.strip()
        
        merged_df = pd.merge(df1, df2, on=columns_to_use, how='inner')
        merged_df = merged_df.drop_duplicates(subset=columns_to_use).reset_index(drop=True)
        
        if merged_df.empty:
            log_callback("Bilgi: Ortak kayıt bulunamadı.", "INFO")
            return pd.DataFrame()
        
        for col, replacements_map in REPLACEMENTS.items():
            if col in merged_df.columns:
                for old, new in replacements_map.items():
                    merged_df[col] = merged_df[col].str.replace(old, new, case=False, regex=False)
        
        if "Dosya Türü" in merged_df.columns:
            filtered_df = merged_df[merged_df["Dosya Türü"].isin(VALID_DOSYA_TURU)]
        else:
            filtered_df = merged_df
        
        if filtered_df.empty:
            log_callback("Bilgi: Filtreleme sonrası geçerli kayıt bulunamadı.", "INFO")
            return pd.DataFrame()
            
        if 'Dosya No' in filtered_df.columns:
            try:
                temp_df = filtered_df.copy()
                split_data = temp_df['Dosya No'].astype(str).str.split('/', n=1, expand=True)
                temp_df['_Yil'] = pd.to_numeric(split_data[0].str.strip(), errors='coerce')
                
                if split_data.shape[1] > 1:
                    temp_df['_No'] = pd.to_numeric(split_data[1].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce')
                else:
                    temp_df['_No'] = 0
                
                sort_cols = ['_Yil', '_No']
                if 'Birim Adı' in temp_df.columns:
                    sort_cols.insert(0, 'Birim Adı')
                    
                temp_df = temp_df.sort_values(by=sort_cols, na_position='last')
                filtered_df = temp_df.drop(columns=['_Yil', '_No'], errors='ignore')
            except Exception as e:
                log_callback(f"Sıralama uyarısı: {e}", "WARN")
        
        final_df = filtered_df.copy()
        final_df.insert(0, 'Sıra No', range(1, len(final_df) + 1))
        return final_df
        
    except Exception as e:
        log_callback(f"Karşılaştırma sırasında hata: {e}", "ERROR")
        log_callback(f"Detay: {traceback.format_exc()}", "DEBUG")
        return None

# --- SÜTUN SEÇİCİ PENCERESİ ---

class ColumnSelectorDialog:
    def __init__(self, parent, all_columns, currently_selected, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Görünümü Özelleştir (Analist Modu)")
        self.top.geometry("500x600")
        self.callback = callback
        self.all_columns = all_columns
        self.vars = {}
        
        for col in all_columns:
            is_selected = (col in currently_selected) if currently_selected is not None else True
            self.vars[col] = tk.BooleanVar(value=is_selected)
        
        lbl = ttk.Label(self.top, text="Analiz etmek istediğiniz sütunları seçin:", font=('Arial', 10, 'bold'))
        lbl.pack(pady=10)
        
        filter_frame = ttk.Frame(self.top)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(filter_frame, text="Filtrele:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_list)
        entry = ttk.Entry(filter_frame, textvariable=self.search_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        frame_container = ttk.Frame(self.top)
        frame_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(frame_container)
        scrollbar = ttk.Scrollbar(frame_container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.create_checkbuttons()

        btn_frame = ttk.Frame(self.top)
        btn_frame.pack(fill=tk.X, pady=10, padx=10)
        
        ttk.Button(btn_frame, text="Tümünü Seç", command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Tümünü Kaldır", command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="UYGULA ve RAPORLA", command=self.apply_selection).pack(side=tk.RIGHT, padx=5)

    def create_checkbuttons(self, filter_text=""):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        filter_text = filter_text.lower()
        for col in self.all_columns:
            if filter_text and filter_text not in col.lower(): continue
            cb = ttk.Checkbutton(self.scrollable_frame, text=col, variable=self.vars[col])
            cb.pack(anchor='w', pady=2)

    def filter_list(self, *args): self.create_checkbuttons(self.search_var.get())
    def select_all(self):
        for col in self.vars: self.vars[col].set(True)
    def deselect_all(self):
        for col in self.vars: self.vars[col].set(False)
    def apply_selection(self):
        selected = [col for col in self.all_columns if self.vars[col].get()]
        if not selected:
            messagebox.showwarning("Uyarı", "En az bir sütun seçmelisiniz.")
            return
        self.callback(selected)
        self.top.destroy()

# --- ANA UYGULAMA ---

class PasteComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Veri Karşılaştırma ve Gelişmiş PDF Aracı")
        self.root.geometry("1400x850")
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')
        self.hide_empty_cols_var = tk.BooleanVar(value=False)
        
        main_container = ttk.Frame(self.root, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="Excel Dosya Karşılaştırma ve Raporlama Aracı", font=('Arial', 14, 'bold')).pack()
        
        self.paned_window = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Sol Panel
        self.left_frame = ttk.LabelFrame(self.paned_window, text="📋 İlk Excel Verisi", padding="5")
        self.paned_window.add(self.left_frame, weight=1)
        name_frame1 = ttk.Frame(self.left_frame)
        name_frame1.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(name_frame1, text="Dosya Adı:").pack(side=tk.LEFT)
        self.name_entry1 = ttk.Entry(name_frame1)
        self.name_entry1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.name_entry1.insert(0, "Excel_1.xlsx")
        tree_cont1 = ttk.Frame(self.left_frame)
        tree_cont1.pack(fill=tk.BOTH, expand=True)
        self.tree1 = self.create_treeview(tree_cont1)
        btn_frame1 = ttk.Frame(self.left_frame)
        btn_frame1.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame1, text="Yapıştır (Ctrl+V)", command=lambda: self.paste_data(1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame1, text="Temizle", command=lambda: self.clear_tree(1)).pack(side=tk.LEFT, padx=2)
        self.count_label1 = ttk.Label(btn_frame1, text="Satır: 0", foreground='blue')
        self.count_label1.pack(side=tk.RIGHT)

        # Sağ Panel
        self.right_frame = ttk.LabelFrame(self.paned_window, text="📋 İkinci Excel Verisi", padding="5")
        self.paned_window.add(self.right_frame, weight=1)
        name_frame2 = ttk.Frame(self.right_frame)
        name_frame2.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(name_frame2, text="Dosya Adı:").pack(side=tk.LEFT)
        self.name_entry2 = ttk.Entry(name_frame2)
        self.name_entry2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.name_entry2.insert(0, "Excel_2.xlsx")
        tree_cont2 = ttk.Frame(self.right_frame)
        tree_cont2.pack(fill=tk.BOTH, expand=True)
        self.tree2 = self.create_treeview(tree_cont2)
        btn_frame2 = ttk.Frame(self.right_frame)
        btn_frame2.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame2, text="Yapıştır (Ctrl+V)", command=lambda: self.paste_data(2)).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame2, text="Temizle", command=lambda: self.clear_tree(2)).pack(side=tk.LEFT, padx=2)
        self.count_label2 = ttk.Label(btn_frame2, text="Satır: 0", foreground='blue')
        self.count_label2.pack(side=tk.RIGHT)
        
        control_frame = ttk.Frame(main_container)
        control_frame.pack(fill=tk.X, pady=5)
        ttk.Button(control_frame, text="🔍 Karşılaştır", command=self.compare_data).pack(side=tk.LEFT, padx=5)
        self.btn_customize = ttk.Button(control_frame, text="🛠️ Sütunları Seç", command=self.open_column_selector, state=tk.DISABLED)
        self.btn_customize.pack(side=tk.LEFT, padx=5)
        self.btn_pdf = ttk.Button(control_frame, text="📄 PDF Önizle ve Kaydet", command=self.open_pdf_editor, state=tk.DISABLED)
        self.btn_pdf.pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="📋 Excel/Kopyala", command=self.copy_result_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="🗑️ Tümünü Temizle", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Separator(control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Checkbutton(control_frame, text="Boş Sütunları Gizle", variable=self.hide_empty_cols_var, command=self.refresh_all_views).pack(side=tk.LEFT, padx=5)
        
        result_frame = ttk.LabelFrame(main_container, text="📊 Karşılaştırma Sonucu", padding="5")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.stats_label = ttk.Label(result_frame, text="Henüz karşılaştırma yapılmadı.", foreground='gray', font=('Arial', 9, 'italic'))
        self.stats_label.pack(anchor=tk.W)
        res_tree_cont = ttk.Frame(result_frame)
        res_tree_cont.pack(fill=tk.BOTH, expand=True)
        self.result_tree = self.create_treeview(res_tree_cont)
        
        log_frame = ttk.LabelFrame(main_container, text="📝 İşlem Logları", padding="5")
        log_frame.pack(fill=tk.X, pady=(5, 0))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, font=('Consolas', 8), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.setup_log_tags()
        
        if font_regular == "Helvetica":
            self.log_status("Uyarı: 'DejaVuSans.ttf' bulunamadı. Türkçe karakterler PDF'te hatalı görünebilir.", "WARN")

        self.result_df = None
        self.display_df = None
        self.df1 = None
        self.df2 = None
        self.current_selected_columns = None 
        self.root.bind('<Control-v>', self.handle_paste_shortcut)

    def create_treeview(self, parent):
        sby = ttk.Scrollbar(parent, orient=tk.VERTICAL)
        sbx = ttk.Scrollbar(parent, orient=tk.HORIZONTAL)
        tree = ttk.Treeview(parent, yscrollcommand=sby.set, xscrollcommand=sbx.set, show='tree headings', selectmode='extended')
        sby.config(command=tree.yview)
        sbx.config(command=tree.xview)
        tree.grid(row=0, column=0, sticky='nsew')
        sby.grid(row=0, column=1, sticky='ns')
        sbx.grid(row=1, column=0, sticky='ew')
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        return tree

    def setup_log_tags(self):
        self.log_text.tag_configure("INFO", foreground="black")
        self.log_text.tag_configure("WARN", foreground="orange")
        self.log_text.tag_configure("ERROR", foreground="red")
        self.log_text.tag_configure("DEBUG", foreground="gray")
        self.log_text.tag_configure("SUCCESS", foreground="green", font=('TkDefaultFont', 9, 'bold'))

    def log_status(self, message, level="INFO"):
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def handle_paste_shortcut(self, event):
        focused = self.root.focus_get()
        if str(self.tree1) in str(focused) or str(self.left_frame) in str(focused):
            self.paste_data(1)
        elif str(self.tree2) in str(focused) or str(self.right_frame) in str(focused):
            self.paste_data(2)

    def paste_data(self, tree_num):
        try:
            clipboard_data = self.root.clipboard_get()
            if not clipboard_data: return
            new_df = parse_clipboard_data(clipboard_data, self.log_status)
            if new_df is None: return

            if tree_num == 1:
                if self.df1 is not None and not self.df1.empty:
                    self.log_status(f"1. alana {len(new_df)} satır daha ekleniyor...", "INFO")
                    self.df1 = pd.concat([self.df1, new_df], ignore_index=True)
                else:
                    self.log_status(f"1. alana veri yapıştırıldı ({len(new_df)} satır).", "INFO")
                    self.df1 = new_df
                self.populate_tree(self.tree1, self.df1)
                self.count_label1.config(text=f"Satır: {len(self.df1)}")
            else:
                if self.df2 is not None and not self.df2.empty:
                    self.log_status(f"2. alana {len(new_df)} satır daha ekleniyor...", "INFO")
                    self.df2 = pd.concat([self.df2, new_df], ignore_index=True)
                else:
                    self.log_status(f"2. alana veri yapıştırıldı ({len(new_df)} satır).", "INFO")
                    self.df2 = new_df
                self.populate_tree(self.tree2, self.df2)
                self.count_label2.config(text=f"Satır: {len(self.df2)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Yapıştırma hatası: {e}")
            print(traceback.format_exc())

    def populate_tree(self, tree, df):
        tree.delete(*tree.get_children())
        if df is None or df.empty: return
        display_df_local = df.copy()
        
        if self.hide_empty_cols_var.get():
            def is_col_not_empty(series): return series.astype(str).str.strip().ne('').any()
            non_empty_cols = [col for col in display_df_local.columns if is_col_not_empty(display_df_local[col])]
            if non_empty_cols: display_df_local = display_df_local[non_empty_cols]

        columns = list(display_df_local.columns)
        tree['columns'] = columns
        tree.column('#0', width=0, stretch=tk.NO)
        for col in columns:
            tree.heading(col, text=col, anchor=tk.W)
            width = min(max(100, len(str(col)) * 10), 300)
            tree.column(col, width=width, anchor=tk.W)
        for _, row in display_df_local.iterrows():
            values = [str(val) for val in row]
            tree.insert('', tk.END, values=values)

    def refresh_all_views(self):
        if self.df1 is not None: self.populate_tree(self.tree1, self.df1)
        if self.df2 is not None: self.populate_tree(self.tree2, self.df2)
        if self.display_df is not None: self.populate_tree(self.result_tree, self.display_df)
        elif self.result_df is not None: self.populate_tree(self.result_tree, self.result_df)

    def clear_tree(self, tree_num):
        if tree_num == 1:
            self.tree1.delete(*self.tree1.get_children())
            self.tree1['columns'] = []
            self.df1 = None
            self.count_label1.config(text="Satır: 0")
        else:
            self.tree2.delete(*self.tree2.get_children())
            self.tree2['columns'] = []
            self.df2 = None
            self.count_label2.config(text="Satır: 0")

    def clear_all(self):
        self.clear_tree(1)
        self.clear_tree(2)
        self.result_tree.delete(*self.result_tree.get_children())
        self.result_df = None
        self.display_df = None
        self.current_selected_columns = None
        self.stats_label.config(text="Temizlendi.")
        self.btn_customize.config(state=tk.DISABLED)
        self.btn_pdf.config(state=tk.DISABLED)

    def compare_data(self):
        if self.df1 is None or self.df2 is None:
            messagebox.showwarning("Eksik Veri", "Her iki alana da veri yapıştırmalısınız.")
            return
        result = process_comparison(self.df1.copy(), self.df2.copy(), BASE_COLUMNS, self.log_status)
        if result is not None and not result.empty:
            self.result_df = result
            self.display_df = result
            self.current_selected_columns = list(result.columns)
            self.populate_tree(self.result_tree, result)
            msg = f"Toplam {len(result)} ortak kayıt bulundu."
            self.stats_label.config(text=msg, foreground='green', font=('Arial', 9, 'bold'))
            self.btn_customize.config(state=tk.NORMAL)
            self.btn_pdf.config(state=tk.NORMAL)
            messagebox.showinfo("Başarılı", msg)
        else:
            self.result_df = None
            self.display_df = None
            self.result_tree.delete(*self.result_tree.get_children())
            self.stats_label.config(text="Ortak kayıt bulunamadı.", foreground='red')
            self.btn_customize.config(state=tk.DISABLED)
            self.btn_pdf.config(state=tk.DISABLED)
            messagebox.showinfo("Sonuç", "Ortak kayıt bulunamadı.")

    def open_column_selector(self):
        if self.result_df is None: return
        all_columns = list(self.result_df.columns)
        initial_selection = self.current_selected_columns if self.current_selected_columns is not None else all_columns
        if self.hide_empty_cols_var.get():
            def is_col_not_empty(series): return series.astype(str).str.strip().ne('').any()
            non_empty_cols = [col for col in all_columns if is_col_not_empty(self.result_df[col])]
            initial_selection = [col for col in initial_selection if col in non_empty_cols]
        ColumnSelectorDialog(self.root, all_columns, initial_selection, self.apply_custom_view)

    def apply_custom_view(self, selected_columns):
        if self.result_df is None: return
        try:
            self.current_selected_columns = selected_columns
            self.display_df = self.result_df[selected_columns].copy()
            self.populate_tree(self.result_tree, self.display_df)
            self.log_status(f"Görünüm özelleştirildi: {len(selected_columns)} sütun gösteriliyor.", "INFO")
        except Exception as e:
            self.log_status(f"Görünüm güncellenirken hata: {e}", "ERROR")

    def copy_result_to_clipboard(self):
        df_to_copy = self.display_df if self.display_df is not None else self.result_df
        if df_to_copy is not None:
            self.root.clipboard_clear()
            self.root.clipboard_append(df_to_copy.to_csv(sep='\t', index=False))
            self.root.update()
            messagebox.showinfo("Kopyalandı", "Görüntülenen sonuçlar panoya kopyalandı.")
        else:
            messagebox.showwarning("Uyarı", "Kopyalanacak sonuç yok.")

    def open_pdf_editor(self):
        df_to_export = self.display_df if self.display_df is not None else self.result_df
        if df_to_export is None or df_to_export.empty:
            messagebox.showwarning("Uyarı", "PDF'e aktarılacak veri yok.")
            return
        PDFLayoutEditor(self.root, df_to_export, None)

if __name__ == "__main__":
    if not os.path.exists("DejaVuSans.ttf"):
        print("UYARI: 'DejaVuSans.ttf' dosyası bulunamadı. Türkçe karakterler PDF'te görünmeyebilir.")
    root = tk.Tk()
    app = PasteComparisonApp(root)
    root.mainloop()