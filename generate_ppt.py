from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Inisialisasi Presentasi Widescreen 16:9
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Palet Warna Tema (Dark Executive Style)
BG_COLOR = RGBColor(15, 23, 42)      # Slate 900
TEXT_MAIN = RGBColor(241, 245, 249)  # Slate 100
TEXT_MUTED = RGBColor(148, 163, 184) # Slate 400
ACCENT_CYAN = RGBColor(6, 182, 212)  # Cyan 500
ACCENT_GREEN = RGBColor(16, 185, 129)# Emerald 500

def set_slide_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

slide_layout = prs.slide_layouts[6] # Blank Layout

# --- SLIDE 1: COVER ---
slide1 = prs.slides.add_slide(slide_layout)
set_slide_background(slide1)
txBox1 = slide1.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(11), Inches(3.5))
tf1 = txBox1.text_frame
tf1.word_wrap = True

p0 = tf1.paragraphs[0]
p0.text = "AVAILABLE FOR FREELANCE PROJECT"
p0.font.size = Pt(14)
p0.font.bold = True
p0.font.color.rgb = ACCENT_CYAN
p0.space_after = Pt(20)

p1 = tf1.add_paragraph()
p1.text = "Muhammad Ulul Albab"
p1.font.size = Pt(44)
p1.font.bold = True
p1.font.color.rgb = TEXT_MAIN
p1.space_after = Pt(10)

p2 = tf1.add_paragraph()
p2.text = "Web Application Developer & Full-Stack Enthusiast"
p2.font.size = Pt(20)
p2.font.color.rgb = TEXT_MUTED
p2.space_after = Pt(20)

p3 = tf1.add_paragraph()
p3.text = "Portofolio Profesional & Rekam Jejak Solusi Digital Kustom"
p3.font.size = Pt(14)
p3.font.color.rgb = TEXT_MUTED


# --- SLIDE 2: KEAHLIAN & CORE TECH STACK ---
slide2 = prs.slides.add_slide(slide_layout)
set_slide_background(slide2)
tb2 = slide2.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11), Inches(5.5))
tf2 = tb2.text_frame
tf2.word_wrap = True

p = tf2.paragraphs[0]
p.text = "Keahlian & Core Tech Stack"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = TEXT_MAIN
p.space_after = Pt(30)

skills = [
    ("Frontend Architecture", "HTML5, CSS3, Modern JavaScript (ES6+), Tailwind CSS, dan desain UI/UX responsif yang bersih."),
    ("Web Application & Dashboard", "Pengembangan sistem manajemen data, kalkulator parameter fungsional, dan panel kontrol interaktif."),
    ("Backend & Database", "Integrasi database terstruktur, manajemen endpoint, dan arsitektur logika program yang efisien.")
]

for title, desc in skills:
    p_title = tf2.add_paragraph()
    p_title.text = f"• {title}"
    p_title.font.size = Pt(18)
    p_title.font.bold = True
    p_title.font.color.rgb = ACCENT_CYAN
    
    p_desc = tf2.add_paragraph()
    p_desc.text = f"   {desc}"
    p_desc.font.size = Pt(14)
    p_desc.font.color.rgb = TEXT_MUTED
    p_desc.space_after = Pt(15)


# --- SLIDE 3: FEATURED PROJECT 1 ---
slide3 = prs.slides.add_slide(slide_layout)
set_slide_background(slide3)
tb3 = slide3.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11), Inches(5.5))
tf3 = tb3.text_frame
tf3.word_wrap = True

p = tf3.paragraphs[0]
p.text = "Featured Project: Clinical Suite Dashboard"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = TEXT_MAIN
p.space_after = Pt(20)

p_sub = tf3.add_paragraph()
p_sub.text = "Sistem Manajemen & Kalkulator Parameter Medis Fungsional"
p_sub.font.size = Pt(18)
p_sub.font.color.rgb = ACCENT_CYAN
p_sub.space_after = Pt(20)

features1 = [
    "Aplikasi web interaktif dengan fitur pencatatan data dan parameter terstruktur.",
    "Dilengkapi kalkulator parameter otomatis untuk akurasi perhitungan.",
    "Antarmuka bersih, responsif, dan dirancang untuk efisiensi alur kerja operasional."
]
for f in features1:
    pf = tf3.add_paragraph()
    pf.text = f"✓ {f}"
    pf.font.size = Pt(15)
    pf.font.color.rgb = TEXT_MUTED
    pf.space_after = Pt(10)


# --- SLIDE 4: FEATURED PROJECT 2 ---
slide4 = prs.slides.add_slide(slide_layout)
set_slide_background(slide4)
tb4 = slide4.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11), Inches(5.5))
tf4 = tb4.text_frame
tf4.word_wrap = True

p = tf4.paragraphs[0]
p.text = "Featured Project: Sales & Warehouse Tracker"
p.font.size = Pt(32)
p.font.bold = True
p.font.color.rgb = TEXT_MAIN
p.space_after = Pt(20)

p_sub = tf4.add_paragraph()
p_sub.text = "Aplikasi Website Laporan Penjualan Realtime & Gudang"
p_sub.font.size = Pt(18)
p_sub.font.color.rgb = ACCENT_CYAN
p_sub.space_after = Pt(20)

features2 = [
    "Rekapitulasi transaksi terpusat lintas platform penjualan digital.",
    "Manajemen stok gudang otomatis (stok terpantau secara real-time).",
    "Dashboard laporan penjualan dan visualisasi data yang memudahkan pemantauan bisnis."
]
for f in features2:
    pf = tf4.add_paragraph()
    pf.text = f"✓ {f}"
    pf.font.size = Pt(15)
    pf.font.color.rgb = TEXT_MUTED
    pf.space_after = Pt(10)


# --- SLIDE 5: CONTACT & CTA ---
slide5 = prs.slides.add_slide(slide_layout)
set_slide_background(slide5)
tb5 = slide5.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11), Inches(4.5))
tf5 = tb5.text_frame
tf5.word_wrap = True

p = tf5.paragraphs[0]
p.text = "Tertarik Bekerjasama?"
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = TEXT_MAIN
p.space_after = Pt(15)

p_desc = tf5.add_paragraph()
p_desc.text = "Mari diskusikan kebutuhan proyek, pembuatan aplikasi custom, atau pengembangan sistem digital Anda bersama saya."
p_desc.font.size = Pt(16)
p_desc.font.color.rgb = TEXT_MUTED
p_desc.space_after = Pt(35)

p_wa = tf5.add_paragraph()
p_wa.text = "📱 WhatsApp: +62 895-4147-81707"
p_wa.font.size = Pt(22)
p_wa.font.bold = True
p_wa.font.color.rgb = ACCENT_GREEN

# Simpan file presentasi
prs.save("Portfolio_Muhammad_Ulul_Albab.pptx")
print("File PowerPoint berhasil diperbarui: Portfolio_Muhammad_Ulul_Albab.pptx")