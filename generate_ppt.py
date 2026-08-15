import sys
import os

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("Library python-pptx belum terinstall!")
    print("Silakan jalankan perintah ini di terminal VS Code/CMD:")
    print("pip install python-pptx")
    sys.exit(1)

def create_portfolio_pptx():
    prs = Presentation()
    # Set Slide Size to Widescreen 16:9 (13.33 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6] # Blank Layout

    # Color Palette (Dark Theme Modern)
    BG_COLOR = RGBColor(15, 23, 42)       # Slate 900 (#0F172A)
    CARD_BG = RGBColor(30, 41, 59)        # Slate 800 (#1E293B)
    CARD_BORDER = RGBColor(51, 65, 85)    # Slate 700 (#334155)
    PRIMARY_TEXT = RGBColor(248, 250, 252)# Slate 50 (#F8FAFC)
    SECONDARY_TEXT = RGBColor(148, 163, 184) # Slate 400 (#94A3B8)
    
    CYAN_ACCENT = RGBColor(6, 182, 212)   # Cyan 500 (#06B6D4)
    ROSE_ACCENT = RGBColor(244, 63, 94)   # Rose 500 (#F43F5E)
    ORANGE_ACCENT = RGBColor(249, 115, 22) # Orange 500 (#F97316)
    SKY_ACCENT = RGBColor(2, 132, 199)    # Sky 600 (#0284C7)
    PURPLE_ACCENT = RGBColor(139, 92, 246) # Purple 500 (#8B5CF6)
    EMERALD_ACCENT = RGBColor(16, 185, 129) # Emerald 500 (#10B981)

    def set_slide_background(slide):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = BG_COLOR

    def add_header(slide, title_text, category_text="PORTOFOLIO DIGITAL"):
        # Header Box
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.733), Inches(1.0))
        tf = header_box.text_frame
        tf.word_wrap = True
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = CYAN_ACCENT
        p_cat.font.name = "Arial"
        
        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY_TEXT
        p_title.font.name = "Arial"

    # ==========================================
    # SLIDE 1: COVER
    # ==========================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1)
    
    # Decorative Accent Bar
    bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.2), Inches(0.15), Inches(3.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = CYAN_ACCENT
    bar.line.color.rgb = CYAN_ACCENT
    
    # Title Text Box
    tb1 = slide1.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "AVAILABLE FOR FREELANCE & COLLABORATION"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = CYAN_ACCENT
    
    p2 = tf1.add_paragraph()
    p2.text = "Muhammad Ulul Albab"
    p2.font.size = Pt(44)
    p2.font.bold = True
    p2.font.color.rgb = PRIMARY_TEXT
    
    p3 = tf1.add_paragraph()
    p3.text = "Fullstack & AI Application Developer"
    p3.font.size = Pt(22)
    p3.font.color.rgb = SECONDARY_TEXT
    
    p4 = tf1.add_paragraph()
    p4.text = "Perpaduan Ketelitian Analisis Farmasi & Rekayasa Perangkat Lunak Modern"
    p4.font.size = Pt(14)
    p4.font.italic = True
    p4.font.color.rgb = CYAN_ACCENT

    # ==========================================
    # SLIDE 2: ABOUT ME & BACKGROUND
    # ==========================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2)
    add_header(slide2, "Tentang Saya & Latar Belakang Analitis", "Profil Profesional")

    # Left Box (Farmasi Background)
    shape_left = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    shape_left.fill.solid()
    shape_left.fill.fore_color.rgb = CARD_BG
    shape_left.line.color.rgb = CARD_BORDER
    
    tf_l = shape_left.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = Inches(0.3)
    tf_l.margin_top = Inches(0.3)
    
    p = tf_l.paragraphs[0]
    p.text = "Pendidikan & Fondasi Eksakta"
    p.font.bold = True
    p.font.size = Pt(18)
    p.font.color.rgb = CYAN_ACCENT
    
    p = tf_l.add_paragraph()
    p.text = "• Universitas Pakuan — Program Studi Farmasi"
    p.font.size = Pt(14)
    p.font.color.rgb = PRIMARY_TEXT
    
    p = tf_l.add_paragraph()
    p.text = "\nLatar belakang Farmasi membentuk ketelitian tinggi, logika eksploratif, serta pola pikir analitis eksakta yang kuat dalam memecahkan masalah kompleks."
    p.font.size = Pt(13)
    p.font.color.rgb = SECONDARY_TEXT

    # Right Box (Developer Edge)
    shape_right = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8))
    shape_right.fill.solid()
    shape_right.fill.fore_color.rgb = CARD_BG
    shape_right.line.color.rgb = CARD_BORDER
    
    tf_r = shape_right.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = Inches(0.3)
    tf_r.margin_top = Inches(0.3)
    
    p = tf_r.paragraphs[0]
    p.text = "Transisi Ke Software Engineering"
    p.font.bold = True
    p.font.size = Pt(18)
    p.font.color.rgb = ROSE_ACCENT
    
    points = [
        "Fokus pada pembentukan arsitektur sistem yang bersih, efisien, dan siap pakai di lingkungan cloud.",
        "Menguasai pemecahan masalah teknis end-to-end: dari profil memori PyTorch/Docker hingga optimasi database.",
        "Komitmen pada pembelajaran berkelanjutan tanpa ragu menghadapi tantangan teknologi baru."
    ]
    for pt in points:
        p = tf_r.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(13)
        p.font.color.rgb = SECONDARY_TEXT

    # ==========================================
    # SLIDE 3: TECH STACK MATRIX
    # ==========================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3)
    add_header(slide3, "Keahlian & Ekosistem Teknologi", "Technical Stack")

    stacks = [
        ("Frontend & Mobile", "React, React Native, Expo, Tailwind CSS, Vite, HTML/JS", CYAN_ACCENT, 0.8, 1.8),
        ("Backend & API", "FastAPI (Python), Laravel 11 (PHP), Node.js / Express", ROSE_ACCENT, 6.8, 1.8),
        ("AI & Data Intelligence", "Gemini RAG API, PyTorch, SentenceTransformers, pgvector", PURPLE_ACCENT, 0.8, 4.3),
        ("Database & Cloud", "PostgreSQL, MySQL, Firebase Firestore, SQLite, Docker", EMERALD_ACCENT, 6.8, 4.3)
    ]

    for title, desc, color, x, y in stacks:
        card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(5.7), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.3)
        tf.margin_top = Inches(0.2)
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.color.rgb = color
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(13)
        p2.font.color.rgb = PRIMARY_TEXT

    # ==========================================
    # SLIDE 4: PROJECT 1 - SCENTDNA
    # ==========================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4)
    add_header(slide4, "ScentDNA — AI Fragrance Discovery Engine", "Featured AI Project")

    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.8))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = ROSE_ACCENT

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.4)
    tf.margin_top = Inches(0.3)

    p = tf.paragraphs[0]
    p.text = "Mesin Rekomendasi Aroma Vector Search & RAG AI (Dockerized)"
    p.font.bold = True
    p.font.size = Pt(18)
    p.font.color.rgb = ROSE_ACCENT

    p = tf.add_paragraph()
    p.text = "Tech Stack: FastAPI (Python), PyTorch, SentenceTransformers, pgvector (PostgreSQL), Docker, Gemini 2.5 Flash API\n"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = CYAN_ACCENT

    highlights = [
        "Semantic Vector Search: Menggunakan pembacaan kemiripan kosinus (<=>) berbasis pgvector PostgreSQL.",
        "RAG Gemini Consultant: Rekomendasi kontekstual terstruktur dari hasil pencarian vektor produk.",
        "Singleton Pattern (Dependency Injection): Restrukturisasi lifespan FastAPI untuk hemat alokasi RAM.",
        "Memory Hardening: Optimasi PyTorch CPU single-threading & garbage collection aktif untuk stabilitas cloud deployment."
    ]
    for h in highlights:
        p = tf.add_paragraph()
        p.text = f"✓  {h}"
        p.font.size = Pt(13)
        p.font.color.rgb = PRIMARY_TEXT

    # ==========================================
    # SLIDE 5: PROJECT 2 - PODLEARN AI & TELEGRAM BOT
    # ==========================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5)
    add_header(slide5, "AI Solutions: PodLearn AI & Telegram Assistant", "Featured AI Projects")

    # Podlearn
    c1 = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = CARD_BG
    c1.line.color.rgb = PURPLE_ACCENT
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.3)
    tf1.margin_top = Inches(0.3)

    p = tf1.paragraphs[0]
    p.text = "PodLearn AI — Podcast & Quiz"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = PURPLE_ACCENT

    p = tf1.add_paragraph()
    p.text = "Mengubah file PDF/teks menjadi audio podcast 2 orang (Host & Expert) secara otomatis.\n"
    p.font.size = Pt(12)
    p.font.color.rgb = SECONDARY_TEXT

    pts = [
        "Integrasi Gemini AI untuk naskah dialog.",
        "Microsoft Edge Neural TTS multi-suara.",
        "FFmpeg Concat untuk penggabungan audio.",
        "Generator 10 kuis evaluasi interaktif."
    ]
    for pt in pts:
        p = tf1.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # Telegram Bot
    c2 = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8))
    c2.fill.solid()
    c2.fill.fore_color.rgb = CARD_BG
    c2.line.color.rgb = PURPLE_ACCENT
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.3)
    tf2.margin_top = Inches(0.3)

    p = tf2.paragraphs[0]
    p.text = "Telegram AI Assistant (Live 24/7)"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = PURPLE_ACCENT

    p = tf2.add_paragraph()
    p.text = "Bot Telegram personal cerdas yang aktif 24/7 di cloud server tanpa henti.\n"
    p.font.size = Pt(12)
    p.font.color.rgb = SECONDARY_TEXT

    pts2 = [
        "Model LLM Llama-3.3 70B via Groq API.",
        "Respon cerdas multibahasa real-time.",
        "System prompt kustom untuk persona unik.",
        "Deployed 24/7 di Railway Cloud."
    ]
    for pt in pts2:
        p = tf2.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # ==========================================
    # SLIDE 6: PROJECT 3 - LDR ANCHOR & CLINICAL SUITE
    # ==========================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6)
    add_header(slide6, "Mobile & Web Dashboard Systems", "Functional Systems")

    # LDR Anchor
    c1 = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = CARD_BG
    c1.line.color.rgb = ORANGE_ACCENT
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.3)
    tf1.margin_top = Inches(0.3)

    p = tf1.paragraphs[0]
    p.text = "LDR Anchor (React Native App)"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = ORANGE_ACCENT

    pts = [
        "Aplikasi mobile Android khusus pasangan jarak jauh (LDR).",
        "Fitur Mood sharing real-time & pelukan virtual interaktif.",
        "Integrasi FCM V1 Push Notification via Expo.",
        "Firebase Firestore & Auth backend integration."
    ]
    for pt in pts:
        p = tf1.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # Clinical Suite
    c2 = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8))
    c2.fill.solid()
    c2.fill.fore_color.rgb = CARD_BG
    c2.line.color.rgb = CYAN_ACCENT
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.3)
    tf2.margin_top = Inches(0.3)

    p = tf2.paragraphs[0]
    p.text = "Clinical Suite Dashboard"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = CYAN_ACCENT

    pts2 = [
        "Sistem rekapitulasi data medis & manajemen klinis.",
        "Kalkulator parameter medis fungsional otomatis.",
        "Antarmuka bersih & presisi tinggi berbasis Tailwind.",
        "Dirancang khusus untuk efisiensi operasional medis."
    ]
    for pt in pts2:
        p = tf2.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # ==========================================
    # SLIDE 7: PROJECT 4 - CLOUD INVENTORY & DOMAIN EXPLORER
    # ==========================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7)
    add_header(slide7, "Enterprise Serverless & Laravel Systems", "Web Applications")

    # Cloud Inventory
    c1 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    c1.fill.solid()
    c1.fill.fore_color.rgb = CARD_BG
    c1.line.color.rgb = SKY_ACCENT
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.3)
    tf1.margin_top = Inches(0.3)

    p = tf1.paragraphs[0]
    p.text = "Cloud Inventory System (Serverless)"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = SKY_ACCENT

    pts = [
        "Sistem inventaris berbasis Google Apps Script.",
        "Concurrency Protection via LockService.",
        "Role-Based Access Control (RBAC) & Anti-XSS.",
        "Audit Trail log & snapshot auto-backup ke Drive."
    ]
    for pt in pts:
        p = tf1.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # Domain Explorer
    c2 = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8))
    c2.fill.solid()
    c2.fill.fore_color.rgb = CARD_BG
    c2.line.color.rgb = EMERALD_ACCENT
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.3)
    tf2.margin_top = Inches(0.3)

    p = tf2.paragraphs[0]
    p.text = "Domain Explorer (Laravel 11)"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = EMERALD_ACCENT

    pts2 = [
        "Aplikasi manajemen domain berbasis Laravel 11.",
        "Arsitektur MVC bersih dengan Blade views.",
        "Penggunaan SQLite database untuk performa ringan.",
        "Database Migrations & Seeders otomatis."
    ]
    for pt in pts2:
        p = tf2.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.font.color.rgb = PRIMARY_TEXT

    # ==========================================
    # SLIDE 8: CONTACT & CLOSING
    # ==========================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8)
    
    card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.5), Inches(10.333), Inches(4.5))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CYAN_ACCENT

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.5)
    tf.margin_top = Inches(0.5)

    p = tf.paragraphs[0]
    p.text = "Mari Berkolaborasi!"
    p.font.bold = True
    p.font.size = Pt(28)
    p.font.color.rgb = CYAN_ACCENT

    p = tf.add_paragraph()
    p.text = "Terbuka untuk proyek freelance, pengembangan sistem kustom, maupun posisi software engineer.\n"
    p.font.size = Pt(14)
    p.font.color.rgb = PRIMARY_TEXT

    contacts = [
        "📱 WhatsApp : +62 895-4147-81707",
        "✉️ Email    : ulula2812@gmail.com",
        "🌐 Portofolio: https://ulul-portofolio-3odd.vercel.app/",
        "📍 Lokasi   : Bogor, Jawa Barat, Indonesia"
    ]
    for c in contacts:
        p = tf.add_paragraph()
        p.text = c
        p.font.size = Pt(15)
        p.font.color.rgb = SECONDARY_TEXT

    # Save
    output_filename = "Portfolio_Muhammad_Ulul_Albab.pptx"
    prs.save(output_filename)
    print(f"SUKSES! Slide PPT Portofolio berhasil diperbarui ke file: {output_filename}")

if __name__ == "__main__":
    create_portfolio_pptx()