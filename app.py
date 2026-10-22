import os
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["ONNXRUNTIME_DISABLE_THREAD_AFFINITY"] = "1"

import io
import re
import random
import warnings
warnings.filterwarnings("ignore")

import streamlit as st
from docx import Document
import pymupdf as fitz

# =====================================================================
# 1. DATABASE AI PATTERNS LENGKAP
# (Kompilasi GPTZero, Wikipedia, SearchAtlas, Grammarly, & blader/humanizer)
# =====================================================================

AI_PATTERNS_EN = {
    r"\bdelve into\b": ["explore", "examine", "look into", "analyze"],
    r"\bdelves into\b": ["explores", "examines", "looks into", "analyzes"],
    r"\bdelving into\b": ["exploring", "examining", "looking into"],
    r"\btapestry\b": ["mix", "collection", "combination", "structure"],
    r"\btestament to\b": ["proof of", "evidence of", "sign of", "example of"],
    r"\bbeacon of\b": ["source of", "example of", "symbol of"],
    r"\brealm of\b": ["field of", "area of", "domain of"],
    r"\blandscape of\b": ["field of", "state of", "environment of"],
    r"\bever-evolving\b": ["changing", "developing", "fast-changing"],
    r"\bcrucial\b": ["important", "key", "main", "major"],
    r"\bpivotal\b": ["important", "central", "key", "decisive"],
    r"\bparamount\b": ["essential", "top", "most important"],
    r"\bunderscore\b": ["highlight", "show", "emphasize", "point out"],
    r"\bunderscores\b": ["highlights", "shows", "emphasizes", "points out"],
    r"\bunderscoring\b": ["highlighting", "showing", "emphasizing"],
    r"\bmeticulous\b": ["careful", "detailed", "thorough"],
    r"\bmeticulously\b": ["carefully", "thoroughly", "in detail"],
    r"\bharness\b": ["use", "apply", "make use of"],
    r"\bharnessing\b": ["using", "applying", "making use of"],
    r"\bleverage\b": ["use", "apply", "utilize"],
    r"\bleveraging\b": ["using", "applying", "utilizing"],
    r"\bfoster\b": ["encourage", "support", "build", "promote"],
    r"\bfostering\b": ["encouraging", "supporting", "building"],
    r"\bembark on\b": ["start", "begin", "undertake"],
    r"\bnavigate the complexities\b": ["handle the challenges", "deal with the details", "manage the difficulties"],
    r"\bmultifaceted\b": ["complex", "many-sided", "varied"],
    r"\bnuanced\b": ["subtle", "detailed", "balanced"],
    r"\bintricate\b": ["complex", "detailed"],
    r"\bgame-changer\b": ["major shift", "big step forward", "breakthrough"],
    r"\bgame-changing\b": ["significant", "groundbreaking", "major"],
    r"\bparadigm shift\b": ["fundamental change", "major shift"],
    r"\bsymphony of\b": ["blend of", "combination of"],
    r"\bkaleidoscope of\b": ["variety of", "wide range of"],
    r"\bunleash\b": ["release", "unlock", "open up"],
    r"\bunlock the potential\b": ["make the most", "open new possibilities"],
    r"\bpave the way\b": ["prepare", "make it easier", "open the door"],
    r"\bcomprehensive\b": ["thorough", "complete", "full", "in-depth"],
    r"\boverarching\b": ["main", "overall", "broad"],
    r"\bseamless\b": ["smooth", "direct", "easy"],
    r"\bseamlessly\b": ["smoothly", "directly", "easily"],
    r"\bseamless integration\b": ["smooth compatibility", "direct integration"],
    r"\brobust\b": ["strong", "stable", "reliable"],
    r"\bendeavor\b": ["effort", "attempt", "work"],
    r"\butilize\b": ["use", "apply"],
    r"\butilizing\b": ["using", "applying"],
    r"\bfacilitate\b": ["help", "support", "enable"],
    r"\bfacilitates\b": ["helps", "supports", "enables"],
    r"\billuminate\b": ["explain", "clarify"],
    r"\billuminates\b": ["explains", "clarifies"],
    r"\bshed light on\b": ["clarify", "explain"],
    r"\bsheds light on\b": ["clarifies", "explains"],
    r"\bbolster\b": ["support", "strengthen", "reinforce"],
    r"\bbolsters\b": ["supports", "strengthens", "reinforces"],
    r"\bstreamline\b": ["simplify", "speed up"],
    r"\bstreamlines\b": ["simplifies", "speeds up"],
    r"\brevolutionize\b": ["transform", "change"],
    r"\bcutting-edge\b": ["advanced", "modern"],
    r"\btransformative\b": ["impactful", "major"],
    r"\bshowcasing\b": ["showing", "demonstrating", "presenting"],
    r"\balign with\b": ["match", "fit with", "support"],
    r"\baligns with\b": ["matches", "fits with", "supports"],

    r"\bMoreover,\s*": ["Also, ", "Beyond that, ", "In addition, "],
    r"\bFurthermore,\s*": ["Plus, ", "In addition, ", "Also, "],
    r"\bAdditionally,\s*": ["Also, ", "Alongside this, "],
    r"\bConsequently,\s*": ["As a result, ", "Therefore, ", "Because of this, "],
    r"\bNevertheless,\s*": ["Still, ", "Even so, ", "However, "],
    r"\bThat being said,\s*": ["However, ", "Even so, "],
    r"\bAt its core,\s*": ["Essentially, ", "Fundamentally, "],
    r"\bTo put it simply,\s*": ["In simple terms, ", "Simply put, "],
    r"\bA key takeaway is\b": ["The main point is", "One important point is"],
    r"\bFrom a broader perspective,\s*": ["On a larger scale, ", "More broadly, "],
    r"\bGenerally speaking,\s*": ["In most cases, ", "Usually, "],
    r"\bBroadly speaking,\s*": ["In general, ", "Overall, "],
    r"\bIt is important to note that\b": ["Note that", "Keep in mind that"],
    r"\bIt is worth noting that\b": ["Interestingly,", "Note that"],
    r"\bIt bears mentioning that\b": ["Also,", "Note that"],
    r"\bIn today's fast-paced world,\s*": ["Today, ", "Currently, "],
    r"\bIn today's digital age,\s*": ["Today, ", "With modern technology, "],
    r"\bIn conclusion,\s*": ["Overall, ", "To sum up, "],
    r"\bIn summary,\s*": ["In short, ", "Overall, "],
    r"\bA myriad of\b": ["Many", "Numerous", "A wide range of"],
    r"\bA plethora of\b": ["Many", "Plenty of", "Several"],
    r"\bPlays a crucial role\b": ["is important", "matters greatly", "is a key part"],
    r"\bPlays a pivotal role\b": ["is central", "is key", "is very important"],
}

AI_PATTERNS_ID = {
    r"\bDi era digital yang berkembang pesat saat ini,?\s*": ["Saat ini, ", "Sekarang, ", "Dewasa ini, "],
    r"\bDi tengah perkembangan zaman,?\s*": ["Saat ini, ", "Kini, "],
    r"\bPenting untuk dicatat bahwa\b": ["Perlu diingat bahwa", "Sebagai catatan,"],
    r"\bPatut dicatat bahwa\b": ["Perlu diperhatikan bahwa", "Menariknya,"],
    r"\bTidak dapat dipungkiri bahwa\b": ["Memang benar bahwa", "Jelas bahwa", "Faktanya,"],
    r"\bTak dapat dimungkiri bahwa\b": ["Memang,", "Faktanya,"],
    r"\bMemainkan peran krusial\b": ["sangat penting", "berperan besar", "menjadi kunci"],
    r"\bMemainkan peran penting\b": ["berperan penting", "sangat berpengaruh"],
    r"\bKrusial\b": ["penting", "utama", "kunci"],
    r"\bKomprehensif\b": ["lengkap", "menyeluruh", "mendalam"],
    r"\bSignifikan\b": ["besar", "nyata", "cukup terasa"],
    r"\bFundamental\b": ["mendasar", "utama", "pokok"],
    r"\bHolistik\b": ["menyeluruh", "utuh"],
    r"\bMenavigasi kompleksitas\b": ["menghadapi tantangan", "mengatasi kerumitan"],
    r"\bLanskap\b": ["kondisi", "situasi", "bidang"],
    r"\bPermadani\b": ["rangkaian", "keragaman", "campuran"],
    r"\bBukti nyata dari\b": ["bukti dari", "contoh dari"],
    r"\bPada intinya,\s*": ["Pada dasarnya, ", "Intinya, "],
    r"\bSecara garis besar,\s*": ["Umumnya, ", "Pada dasarnya, "],
    r"\bLebih lanjut,\s*": ["Selain itu, ", "Di samping itu, "],
    r"\bSelanjutnya,\s*": ["Berikutnya, ", "Lalu, "],
    r"\bSebagai kesimpulan,\s*": ["Singkatnya, ", "Pada akhirnya, ", "Secara keseluruhan, "],
    r"\bKesimpulannya,\s*": ["Intinya, ", "Secara garis besar, "],
    r"\bSecara keseluruhan,\s*": ["Pada dasarnya, ", "Intinya, "],
    r"\bMemanfaatkan\b": ["menggunakan", "memakai"],
    r"\bMengoptimalkan\b": ["memaksimalkan", "meningkatkan"],
    r"\bMendorong\b": ["membantu", "mendukung", "memicu"],
    r"\bMenyelami lebih dalam\b": ["membahas lebih jauh", "mempelajari", "mengulas"],
}


# =====================================================================
# 2. EKSTRAKSI LENGKAP (PARAGRAF + TABEL + TEXTBOX XML + HEADER/FOOTER)
# =====================================================================

def iter_all_docx_paragraphs(doc):
    for p in doc.paragraphs:
        yield p

    def iter_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
                    yield from iter_tables(cell.tables)

    yield from iter_tables(doc.tables)

    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        yield from iter_tables(section.header.tables)
        for p in section.footer.paragraphs:
            yield p
        yield from iter_tables(section.footer.tables)


def extract_full_text_from_docx_bytes(docx_bytes: bytes) -> tuple[str, int]:
    doc = Document(io.BytesIO(docx_bytes))
    all_xml_texts = []
    for p_elem in doc.element.body.iter():
        if p_elem.tag.endswith('}p'):
            texts = [t.text for t in p_elem.iter() if t.tag.endswith('}t') and t.text]
            joined = "".join(texts).strip()
            if joined:
                all_xml_texts.append(joined)

    full_text = "\n\n".join(all_xml_texts)
    ms_word_tokens = re.findall(r'\S+', full_text.replace("—", " ").replace("–", " "))
    return full_text, len(ms_word_tokens)


def extract_text_from_file(uploaded_file) -> tuple[str, int, bytes | None]:
    filename = uploaded_file.name.lower()
    raw_bytes = uploaded_file.read()

    if filename.endswith(".docx"):
        full_text, word_count = extract_full_text_from_docx_bytes(raw_bytes)
        return full_text, word_count, raw_bytes
    elif filename.endswith(".pdf"):
        doc = fitz.open(stream=raw_bytes, filetype="pdf")
        pages_text = [page.get_text("text") for page in doc]
        full_text = "\n\n".join(pages_text)
        word_count = len(re.findall(r'\S+', full_text))
        return full_text, word_count, None
    else:
        full_text = raw_bytes.decode("utf-8", errors="ignore")
        word_count = len(re.findall(r'\S+', full_text))
        return full_text, word_count, None


# =====================================================================
# 3. ANALISIS DETEKSI POLA AI & BURSTINESS
# =====================================================================

def analyze_ai_metrics(text: str) -> dict:
    if not text.strip():
        return {"ai_hits": 0, "em_dashes": 0, "burstiness": 0.0, "risk_score": 0, "detected_phrases": []}

    detected_phrases = []
    combined_patterns = {**AI_PATTERNS_EN, **AI_PATTERNS_ID}

    for pattern in combined_patterns.keys():
        matches = re.findall(pattern, text, flags=re.IGNORECASE)
        if matches:
            detected_phrases.extend(matches)

    em_dashes = text.count("—") + text.count(" -- ")

    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', text) if len(s.strip().split()) > 3]
    if len(sentences) > 1:
        lengths = [len(s.split()) for s in sentences]
        mean_len = sum(lengths) / len(lengths)
        variance = sum((x - mean_len) ** 2 for x in lengths) / len(lengths)
        burstiness = round(variance ** 0.5, 2)
    else:
        burstiness = 10.0

    word_count = max(len(re.findall(r'\S+', text)), 1)
    phrase_density = (len(detected_phrases) / word_count) * 1000
    low_burstiness_penalty = max(0, (9.0 - burstiness) * 3.5)
    em_dash_penalty = min(25, em_dashes * 2.5)

    risk_score = min(100, int(phrase_density * 5 + low_burstiness_penalty + em_dash_penalty))

    return {
        "ai_hits": len(detected_phrases),
        "em_dashes": em_dashes,
        "burstiness": burstiness,
        "risk_score": risk_score,
        "detected_phrases": sorted(list(set(m.strip() for m in detected_phrases))),
    }


# =====================================================================
# 4. ENGINE HUMANIZER (100% PROGRAMMATIC + IN-PLACE DOCX REWRITER)
# =====================================================================

def preserve_case_replace(match, replacements):
    original = match.group(0)
    chosen = random.choice(replacements)
    if not chosen:
        return ""
    if original.isupper():
        return chosen.upper()
    elif original[0].isupper():
        return chosen[0].upper() + chosen[1:]
    return chosen.lower()


def improve_burstiness(paragraph: str) -> str:
    sentences = re.split(r'(?<=[.!?])\s+', paragraph)
    new_sentences = []

    for sent in sentences:
        words = sent.split()
        if len(words) > 26:
            if re.search(r',\s+(which|while|whereas|although|and it)\s+', sent):
                sent = re.sub(
                    r',\s+which\s+',
                    lambda _: random.choice([". This ", ". That "]),
                    sent,
                    count=1
                )
                sent = re.sub(r',\s+while\s+', ". Meanwhile, ", sent, count=1)
                sent = re.sub(r',\s+whereas\s+', ". In contrast, ", sent, count=1)
            elif re.search(r',\s+(yang mana|sedangkan|sementara itu|meskipun)\s+', sent, re.IGNORECASE):
                sent = re.sub(r',\s+yang mana\s+', ". Hal ini ", sent, count=1, flags=re.IGNORECASE)
                sent = re.sub(r',\s+sementara itu\s+', ". Di sisi lain, ", sent, count=1, flags=re.IGNORECASE)
        new_sentences.append(sent)

    return " ".join(new_sentences)


def programmatic_humanize(
    text: str,
    lang: str = "Auto",
    fix_em_dash: bool = True,
    boost_burstiness: bool = True,
    strip_markdown_bold: bool = True
) -> str:
    if not text.strip():
        return text

    result = text

    if strip_markdown_bold:
        result = re.sub(r'\*\*(.*?)\*\*', r'\1', result)

    result = result.replace("\u200b", "").replace("\u00a0", " ")

    if fix_em_dash:
        result = re.sub(
            r'(?<!Abstract)\s*—\s*',
            lambda _: random.choice([", ", " - ", ": "]),
            result
        )

    patterns_to_use = {}
    if lang in ["Auto", "English"]:
        patterns_to_use.update(AI_PATTERNS_EN)
    if lang in ["Auto", "Indonesia"]:
        patterns_to_use.update(AI_PATTERNS_ID)

    for pattern, replacements in patterns_to_use.items():
        result = re.sub(
            pattern,
            lambda m, r=replacements: preserve_case_replace(m, r),
            result,
            flags=re.IGNORECASE
        )

    if boost_burstiness:
        paragraphs = result.split("\n\n")
        paragraphs = [improve_burstiness(p) for p in paragraphs]
        result = "\n\n".join(paragraphs)

    result = re.sub(r'[ \t]{2,}', ' ', result)
    result = re.sub(r'\s+([.,;!?])', r'\1', result)
    return result


def humanize_docx_in_place(
    docx_bytes: bytes,
    lang: str,
    fix_em_dash: bool,
    boost_burstiness: bool,
    strip_markdown_bold: bool
) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    seen_elements = set()

    for p in iter_all_docx_paragraphs(doc):
        if id(p._element) in seen_elements:
            continue
        seen_elements.add(id(p._element))

        orig_text = p.text
        if not orig_text or len(orig_text.strip()) < 4:
            continue

        if "@" in orig_text and len(orig_text.split()) < 15:
            continue

        new_text = programmatic_humanize(
            orig_text,
            lang=lang,
            fix_em_dash=fix_em_dash,
            boost_burstiness=boost_burstiness,
            strip_markdown_bold=strip_markdown_bold
        )

        if new_text != orig_text and p.runs:
            p.runs[0].text = new_text
            for r in p.runs[1:]:
                r.text = ""

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


def create_plain_docx_bytes(text: str) -> bytes:
    doc = Document()
    for para in text.split("\n\n"):
        clean_p = para.strip()
        if clean_p:
            doc.add_paragraph(clean_p)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =====================================================================
# 5. ANTARMUKA WEB UI (STREAMLIT)
# =====================================================================

st.set_page_config(
    page_title="Turnitout Server - AI Humanizer & De-AI Pattern",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Self-Hosted AI Humanizer & Similarity Evader")
st.caption(
    "100% Programmatic Engine berdasarkan standar **GPTZero**, **Wikipedia**, **SearchAtlas**, & **Grammarly** "
    "— Tanpa batas kata & menjaga layout asli Word (`.docx`)."
)

with st.sidebar:
    st.header("⚙️ Pengaturan Filter")
    lang_choice = st.selectbox("Bahasa Dokumen:", ["Auto", "English", "Indonesia"])

    st.subheader("Aturan De-AI Pattern")
    opt_em_dash = st.checkbox("Hapus Em-Dash (—) Berlebih", value=True)
    opt_burstiness = st.checkbox("Tingkatkan Burstiness (Variasi Kalimat)", value=True)
    opt_strip_bold = st.checkbox("Bersihkan Bold Markdown (**)", value=True)

tab_upload, tab_paste = st.tabs(["📂 Upload Dokumen (Word / PDF / TXT)", "✍️ Paste Teks Langsung"])

input_text = ""
raw_docx_bytes = None
total_words_detected = 0

with tab_upload:
    uploaded_file = st.file_uploader(
        "Pilih file dokumen (.docx, .pdf, .txt, .md, .tex) — Tidak ada batas jumlah kata",
        type=["docx", "pdf", "txt", "md", "tex"]
    )
    if uploaded_file is not None:
        try:
            input_text, total_words_detected, raw_docx_bytes = extract_text_from_file(uploaded_file)
            st.success(
                f"Berhasil membaca seluruh elemen **{uploaded_file.name}** "
                f"(Total: **{total_words_detected:,} kata** termasuk isi Tabel, Caption & Textbox — Layout asli terjaga)."
            )
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

with tab_paste:
    pasted_text = st.text_area(
        "Masukkan atau tempel teks di sini (Bebas tanpa batas karakter):",
        height=260,
        placeholder="Tempel teks panjang atau seluruh isi jurnal di sini..."
    )
    if pasted_text.strip():
        input_text = pasted_text
        total_words_detected = len(re.findall(r'\S+', input_text))
        raw_docx_bytes = None

if st.button("🚀 Proses & Humanize Dokumen", type="primary", use_container_width=True):
    if not input_text.strip():
        st.warning("Silakan unggah dokumen atau tempelkan teks terlebih dahulu.")
    else:
        with st.spinner("Menganalisis pola AI dan memproses dokumen..."):
            before_metrics = analyze_ai_metrics(input_text)

            if raw_docx_bytes is not None:
                output_docx_bytes = humanize_docx_in_place(
                    raw_docx_bytes,
                    lang=lang_choice,
                    fix_em_dash=opt_em_dash,
                    boost_burstiness=opt_burstiness,
                    strip_markdown_bold=opt_strip_bold
                )
                output_text, after_word_count = extract_full_text_from_docx_bytes(output_docx_bytes)
            else:
                output_text = programmatic_humanize(
                    input_text,
                    lang=lang_choice,
                    fix_em_dash=opt_em_dash,
                    boost_burstiness=opt_burstiness,
                    strip_markdown_bold=opt_strip_bold
                )
                output_docx_bytes = create_plain_docx_bytes(output_text)
                after_word_count = len(re.findall(r'\S+', output_text.replace("—", " ")))

            after_metrics = analyze_ai_metrics(output_text)

        st.subheader("📊 Hasil Analisis Pola AI")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(
            "Total Kata Diproses",
            f"{after_word_count:,}",
            delta=f"{after_word_count - total_words_detected} kata"
        )
        m2.metric(
            "Skor Indikator Pola AI",
            f"{after_metrics['risk_score']}%",
            delta=f"{after_metrics['risk_score'] - before_metrics['risk_score']}%",
            delta_color="inverse"
        )
        m3.metric(
            "Frasa Klise AI Ditemukan",
            after_metrics["ai_hits"],
            delta=after_metrics["ai_hits"] - before_metrics["ai_hits"],
            delta_color="inverse"
        )
        m4.metric(
            "Skor Burstiness",
            after_metrics["burstiness"],
            delta=round(after_metrics["burstiness"] - before_metrics["burstiness"], 2)
        )
        m5.metric(
            "Jumlah Em-Dash (—)",
            after_metrics["em_dashes"],
            delta=after_metrics["em_dashes"] - before_metrics["em_dashes"],
            delta_color="inverse"
        )

        if before_metrics["detected_phrases"]:
            with st.expander(f"🔍 Lihat Daftar Kata/Frasa AI yang Terdeteksi & Diganti ({len(before_metrics['detected_phrases'])} jenis)"):
                st.write(", ".join([f"`{p}`" for p in before_metrics["detected_phrases"]]))

        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown(f"### 📄 Teks Asli ({total_words_detected:,} kata)")
            st.text_area("Before", value=input_text, height=450, label_visibility="collapsed")
        with col_right:
            st.markdown(f"### ✨ Hasil Humanizer ({after_word_count:,} kata)")
            st.text_area("After", value=output_text, height=450, label_visibility="collapsed")

        st.divider()
        st.subheader("💾 Unduh Dokumen Hasil (Format & Tabel Utuh)")
        dl_col1, dl_col2 = st.columns(2)

        with dl_col1:
            st.download_button(
                label="📥 Download Microsoft Word (.docx) — Format Asli Terjaga",
                data=output_docx_bytes,
                file_name="Humanized_Document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                label="📥 Download File Teks (.txt)",
                data=output_text.encode("utf-8"),
                file_name="Humanized_Document.txt",
                mime="text/plain",
                use_container_width=True
            )