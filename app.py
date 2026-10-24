import os
os.environ["OMP_NUM_THREADS"] = "1"
os.environ["ONNXRUNTIME_DISABLE_THREAD_AFFINITY"] = "1"

import io
import re
import html
import difflib
import random
import warnings
warnings.filterwarnings("ignore")

import streamlit as st
from docx import Document
import pymupdf as fitz

# =====================================================================
# 1. DATABASE HUMANIZER V2.5.2 (WIKIPEDIA & GPTZERO PATTERNS)
# =====================================================================

AI_PATTERNS_EN = {
    r"\bplays a crucial role\b": ["is very important", "is a key factor", "plays an important part"],
    r"\bplays a pivotal role\b": ["is central", "is a key factor", "is essential"],
    r"\bplays a vital role\b": ["is essential", "is very important", "matters greatly"],
    r"\bserves as a testament to\b": ["demonstrates", "reflects", "shows clearly"],
    r"\bstands as a testament to\b": ["reflects", "shows", "highlights"],
    r"\btestament to\b": ["evidence of", "proof of", "reflection of"],
    r"\bindelible mark\b": ["lasting impact", "clear mark", "strong effect"],
    r"\bbeacon of hope\b": ["positive sign", "encouragement", "source of optimism"],
    r"\bbeacon of\b": ["source of", "symbol of", "example of"],
    r"\brich tapestry\b": ["wide variety", "broad combination", "mix"],
    r"\btapestry of\b": ["combination of", "mix of", "range of"],
    r"\bdelve deeper into\b": ["examine closely", "explore further", "look into"],
    r"\bdelve into\b": ["explore", "examine", "analyze", "look into"],
    r"\bdelves into\b": ["explores", "examines", "analyzes", "looks into"],
    r"\bdelving into\b": ["exploring", "examining", "looking into"],
    r"\bnavigate the complexities of\b": ["handle the challenges of", "manage", "deal with"],
    r"\bpave the way for\b": ["prepare for", "enable", "support"],
    r"\bunlock the potential\b": ["maximize the potential", "open new possibilities"],
    r"\bparadigm shift\b": ["major change", "significant shift", "fundamental change"],
    r"\bgame-changer\b": ["major breakthrough", "significant advance", "turning point"],
    r"\bseamless integration\b": ["smooth integration", "direct connection"],
    r"\bshed light on\b": ["clarify", "explain", "shed light upon"],
    r"\bsheds light on\b": ["clarifies", "explains", "highlights"],
    r"\ba myriad of\b": ["many", "numerous", "various"],
    r"\ba plethora of\b": ["many", "plenty of", "several"],
    r"\bin today's fast-paced world,?\s*": ["Today, ", "Currently, "],
    r"\bin today's digital age,?\s*": ["In this digital era, ", "Today, "],
    r"\bit is important to note that\b": ["Notably,", "It is worth noting that", "Keep in mind that"],
    r"\bit is worth noting that\b": ["Interestingly,", "Notably,", "It should be noted that"],
    r"\bmoreover,\s*": ["Additionally, ", "Furthermore, ", "Also, "],
    r"\bfurthermore,\s*": ["Additionally, ", "Moreover, ", "Also, "],
    r"\bconsequently,\s*": ["Therefore, ", "As a result, "],
    r"\bnevertheless,\s*": ["However, ", "Even so, "],
    r"\brealm of\b": ["field of", "domain of", "area of"],
    r"\blandscape of\b": ["environment of", "field of", "state of"],
    r"\bever-evolving\b": ["changing", "fast-changing", "dynamic"],
    r"\bcrucial\b": ["important", "key", "vital"],
    r"\bpivotal\b": ["central", "key", "main"],
    r"\bparamount\b": ["essential", "top", "fundamental"],
    r"\bunderscore\b": ["highlight", "emphasize", "show"],
    r"\bunderscores\b": ["highlights", "emphasizes", "shows"],
    r"\bunderscoring\b": ["highlighting", "emphasizing", "showing"],
    r"\bmeticulous\b": ["careful", "thorough", "detailed"],
    r"\bmeticulously\b": ["carefully", "thoroughly", "in detail"],
    r"\bharness\b": ["use", "apply", "utilize"],
    r"\bharnessing\b": ["using", "applying", "utilizing"],
    r"\bleverage\b": ["use", "employ", "utilize"],
    r"\bleveraging\b": ["using", "employing", "utilizing"],
    r"\bfoster\b": ["support", "encourage", "promote"],
    r"\bfostering\b": ["supporting", "encouraging", "promoting"],
    r"\bcomprehensive\b": ["thorough", "complete", "detailed"],
    r"\bseamless\b": ["smooth", "continuous", "direct"],
    r"\bseamlessly\b": ["smoothly", "directly", "continuously"],
    r"\brobust\b": ["strong", "reliable", "solid"],
    r"\butilize\b": ["use", "employ"],
    r"\butilizing\b": ["using", "employing"],
    r"\bfacilitate\b": ["help", "support", "enable"],
    r"\bfacilitates\b": ["helps", "supports", "enables"],
    r"\billuminate\b": ["explain", "clarify"],
    r"\bbolster\b": ["strengthen", "support"],
    r"\bstreamline\b": ["simplify", "improve"],
    r"\brevolutionize\b": ["transform", "change significantly"],
    r"\bcutting-edge\b": ["advanced", "modern"],
    r"\btransformative\b": ["impactful", "significant"],
    r"\bshowcasing\b": ["demonstrating", "displaying", "showing"],
}

AI_PATTERNS_ID = {
    r"\bdi era digital yang berkembang pesat saat ini,?\s*": ["Saat ini, ", "Kini, ", "Dewasa ini, "],
    r"\bdi tengah perkembangan zaman,?\s*": ["Saat ini, ", "Kini, "],
    r"\bpenting untuk dicatat bahwa\b": ["Perlu diingat bahwa", "Perlu dicatat bahwa"],
    r"\bpatut dicatat bahwa\b": ["Perlu diperhatikan bahwa", "Menariknya,"],
    r"\btidak dapat dipungkiri bahwa\b": ["Memang benar bahwa", "Jelas bahwa"],
    r"\btak dapat dimungkiri bahwa\b": ["Faktanya,", "Memang,"],
    r"\bmemainkan peran krusial\b": ["berperan penting", "sangat menentukan", "menjadi kunci"],
    r"\bmemainkan peran penting\b": ["berperan besar", "sangat berpengaruh"],
    r"\bmenavigasi kompleksitas\b": ["menghadapi tantangan", "mengatasi kendala"],
    r"\bmenyelami lebih dalam\b": ["mempelajari lebih lanjut", "mengkaji lebih jauh"],
    r"\bbukti nyata dari\b": ["bukti dari", "contoh dari"],
    r"\bpada intinya,\s*": ["Pada dasarnya, ", "Intinya, "],
    r"\bsecara garis besar,\s*": ["Umumnya, ", "Secara umum, "],
    r"\blebih lanjut,\s*": ["Selain itu, ", "Di samping itu, "],
    r"\bsebagai kesimpulan,\s*": ["Singkatnya, ", "Secara keseluruhan, "],
    r"\bkesimpulannya,\s*": ["Intinya, ", "Secara ringkas, "],
    r"\bsecara keseluruhan,\s*": ["Pada dasarnya, ", "Secara umum, "],
    r"\bkrusial\b": ["penting", "utama", "mendasar"],
    r"\bkomprehensif\b": ["menyeluruh", "lengkap", "utuh"],
    r"\bsignifikan\b": ["besar", "nyata", "cukup berarti"],
    r"\bfundamental\b": ["mendasar", "pokok", "utama"],
    r"\bholistik\b": ["menyeluruh", "terpadu"],
    r"\blanskap\b": ["kondisi", "situasi", "bidang"],
    r"\bpermadani\b": ["rangkaian", "kombinasi"],
    r"\bmemanfaatkan\b": ["menggunakan", "memakai"],
    r"\bmengoptimalkan\b": ["memaksimalkan", "meningkatkan"],
    r"\bmendorong\b": ["membantu", "mendukung", "memicu"],
    r"\bmemiliki karakteristik utama berupa\b": ["memiliki ciri khas berupa", "menunjukkan karakteristik berupa"],
    r"\bdiwujudkan melalui\b": ["diterapkan lewat", "diwujudkan lewat"],
    r"\bdifokuskan pada\b": ["dipusatkan pada", "diarahkan ke"],
    r"\bberlanjut hingga\b": ["berlangsung sampai", "terus berjalan hingga"],
    r"\bberdasarkan hasil pengamatan\b": ["dari hasil pengamatan", "mengacu pada pengamatan"],
    r"\bmempermudah\b": ["memudahkan", "membantu"],
    r"\bterstruktur\b": ["tersusun", "tertata", "tersusun dengan baik"],
}

@st.cache_data
def get_sorted_patterns(lang: str) -> list[tuple[str, list[str]]]:
    combined = {}
    if lang in ["Auto", "English"]:
        combined.update(AI_PATTERNS_EN)
    if lang in ["Auto", "Indonesia"]:
        combined.update(AI_PATTERNS_ID)
    return sorted(combined.items(), key=lambda x: len(x[0]), reverse=True)


# =====================================================================
# 2. ENGINE HUMANIZER V2.5.2 (NATURAL & CLEAN)
# =====================================================================

def preserve_case_replace(
    match: re.Match,
    replacements: list[str],
    prefer_length_match: bool = False,
) -> str:
    original = match.group(0)

    # Untuk PDF: pilih pengganti yang paling dekat panjangnya ke teks asli
    # supaya gap spasi setelah redact tidak terlalu mencolok
    if prefer_length_match and len(replacements) > 1:
        orig_len = len(original.strip())
        chosen = min(replacements, key=lambda r: abs(len(r.strip()) - orig_len))
    else:
        chosen = random.choice(replacements)

    if not chosen:
        return ""

    if original[-1].isspace() and not chosen[-1].isspace():
        chosen += original[-1]
    elif not original[-1].isspace() and chosen[-1].isspace():
        chosen = chosen.rstrip()

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
        if len(words) > 32:
            if re.search(r',\s+(which|while|whereas)\s+', sent):
                sent = re.sub(r',\s+which\s+', lambda _: random.choice([". This ", ". That "]), sent, count=1)
                sent = re.sub(r',\s+while\s+', ". Meanwhile, ", sent, count=1)
                sent = re.sub(r',\s+whereas\s+', ". In contrast, ", sent, count=1)
            elif re.search(r',\s+(yang mana|sementara itu)\s+', sent, re.IGNORECASE):
                sent = re.sub(r',\s+yang mana\s+', ". Hal ini ", sent, count=1, flags=re.IGNORECASE)
                sent = re.sub(r',\s+sementara itu\s+', ". Di sisi lain, ", sent, count=1, flags=re.IGNORECASE)
        new_sentences.append(sent)

    return " ".join(new_sentences)


def programmatic_humanize(
    text: str,
    lang: str = "Auto",
    fix_em_dash: bool = True,
    boost_burstiness: bool = False,
    strip_markdown_bold: bool = False,
    prefer_length_match: bool = False,  # aktifkan saat dipakai di PDF
) -> str:
    if not text or not text.strip():
        return text

    result = text

    if strip_markdown_bold:
        result = re.sub(r'\*\*(.*?)\*\*', r'\1', result)

    result = result.replace("\u200b", "")

    if fix_em_dash:
        result = re.sub(
            r'(?<!Abstract)\s*—\s*',
            lambda _: random.choice([", ", " - ", "; "]),
            result
        )

    for pattern, replacements in get_sorted_patterns(lang):
        result = re.sub(
            pattern,
            lambda m, r=replacements: preserve_case_replace(m, r, prefer_length_match),
            result,
            flags=re.IGNORECASE
        )

    result = re.sub(r'\b(dengan|dari|pada|ke|untuk|dalam)\s+\1\b', r'\1', result, flags=re.IGNORECASE)

    if boost_burstiness:
        lines = result.split("\n")
        result = "\n".join(improve_burstiness(line) if len(line.split()) > 32 else line for line in lines)

    return result


# =====================================================================
# 3. VISUAL DIFF & PAPER-STYLE DOCUMENT VIEWER
# =====================================================================

def normalize_paragraphs_for_display(text: str) -> str:
    blocks = re.split(r'\n\s*\n', text.strip())
    cleaned_blocks = []
    for b in blocks:
        lines = [line.strip() for line in b.split('\n') if line.strip()]
        if not lines:
            continue
        if len(lines) <= 2 and all(len(l) < 60 and not l.endswith('.') for l in lines):
            cleaned_blocks.append("\n".join(lines))
        else:
            cleaned_blocks.append(" ".join(lines))
    return "\n\n".join(cleaned_blocks)


def generate_visual_diffs(orig_text: str, new_text: str) -> tuple[str, str, str, int]:
    norm_orig = normalize_paragraphs_for_display(orig_text)
    norm_new = normalize_paragraphs_for_display(new_text)

    orig_tokens = re.findall(r'\n+|\s+|\S+', norm_orig)
    new_tokens = re.findall(r'\n+|\s+|\S+', norm_new)

    matcher = difflib.SequenceMatcher(None, orig_tokens, new_tokens, autojunk=False)

    before_parts = []
    after_parts = []
    inline_parts = []
    change_count = 0

    style_del_side = "background-color: #fee2e2; color: #991b1b; font-weight: 600; padding: 1px 5px; border-radius: 4px; border: 1px solid #f87171;"
    style_del = "background-color: #fee2e2; color: #991b1b; text-decoration: line-through; padding: 1px 5px; border-radius: 4px; border: 1px solid #f87171;"
    style_ins = "background-color: #d1fae5; color: #065f46; font-weight: 600; padding: 1px 5px; border-radius: 4px; border: 1px solid #34d399;"

    def _fmt(tokens: list[str]) -> str:
        return html.escape("".join(tokens)).replace("\n\n", "</p><p style='margin-bottom:12px; text-indent:28px; text-align:justify;'>").replace("\n", "<br>")

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            seg = _fmt(orig_tokens[i1:i2])
            before_parts.append(seg)
            after_parts.append(seg)
            inline_parts.append(seg)
        elif tag == "replace":
            orig_seg = _fmt(orig_tokens[i1:i2])
            new_seg = _fmt(new_tokens[j1:j2])
            if orig_seg.strip() or new_seg.strip():
                change_count += 1
            before_parts.append(f'<span style="{style_del_side}">{orig_seg}</span>')
            after_parts.append(f'<span style="{style_ins}">{new_seg}</span>')
            inline_parts.append(f'<span style="{style_del}">{orig_seg}</span> <span style="{style_ins}">{new_seg}</span>')
        elif tag == "delete":
            orig_seg = _fmt(orig_tokens[i1:i2])
            if orig_seg.strip():
                change_count += 1
            before_parts.append(f'<span style="{style_del_side}">{orig_seg}</span>')
            inline_parts.append(f'<span style="{style_del}">{orig_seg}</span>')
        elif tag == "insert":
            new_seg = _fmt(new_tokens[j1:j2])
            if new_seg.strip():
                change_count += 1
            after_parts.append(f'<span style="{style_ins}">{new_seg}</span>')
            inline_parts.append(f'<span style="{style_ins}">{new_seg}</span>')

    paper_css = "height: 540px; overflow-y: auto; padding: 36px 44px; border: 1px solid #d1d5db; border-radius: 8px; background-color: #ffffff; color: #111827; font-family: 'Times New Roman', Times, Georgia, serif; line-height: 1.65; font-size: 1.05rem; box-shadow: 0 4px 12px rgba(0,0,0,0.15);"
    p_open = "<p style='margin-bottom:12px; text-align:justify;'>"

    return (
        f'<div style="{paper_css}">{p_open}{"".join(before_parts)}</p></div>',
        f'<div style="{paper_css}">{p_open}{"".join(after_parts)}</p></div>',
        f'<div style="{paper_css}">{p_open}{"".join(inline_parts)}</p></div>',
        change_count
    )


# =====================================================================
# 4. IN-PLACE DOCX REWRITER
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


def replace_in_paragraph_preserving_runs(
    p,
    patterns_list: list[tuple[str, list[str]]],
    fix_em_dash: bool,
    boost_burstiness: bool,
    strip_markdown_bold: bool
):
    runs = p.runs
    if not runs:
        return

    if strip_markdown_bold:
        for r in runs:
            if "**" in r.text:
                r.text = r.text.replace("**", "")

    if fix_em_dash:
        for r in runs:
            if "—" in r.text and "Abstract" not in p.text:
                r.text = re.sub(r'\s*—\s*', lambda _: random.choice([", ", " - "]), r.text)

    for pattern, replacements in patterns_list:
        while True:
            full_text = "".join(r.text for r in runs)
            match = re.search(pattern, full_text, flags=re.IGNORECASE)
            if not match:
                break

            m_start, m_end = match.span()
            replacement_str = preserve_case_replace(match, replacements)

            if replacement_str.lower() == match.group(0).lower():
                break

            char_idx = 0
            affected_runs = []
            for r_idx, r in enumerate(runs):
                r_len = len(r.text)
                r_start = char_idx
                r_end = char_idx + r_len
                if r_end > m_start and r_start < m_end:
                    affected_runs.append((r_idx, r_start, r_end))
                char_idx = r_end

            if not affected_runs:
                break

            first_idx, first_start, _ = affected_runs[0]
            last_idx, last_start, _ = affected_runs[-1]

            if first_idx == last_idx:
                r = runs[first_idx]
                local_s = m_start - first_start
                local_e = m_end - first_start
                r.text = r.text[:local_s] + replacement_str + r.text[local_e:]
            else:
                r_first = runs[first_idx]
                r_last = runs[last_idx]
                local_s = m_start - first_start
                local_e = m_end - last_start

                r_first.text = r_first.text[:local_s] + replacement_str
                for mid_idx, _, _ in affected_runs[1:-1]:
                    runs[mid_idx].text = ""
                r_last.text = r_last.text[local_e:]


def humanize_docx_in_place(
    docx_bytes: bytes,
    lang: str,
    fix_em_dash: bool,
    boost_burstiness: bool,
    strip_markdown_bold: bool
) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    seen_elements = set()
    patterns_list = get_sorted_patterns(lang)

    for p in iter_all_docx_paragraphs(doc):
        elem_id = id(p._element)
        if elem_id in seen_elements:
            continue
        seen_elements.add(elem_id)

        orig_text = p.text
        if not orig_text or len(orig_text.strip()) < 3:
            continue
        if "@" in orig_text and len(orig_text.split()) < 12:
            continue

        replace_in_paragraph_preserving_runs(
            p,
            patterns_list=patterns_list,
            fix_em_dash=fix_em_dash,
            boost_burstiness=boost_burstiness,
            strip_markdown_bold=strip_markdown_bold
        )

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# =====================================================================
# 5. IN-PLACE PDF REWRITER (v2.5.2 — fixes bold, spacing, font mapping)
# =====================================================================

def _map_pdf_font_to_base14(font_name: str, flags: int) -> str:
    """
    Peta font PDF ke nama Base-14 terdekat.
    Fix: case-insensitive, lebih banyak serif families, 'italic' bukan 'It' yang ambigu.
    """
    fn = (font_name or "").lower()
    is_bold   = bool(flags & 16) or any(w in fn for w in ["bold", "black", "heavy"])
    is_italic = bool(flags & 2)  or any(w in fn for w in ["italic", "oblique"])
    is_serif  = bool(flags & 4)  or any(
        s in fn for s in [
            "times", "serif", "roman", "cambria", "georgia", "palatino",
            "nimbus", "garamond", "bookman", "charter", "constantia",
            "minion", "utopia", "caslon",
        ]
    )
    if is_serif:
        if is_bold and is_italic: return "tibi"
        if is_bold:               return "tibo"
        if is_italic:             return "tiit"
        return "tiro"
    else:
        if is_bold and is_italic: return "hebi"
        if is_bold:               return "hebo"
        if is_italic:             return "heit"
        return "helv"


def _int_color_to_rgb(color_int: int) -> tuple:
    r = ((color_int >> 16) & 0xFF) / 255.0
    g = ((color_int >> 8)  & 0xFF) / 255.0
    b = (color_int         & 0xFF) / 255.0
    return (r, g, b)


def _try_register_page_font(
    page,
    doc,
    span_font: str,
    span_flags: int,
    page_font_index: dict,
    registered: dict,
) -> str:
    """
    Coba daftarkan font asli yang ter-embed di PDF, kembalikan nama untuk insert_text.
    Fallback ke Base-14 jika font tidak ter-embed atau registrasi gagal.

    page_font_index : {cleaned_basefont_lower: xref}  — dari page.get_fonts()
    registered      : {cleaned_basefont_lower: fontname}  — cache per halaman
    """
    key = (span_font or "").lower()

    if key in registered:
        return registered[key]

    xref = page_font_index.get(key)
    if xref:
        try:
            font_info = doc.extract_font(xref)
            # font_info: (name, ext, type, content, ...) — content di index 3
            buf = font_info[3] if font_info and len(font_info) > 3 else b""
            if buf:
                reg_name = f"EF{xref}"  # nama unik per font xref
                page.insert_font(fontname=reg_name, fontbuffer=buf)
                registered[key] = reg_name
                return reg_name
        except Exception:
            pass  # jatuh ke fallback Base-14

    fallback = _map_pdf_font_to_base14(span_font, span_flags)
    registered[key] = fallback
    return fallback


def humanize_pdf_in_place(
    pdf_bytes: bytes,
    lang: str,
    fix_em_dash: bool,
    strip_markdown_bold: bool
) -> tuple[bytes, bytes]:
    doc_clean   = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc_preview = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_idx in range(len(doc_clean)):
        page_c = doc_clean[page_idx]
        page_p = doc_preview[page_idx]

        # Bangun index font per halaman: {cleaned_basefont_lower: xref}
        # "ABCDEF+Cambria" → key = "cambria"
        page_font_index: dict[str, int] = {}
        for item in page_c.get_fonts(full=True):
            xref, _ext, _ftype, basefont, _res_name, *_ = item
            if xref and basefont:
                clean = basefont.split("+")[-1].lower()
                if clean and clean not in page_font_index:
                    page_font_index[clean] = xref

        # Cache registrasi font per halaman (terpisah antara doc_clean dan doc_preview)
        registered_c: dict[str, str] = {}
        registered_p: dict[str, str] = {}

        blocks = page_c.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        edits: list[dict] = []

        for block in blocks:
            if block["type"] != 0:
                continue
            for line in block.get("lines", []):
                for span in line["spans"]:
                    orig_text = span["text"]
                    if not orig_text or len(orig_text.strip()) < 3:
                        continue

                    new_text = programmatic_humanize(
                        orig_text,
                        lang=lang,
                        fix_em_dash=fix_em_dash,
                        boost_burstiness=False,
                        strip_markdown_bold=strip_markdown_bold,
                        prefer_length_match=True,  # kurangi gap spasi di PDF
                    )

                    if new_text == orig_text:
                        continue

                    rect       = fitz.Rect(span["bbox"])
                    origin     = span.get("origin", (rect.x0, rect.y1 - 2.0))
                    span_font  = span.get("font", "")
                    span_flags = span.get("flags", 0)
                    orig_size  = span["size"]

                    # Fix Bug 1: gunakan font asli yang ter-embed supaya berat huruf sama
                    fontname = _try_register_page_font(
                        page_c, doc_clean, span_font, span_flags,
                        page_font_index, registered_c
                    )

                    # Fix Bug 2 & 4: scale berbasis rasio karakter, minimum 0.92 (bukan 0.80)
                    orig_len = max(len(orig_text.strip()), 1)
                    new_len  = max(len(new_text.strip()), 1)
                    scale    = min(1.0, max(0.92, orig_len / new_len))

                    edits.append({
                        "rect":       rect,
                        "origin":     origin,
                        "text":       new_text,
                        "fontsize":   orig_size * scale,
                        "fontname":   fontname,
                        "color":      _int_color_to_rgb(span.get("color", 0)),
                        "span_font":  span_font,
                        "span_flags": span_flags,
                    })

        if not edits:
            continue

        # Daftarkan font untuk page_p SEBELUM apply_redactions dipanggil
        for edit in edits:
            edit["fontname_p"] = _try_register_page_font(
                page_p, doc_preview, edit["span_font"], edit["span_flags"],
                page_font_index, registered_p
            )

        # Fix Bug 5: fill=None (bukan fill=False) untuk "tidak ada background fill"
        for edit in edits:
            page_c.add_redact_annot(edit["rect"], text="", fill=None)
            page_p.add_redact_annot(edit["rect"], text="", fill=(0.82, 0.98, 0.89))

        page_c.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        page_p.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        for edit in edits:
            try:
                page_c.insert_text(
                    edit["origin"], edit["text"],
                    fontname=edit["fontname"],   fontsize=edit["fontsize"],
                    color=edit["color"]
                )
                page_p.insert_text(
                    edit["origin"], edit["text"],
                    fontname=edit["fontname_p"], fontsize=edit["fontsize"],
                    color=(0.02, 0.37, 0.27)
                )
            except Exception:
                # Fallback ke Base-14 jika font ter-embed punya subset encoding bermasalah
                b14 = _map_pdf_font_to_base14(edit["span_font"], edit["span_flags"])
                try:
                    page_c.insert_text(
                        edit["origin"], edit["text"],
                        fontname=b14, fontsize=edit["fontsize"],
                        color=edit["color"]
                    )
                    page_p.insert_text(
                        edit["origin"], edit["text"],
                        fontname=b14, fontsize=edit["fontsize"],
                        color=(0.02, 0.37, 0.27)
                    )
                except Exception:
                    pass  # skip span yang benar-benar tidak bisa diinsert

    buf_clean = io.BytesIO()
    doc_clean.save(buf_clean, garbage=3, deflate=True)
    buf_clean.seek(0)
    doc_clean.close()

    buf_preview = io.BytesIO()
    doc_preview.save(buf_preview, garbage=3, deflate=True)
    buf_preview.seek(0)
    doc_preview.close()

    return buf_clean.getvalue(), buf_preview.getvalue()


def render_pdf_pages_as_images(pdf_bytes: bytes, dpi: int = 140) -> list[bytes]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        images.append(pix.tobytes("png"))
    doc.close()
    return images


# =====================================================================
# 6. EKSTRAKSI TEKS & ANALISIS METRIK
# =====================================================================

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
    return full_text, len(re.findall(r'\S+', full_text.replace("—", " ").replace("–", " ")))


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> tuple[str, int]:
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_text = []
    for page in pdf_doc:
        blocks = page.get_text("blocks", sort=True)
        page_paras = []
        for b in blocks:
            if b[6] == 0:
                clean_block = re.sub(r'\s*\n\s*', ' ', b[4].strip())
                if clean_block:
                    page_paras.append(clean_block)
        pages_text.append("\n\n".join(page_paras))
    full_text = "\n\n".join(pages_text)
    pdf_doc.close()
    return full_text, len(re.findall(r'\S+', full_text))


def extract_text_from_file(uploaded_file) -> tuple[str, int, bytes, str]:
    filename = uploaded_file.name.lower()
    raw_bytes = uploaded_file.read()
    ext = os.path.splitext(filename)[1]

    if ext == ".docx":
        full_text, word_count = extract_full_text_from_docx_bytes(raw_bytes)
        return full_text, word_count, raw_bytes, ext
    elif ext == ".pdf":
        full_text, word_count = extract_text_from_pdf_bytes(raw_bytes)
        return full_text, word_count, raw_bytes, ext
    else:
        full_text = raw_bytes.decode("utf-8", errors="ignore")
        return full_text, len(re.findall(r'\S+', full_text)), raw_bytes, ext


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
    risk_score = min(100, int(phrase_density * 5 + max(0, (9.0 - burstiness) * 3.5) + min(25, em_dashes * 2.5)))

    return {
        "ai_hits": len(detected_phrases),
        "em_dashes": em_dashes,
        "burstiness": burstiness,
        "risk_score": risk_score,
        "detected_phrases": sorted(list(set(m.strip() for m in detected_phrases))),
    }


# =====================================================================
# 7. ANTARMUKA WEB UI (STREAMLIT)
# =====================================================================

st.set_page_config(
    page_title="Turnitout Server - AI Humanizer & De-AI Pattern",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Self-Hosted AI Humanizer & Similarity Evader (v2.5.2)")
st.caption(
    "100% Programmatic Engine berbasis standar **Humanizer v2.5.2** & **Wikipedia Signs of AI Writing** "
    "— Kosakata natural, mengalir, dan menjaga tata letak dokumen asli."
)

with st.sidebar:
    st.header("⚙️ Pengaturan Filter")
    lang_choice = st.selectbox("Bahasa Dokumen:", ["Auto", "English", "Indonesia"])

    st.subheader("Aturan De-AI Pattern")
    opt_em_dash = st.checkbox("Ganti Em-Dash (—) Berlebih", value=True)
    opt_burstiness = st.checkbox(
        "Variasi Kalimat Panjang (Burstiness)",
        value=False,
        help="Aktifkan jika ingin memecah kalimat >32 kata."
    )
    opt_strip_bold = st.checkbox("Bersihkan Simbol Markdown (**)", value=False)

tab_upload, tab_paste = st.tabs(["📂 Upload Dokumen (Word / PDF / TXT)", "✍️ Paste Teks Langsung"])

input_text = ""
raw_file_bytes = None
total_words_detected = 0
uploaded_file_stem = "Humanized_Document"
uploaded_file_ext = ".txt"

with tab_upload:
    uploaded_file = st.file_uploader(
        "Pilih file dokumen (.docx, .pdf, .txt, .md, .tex)",
        type=["docx", "pdf", "txt", "md", "tex"]
    )
    if uploaded_file is not None:
        try:
            input_text, total_words_detected, raw_file_bytes, uploaded_file_ext = extract_text_from_file(uploaded_file)
            uploaded_file_stem = os.path.splitext(uploaded_file.name)[0]
            st.success(f"Berhasil membaca **{uploaded_file.name}** (Total: **{total_words_detected:,} kata**).")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

with tab_paste:
    pasted_text = st.text_area(
        "Masukkan atau tempel teks di sini:",
        height=260,
        placeholder="Tempel teks laporan atau jurnal di sini..."
    )
    if pasted_text.strip():
        input_text = pasted_text
        total_words_detected = len(re.findall(r'\S+', input_text))
        raw_file_bytes = None
        uploaded_file_ext = ".txt"

if st.button("🚀 Proses & Humanize Dokumen", type="primary", use_container_width=True):
    if not input_text.strip():
        st.warning("Silakan unggah dokumen atau tempelkan teks terlebih dahulu.")
    else:
        with st.spinner("Memparafrase pola AI menggunakan standar Humanizer v2.5.2..."):
            before_metrics = analyze_ai_metrics(input_text)

            output_docx_bytes = None
            output_pdf_bytes = None
            orig_pdf_images = None
            preview_pdf_images = None

            try:
                if raw_file_bytes is not None and uploaded_file_ext == ".docx":
                    output_docx_bytes = humanize_docx_in_place(
                        raw_file_bytes, lang=lang_choice, fix_em_dash=opt_em_dash,
                        boost_burstiness=opt_burstiness, strip_markdown_bold=opt_strip_bold
                    )
                    output_text, after_word_count = extract_full_text_from_docx_bytes(output_docx_bytes)

                elif raw_file_bytes is not None and uploaded_file_ext == ".pdf":
                    output_pdf_bytes, preview_pdf_bytes = humanize_pdf_in_place(
                        raw_file_bytes, lang=lang_choice, fix_em_dash=opt_em_dash,
                        strip_markdown_bold=opt_strip_bold
                    )
                    output_text, after_word_count = extract_text_from_pdf_bytes(output_pdf_bytes)
                    orig_pdf_images = render_pdf_pages_as_images(raw_file_bytes)
                    preview_pdf_images = render_pdf_pages_as_images(preview_pdf_bytes)

                else:
                    output_text = programmatic_humanize(
                        input_text, lang=lang_choice, fix_em_dash=opt_em_dash,
                        boost_burstiness=opt_burstiness, strip_markdown_bold=opt_strip_bold
                    )
                    output_docx_bytes = create_plain_docx_bytes(output_text)
                    after_word_count = len(re.findall(r'\S+', output_text.replace("—", " ")))

                after_metrics = analyze_ai_metrics(output_text)
                html_before, html_after, html_inline, total_changes = generate_visual_diffs(input_text, output_text)

                st.session_state["humanize_result"] = {
                    "input_text": input_text, "output_text": output_text,
                    "total_words_detected": total_words_detected, "after_word_count": after_word_count,
                    "before_metrics": before_metrics, "after_metrics": after_metrics,
                    "html_before": html_before, "html_after": html_after, "html_inline": html_inline,
                    "total_changes": total_changes, "output_docx_bytes": output_docx_bytes,
                    "output_pdf_bytes": output_pdf_bytes, "orig_pdf_images": orig_pdf_images,
                    "preview_pdf_images": preview_pdf_images, "uploaded_file_ext": uploaded_file_ext,
                    "uploaded_file_stem": uploaded_file_stem,
                }
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")

if "humanize_result" in st.session_state:
    res = st.session_state["humanize_result"]
    before_metrics = res["before_metrics"]
    after_metrics = res["after_metrics"]

    st.subheader("📊 Hasil Analisis Pola AI")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total Kata", f"{res['after_word_count']:,}", delta=f"{res['after_word_count'] - res['total_words_detected']} kata")
    m2.metric("Titik Kata Diubah", f"{res['total_changes']:,}", delta="Terparafrase", delta_color="normal")
    m3.metric("Skor Indikator AI", f"{after_metrics['risk_score']}%", delta=f"{after_metrics['risk_score'] - before_metrics['risk_score']}%", delta_color="inverse")
    m4.metric("Frasa Klise AI", after_metrics["ai_hits"], delta=after_metrics["ai_hits"] - before_metrics["ai_hits"], delta_color="inverse")
    m5.metric("Skor Burstiness", after_metrics["burstiness"], delta=round(after_metrics["burstiness"] - before_metrics["burstiness"], 2))
    m6.metric("Em-Dash (—)", after_metrics["em_dashes"], delta=after_metrics["em_dashes"] - before_metrics["em_dashes"], delta_color="inverse")

    if before_metrics["detected_phrases"]:
        with st.expander(f"🔍 Lihat Daftar Kata/Frasa AI yang Terdeteksi & Diganti ({len(before_metrics['detected_phrases'])} jenis)"):
            st.write(", ".join([f"`{p}`" for p in before_metrics["detected_phrases"]]))

    st.subheader("🎨 Visual Diff & Tampilan Dokumen")

    if res["uploaded_file_ext"] == ".pdf" and res["orig_pdf_images"] is not None:
        tab_pdf_layout, diff_tab_side, diff_tab_inline, diff_tab_raw = st.tabs([
            "📄 Layout Halaman PDF Asli vs Hasil (100% Identik)",
            "⚡ Kertas Visual Diff (Side-by-Side Highlight)",
            "🔀 Gabungan (Inline Diff Coret & Ganti)",
            "📝 Teks Polos"
        ])

        with tab_pdf_layout:
            st.caption("Perbandingan tampilan halaman PDF asli dengan hasil humanizer (kata yang diganti ditandai warna hijau di sebelah kanan).")
            num_pages = len(res["orig_pdf_images"])
            selected_page = st.slider("Pilih Halaman PDF:", 1, num_pages, 1) if num_pages > 1 else 1
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                st.markdown(f"#### 📄 Halaman {selected_page} — PDF Asli")
                st.image(res["orig_pdf_images"][selected_page - 1], use_container_width=True)
            with p_col2:
                st.markdown(f"#### ✨ Halaman {selected_page} — PDF Hasil Humanizer")
                st.image(res["preview_pdf_images"][selected_page - 1], use_container_width=True)
    else:
        diff_tab_side, diff_tab_inline, diff_tab_raw = st.tabs([
            "⚡ Kertas Visual Diff (Side-by-Side Highlight)",
            "🔀 Gabungan (Inline Diff Coret & Ganti)",
            "📝 Teks Polos"
        ])

    with diff_tab_side:
        st.markdown(
            """
            <div style="display:flex; gap:18px; align-items:center; margin-bottom:10px; font-size:0.9rem;">
                <span><b>Legenda Kertas Dokumen:</b></span>
                <span style="background:#fee2e2; color:#991b1b; padding:2px 8px; border-radius:4px; border:1px solid #f87171; font-weight:600;">
                    🔴 Kata Asli (Dihapus)
                </span>
                <span style="background:#d1fae5; color:#065f46; padding:2px 8px; border-radius:4px; border:1px solid #34d399; font-weight:600;">
                    🟢 Kata Baru (Humanizer v2.5.2 Natural)
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )
        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown(f"#### 📄 Teks Asli ({res['total_words_detected']:,} kata)")
            st.markdown(res["html_before"], unsafe_allow_html=True)
        with col_right:
            st.markdown(f"#### ✨ Hasil Humanizer ({res['after_word_count']:,} kata)")
            st.markdown(res["html_after"], unsafe_allow_html=True)

    with diff_tab_inline:
        st.markdown("#### 🔍 Detail Perubahan Langsung")
        st.markdown(res["html_inline"], unsafe_allow_html=True)

    with diff_tab_raw:
        raw_left, raw_right = st.columns(2)
        with raw_left:
            st.markdown("#### 📄 Teks Asli (Raw)")
            st.text_area("Before Raw", value=res["input_text"], height=420, label_visibility="collapsed")
        with raw_right:
            st.markdown("#### ✨ Hasil Humanizer (Raw)")
            st.text_area("After Raw", value=res["output_text"], height=420, label_visibility="collapsed")

    st.divider()
    st.subheader("💾 Unduh Dokumen Hasil")
    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        if res["uploaded_file_ext"] == ".pdf" and res["output_pdf_bytes"] is not None:
            st.download_button(
                label="📥 Download Humanized PDF (.pdf)",
                data=res["output_pdf_bytes"],
                file_name=f"Humanized_{res['uploaded_file_stem']}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        elif res["output_docx_bytes"] is not None:
            st.download_button(
                label="📥 Download Humanized Word (.docx)",
                data=res["output_docx_bytes"],
                file_name=f"Humanized_{res['uploaded_file_stem']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    with dl_col2:
        out_ext = res["uploaded_file_ext"] if res["uploaded_file_ext"] in [".txt", ".md", ".tex"] else ".txt"
        st.download_button(
            label=f"📥 Download File Teks ({out_ext})",
            data=res["output_text"].encode("utf-8"),
            file_name=f"Humanized_{res['uploaded_file_stem']}{out_ext}",
            mime="text/plain",
            use_container_width=True
        )