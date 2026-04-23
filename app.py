"""
YouTube Transcript Extractor v3
--------------------------------
- yt-dlp uniquement (plus robuste que youtube-transcript-api)
- Liste les sous-titres dispo, construit une liste ordonnée de cibles
- Itère sur les cibles jusqu'à ce qu'une marche (fr manuel → fr auto → en manuel → en auto → fallback)
- User-Agent navigateur + retries pour contourner le rate-limiting YouTube
"""

import os
import re
import html
import glob
import tempfile
from io import BytesIO

import streamlit as st
import yt_dlp
from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================================
# CONFIG
# ============================================================================
st.set_page_config(page_title="YouTube Transcript Extractor", page_icon="🎬", layout="wide")
st.title("🎬 YouTube Transcript Extractor")
st.caption("Colle une URL de chaîne / playlist → récupère les N dernières transcriptions → télécharge un .docx")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
)


# ============================================================================
# LISTING VIDÉOS
# ============================================================================
def normalize_channel_url(url: str) -> str:
    url = url.strip().rstrip("/")
    if "playlist" in url or "watch?v=" in url or "/videos" in url:
        return url
    return url + "/videos"


def list_videos(url: str, n: int):
    ydl_opts = {
        "quiet": True,
        "no_warnings": True,
        "extract_flat": True,
        "playlistend": n,
        "skip_download": True,
        "http_headers": {"User-Agent": USER_AGENT},
        "retries": 3,
    }
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)

    entries = info.get("entries") or []
    videos = []
    for e in entries[:n]:
        vid_id = e.get("id")
        if not vid_id:
            continue
        videos.append({
            "id": vid_id,
            "title": e.get("title", "Sans titre"),
            "url": e.get("url") or f"https://www.youtube.com/watch?v={vid_id}",
        })
    return videos, info.get("title") or "YouTube Transcripts"


# ============================================================================
# MATCHING DE LANGUE
# ============================================================================
def match_lang(pref: str, pool):
    """'fr' matche 'fr', 'fr-FR', 'fr-CA'. 'en' matche 'en', 'en-US', 'en-orig'."""
    pref_lower = pref.lower()
    pool = list(pool)
    # 1. Exact
    for lang in pool:
        if lang.lower() == pref_lower:
            return lang
    # 2. Préfixe avec tiret/underscore
    for lang in pool:
        ll = lang.lower()
        if ll.startswith(pref_lower + "-") or ll.startswith(pref_lower + "_"):
            return lang
    # 3. Base du code
    for lang in pool:
        base = lang.lower().split("-")[0].split("_")[0]
        if base == pref_lower:
            return lang
    return None


# ============================================================================
# PARSING VTT
# ============================================================================
def parse_vtt(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception:
        return ""

    lines = []
    for line in content.split("\n"):
        line = line.strip()
        if not line:
            continue
        if "-->" in line:
            continue
        if line.startswith(("WEBVTT", "Kind:", "Language:", "NOTE", "STYLE", "REGION")):
            continue
        if re.match(r"^\d+$", line):
            continue
        line = re.sub(r"<[^>]+>", "", line)
        line = html.unescape(line).strip()
        if line and (not lines or lines[-1] != line):
            lines.append(line)
    return " ".join(lines)


# ============================================================================
# FETCH — listing préalable + loop de cibles avec fallback
# ============================================================================
def build_targets(manual_subs, auto_subs, lang_pref, include_auto):
    """Construit la liste ordonnée (lang, type) des cibles à essayer."""
    targets = []
    seen = set()

    def add(lang, t):
        key = (lang.lower(), t)
        if key not in seen:
            seen.add(key)
            targets.append((lang, t))

    # 1. Manuels dans langues préférées
    for pref in lang_pref:
        m = match_lang(pref, manual_subs)
        if m:
            add(m, "manuel")

    # 2. Auto dans langues préférées
    if include_auto:
        for pref in lang_pref:
            m = match_lang(pref, auto_subs)
            if m:
                add(m, "auto")

    # 3. Anglais en fallback (manuel)
    for lang in manual_subs:
        if lang.lower().startswith("en"):
            add(lang, "manuel")

    # 4. Anglais en fallback (auto)
    if include_auto:
        for lang in auto_subs:
            if lang.lower().startswith("en"):
                add(lang, "auto")

    # 5. Premier manuel dispo (langue originale quelle qu'elle soit)
    for lang in sorted(manual_subs):
        add(lang, "manuel")

    # 6. Premier auto dispo
    if include_auto:
        for lang in sorted(auto_subs):
            add(lang, "auto")

    return targets


def try_download_one(video_url, target_lang, target_type, tmp):
    """Tente de télécharger UNE cible. Retourne le chemin du .vtt ou None."""
    dl_opts = {
        "quiet": True,
        "no_warnings": True,
        "skip_download": True,
        "writesubtitles": (target_type == "manuel"),
        "writeautomaticsub": (target_type == "auto"),
        "subtitleslangs": [target_lang],
        "subtitlesformat": "vtt",
        "outtmpl": os.path.join(tmp, "%(id)s.%(ext)s"),
        "http_headers": {"User-Agent": USER_AGENT},
        "retries": 3,
        "extractor_retries": 3,
    }
    try:
        with yt_dlp.YoutubeDL(dl_opts) as ydl:
            ydl.download([video_url])
    except Exception as e:
        return None, type(e).__name__

    vtt_files = glob.glob(os.path.join(tmp, "*.vtt"))
    if not vtt_files:
        return None, "no-vtt-file"
    # Prendre le plus récent
    vtt_files.sort(key=os.path.getmtime, reverse=True)
    return vtt_files[0], None


def fetch_transcript(video_id: str, lang_pref: list, include_auto: bool):
    """Retourne (texte, label_langue, message_erreur)."""
    video_url = f"https://www.youtube.com/watch?v={video_id}"

    # 1. Inventaire des sous-titres
    try:
        with yt_dlp.YoutubeDL({
            "quiet": True,
            "no_warnings": True,
            "skip_download": True,
            "http_headers": {"User-Agent": USER_AGENT},
            "retries": 3,
            "extractor_retries": 3,
        }) as ydl:
            info = ydl.extract_info(video_url, download=False)
    except Exception as e:
        return None, None, f"extract_info : {type(e).__name__}"

    manual_subs = list((info.get("subtitles") or {}).keys())
    auto_subs = list((info.get("automatic_captions") or {}).keys())

    if not manual_subs and not auto_subs:
        return None, None, "aucun sous-titre (ni manuel ni auto)"

    # 2. Construire la liste des cibles
    targets = build_targets(manual_subs, auto_subs, lang_pref, include_auto)
    if not targets:
        return None, None, "aucune cible compatible"

    # 3. Itérer jusqu'à succès
    errors = []
    for target_lang, target_type in targets:
        with tempfile.TemporaryDirectory() as tmp:
            vtt_path, err = try_download_one(video_url, target_lang, target_type, tmp)
            if err:
                errors.append(f"{target_lang}/{target_type}:{err}")
                continue
            text = parse_vtt(vtt_path)
            if text:
                return text, f"{target_lang} ({target_type})", None
            errors.append(f"{target_lang}/{target_type}:vtt-vide")

    return None, None, f"échec sur {len(targets)} cibles [{'; '.join(errors[:3])}]"


# ============================================================================
# DOCX
# ============================================================================
def build_docx(source_title: str, channel_url: str, results: list) -> BytesIO:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_heading(source_title, level=0)
    doc.add_paragraph(f"Source : {channel_url}")
    doc.add_paragraph(f"Nombre de vidéos : {len(results)}")
    doc.add_paragraph("")

    for idx, r in enumerate(results, 1):
        doc.add_heading(f"{idx}. {r['title']}", level=1)
        p = doc.add_paragraph()
        p.add_run("URL : ").bold = True
        p.add_run(r["url"])

        if r["transcript"]:
            p = doc.add_paragraph()
            p.add_run("Langue : ").bold = True
            p.add_run(r["lang"] or "inconnue")
            doc.add_paragraph(r["transcript"])
        else:
            p = doc.add_paragraph()
            run = p.add_run(f"❌ Pas de transcription — {r.get('error', 'raison inconnue')}")
            run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        doc.add_paragraph("")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", name)[:60].strip("_") or "transcripts"


# ============================================================================
# UI
# ============================================================================
col1, col2 = st.columns([3, 1])
with col1:
    channel_url = st.text_input(
        "URL de la chaîne ou de la playlist YouTube",
        placeholder="https://www.youtube.com/@NomDeLaChaine",
    )
with col2:
    num_videos = st.number_input("Nombre de vidéos", min_value=1, max_value=50, value=10)

col3, col4 = st.columns([2, 1])
with col3:
    lang_pref = st.multiselect(
        "Langues préférées (ordre prioritaire)",
        options=["fr", "en", "es", "de", "ar", "pt", "it", "ru", "zh", "ja"],
        default=["fr", "en"],
    )
with col4:
    include_auto = st.checkbox("Inclure sous-titres auto", value=True)

st.markdown("---")

if st.button("🚀 Extraire les transcriptions", type="primary", use_container_width=True):
    if not channel_url:
        st.error("Merci de fournir une URL")
        st.stop()
    if not lang_pref:
        st.error("Choisis au moins une langue")
        st.stop()

    # 1. Lister les vidéos
    with st.spinner(f"Récupération des {num_videos} dernières vidéos..."):
        url = normalize_channel_url(channel_url)
        try:
            videos, source_title = list_videos(url, int(num_videos))
        except Exception as e:
            st.error(f"Impossible de lister les vidéos : {e}")
            st.stop()

    if not videos:
        st.error("Aucune vidéo trouvée. Vérifie l'URL.")
        st.stop()

    st.success(f"✅ {len(videos)} vidéos trouvées dans **{source_title}**")

    # 2. Récupérer les transcriptions
    progress = st.progress(0.0)
    status = st.empty()
    results = []

    for i, v in enumerate(videos):
        status.info(f"📝 Transcription {i+1}/{len(videos)} — {v['title']}")
        text, lang, error = fetch_transcript(v["id"], lang_pref, include_auto)
        results.append({
            "id": v["id"],
            "title": v["title"],
            "url": v["url"],
            "transcript": text,
            "lang": lang,
            "error": error,
        })
        progress.progress((i + 1) / len(videos))

    status.empty()
    progress.empty()

    # 3. Métriques
    ok_count = sum(1 for r in results if r["transcript"])
    total_words = sum(len(r["transcript"].split()) for r in results if r["transcript"])

    c1, c2, c3 = st.columns(3)
    c1.metric("Transcriptions OK", f"{ok_count}/{len(results)}")
    c2.metric("Total mots", f"{total_words:,}".replace(",", " "))
    c3.metric("Échecs", len(results) - ok_count)

    # 4. Télécharger
    buf = build_docx(source_title, channel_url, results)
    filename = f"{safe_filename(source_title)}_transcripts.docx"
    st.download_button(
        "📥 Télécharger le .docx",
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

    # 5. Détails des échecs
    failures = [r for r in results if not r["transcript"]]
    if failures:
        with st.expander(f"⚠️ Détails des {len(failures)} échec(s)", expanded=True):
            for r in failures:
                st.markdown(f"- **[{r['title']}]({r['url']})** → `{r.get('error', 'n/a')}`")

    # 6. Aperçu
    with st.expander("👁️ Aperçu des transcriptions"):
        for r in results:
            st.markdown(f"### [{r['title']}]({r['url']})")
            if r["transcript"]:
                st.caption(f"Langue : {r['lang']} · {len(r['transcript'].split())} mots")
                preview = r["transcript"][:2000] + ("…" if len(r["transcript"]) > 2000 else "")
                st.text_area(" ", preview, height=150, key=r["id"], label_visibility="collapsed")
            else:
                st.warning(f"❌ {r.get('error', 'raison inconnue')}")
            st.markdown("---")
