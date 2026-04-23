"""
YouTube Transcript Extractor
----------------------------
Colle l'URL d'une chaîne (ou d'une playlist), choisis le nombre de vidéos,
le script récupère les transcriptions et te génère un .docx.
"""

import os
import re
import glob
import tempfile
from io import BytesIO

import streamlit as st
import yt_dlp
from youtube_transcript_api import YouTubeTranscriptApi
from youtube_transcript_api._errors import TranscriptsDisabled, NoTranscriptFound
from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================================
# CONFIG STREAMLIT
# ============================================================================
st.set_page_config(
    page_title="YouTube Transcript Extractor",
    page_icon="🎬",
    layout="wide",
)

st.title("🎬 YouTube Transcript Extractor")
st.caption("Colle une URL de chaîne ou de playlist → récupère les N dernières transcriptions → télécharge un .docx")


# ============================================================================
# HELPERS
# ============================================================================
def normalize_channel_url(url: str) -> str:
    """Force le suffixe /videos pour récupérer les uploads récents d'une chaîne."""
    url = url.strip().rstrip("/")
    # Si c'est une playlist ou une vidéo, on touche pas
    if "playlist" in url or "watch?v=" in url or "/videos" in url:
        return url
    # Sinon on force /videos
    return url + "/videos"


def list_videos(url: str, n: int):
    """Liste les N dernières vidéos d'une chaîne ou playlist via yt-dlp."""
    ydl_opts = {
        "quiet": True,
        "no_warnings": True,
        "extract_flat": True,   # On ne télécharge pas, on liste juste
        "playlistend": n,
        "skip_download": True,
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
    source_title = info.get("title") or "YouTube Transcripts"
    return videos, source_title


def parse_vtt(path: str) -> str:
    """Extrait le texte propre d'un fichier .vtt (sans timestamps ni balises)."""
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    lines = []
    for line in content.split("\n"):
        line = line.strip()
        if not line:
            continue
        if "-->" in line:
            continue
        if line.startswith(("WEBVTT", "Kind:", "Language:", "NOTE")):
            continue
        if re.match(r"^\d+$", line):  # index numérique
            continue
        # Retirer les balises <c>, <00:00:00.000>, etc.
        line = re.sub(r"<[^>]+>", "", line)
        if line and (not lines or lines[-1] != line):
            lines.append(line)
    return " ".join(lines)


def fetch_transcript_api(video_id: str, lang_pref: list, include_auto: bool):
    """Méthode 1 : youtube-transcript-api (rapide, mais parfois bloqué par YT)."""
    try:
        transcripts = YouTubeTranscriptApi.list_transcripts(video_id)
    except (TranscriptsDisabled, NoTranscriptFound):
        return None, None
    except Exception:
        return None, None

    # D'abord : manuels dans les langues préférées
    for lang in lang_pref:
        try:
            t = transcripts.find_manually_created_transcript([lang])
            text = " ".join(x["text"] for x in t.fetch())
            return text, f"{lang} (manuel)"
        except Exception:
            continue

    # Ensuite : auto-générés si autorisés
    if include_auto:
        for lang in lang_pref:
            try:
                t = transcripts.find_generated_transcript([lang])
                text = " ".join(x["text"] for x in t.fetch())
                return text, f"{lang} (auto)"
            except Exception:
                continue

    # Dernier recours : n'importe quelle langue disponible
    try:
        for t in transcripts:
            if t.is_generated and not include_auto:
                continue
            try:
                text = " ".join(x["text"] for x in t.fetch())
                tag = "auto" if t.is_generated else "manuel"
                return text, f"{t.language_code} ({tag})"
            except Exception:
                continue
    except Exception:
        pass

    return None, None


def fetch_transcript_ytdlp(video_id: str, lang_pref: list, include_auto: bool):
    """Méthode 2 (fallback) : yt-dlp télécharge les .vtt directement."""
    with tempfile.TemporaryDirectory() as tmp:
        opts = {
            "quiet": True,
            "no_warnings": True,
            "skip_download": True,
            "writesubtitles": True,
            "writeautomaticsub": include_auto,
            "subtitleslangs": lang_pref + ["en"],  # toujours fallback en anglais
            "subtitlesformat": "vtt",
            "outtmpl": os.path.join(tmp, "%(id)s.%(ext)s"),
        }
        try:
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([f"https://www.youtube.com/watch?v={video_id}"])
        except Exception:
            return None, None

        # Cherche d'abord dans les langues préférées
        for lang in lang_pref:
            for f in glob.glob(os.path.join(tmp, f"*.{lang}*.vtt")):
                return parse_vtt(f), f"{lang} (yt-dlp)"

        # Sinon n'importe quel .vtt
        vtts = glob.glob(os.path.join(tmp, "*.vtt"))
        if vtts:
            return parse_vtt(vtts[0]), "auto (yt-dlp)"

    return None, None


def get_transcript(video_id: str, lang_pref: list, include_auto: bool):
    """Essaie l'API puis yt-dlp en fallback."""
    text, lang = fetch_transcript_api(video_id, lang_pref, include_auto)
    if text:
        return text, lang
    return fetch_transcript_ytdlp(video_id, lang_pref, include_auto)


def build_docx(source_title: str, channel_url: str, results: list) -> BytesIO:
    """Construit le document Word final."""
    doc = Document()

    # Style de base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # En-tête
    doc.add_heading(source_title, level=0)
    doc.add_paragraph(f"Source : {channel_url}")
    doc.add_paragraph(f"Nombre de vidéos : {len(results)}")
    doc.add_paragraph("")

    # Une section par vidéo
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
            run = p.add_run("❌ Pas de transcription disponible")
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
    include_auto = st.checkbox("Inclure les sous-titres auto", value=True)

st.markdown("---")

if st.button("🚀 Extraire les transcriptions", type="primary", use_container_width=True):
    if not channel_url:
        st.error("Merci de fournir une URL")
        st.stop()
    if not lang_pref:
        st.error("Choisis au moins une langue")
        st.stop()

    # --- Étape 1 : lister les vidéos
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

    # --- Étape 2 : récupérer les transcriptions
    progress = st.progress(0.0)
    status = st.empty()
    results = []

    for i, v in enumerate(videos):
        status.info(f"📝 Transcription {i+1}/{len(videos)} — {v['title']}")
        text, lang = get_transcript(v["id"], lang_pref, include_auto)
        results.append({
            "id": v["id"],
            "title": v["title"],
            "url": v["url"],
            "transcript": text,
            "lang": lang,
        })
        progress.progress((i + 1) / len(videos))

    status.empty()
    progress.empty()

    # --- Étape 3 : construire le docx
    ok_count = sum(1 for r in results if r["transcript"])
    total_words = sum(len(r["transcript"].split()) for r in results if r["transcript"])

    c1, c2, c3 = st.columns(3)
    c1.metric("Transcriptions OK", f"{ok_count}/{len(results)}")
    c2.metric("Total mots", f"{total_words:,}".replace(",", " "))
    c3.metric("Vidéos sans sous-titres", len(results) - ok_count)

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

    # --- Aperçu
    with st.expander("👁️ Aperçu des transcriptions"):
        for r in results:
            st.markdown(f"### [{r['title']}]({r['url']})")
            if r["transcript"]:
                st.caption(f"Langue : {r['lang']} · {len(r['transcript'].split())} mots")
                preview = r["transcript"][:2000] + ("…" if len(r["transcript"]) > 2000 else "")
                st.text_area(" ", preview, height=150, key=r["id"], label_visibility="collapsed")
            else:
                st.warning("Pas de transcription disponible")
            st.markdown("---")
