# ===== 必要ライブラリ =====
import streamlit as st
import pandas as pd
from deep_translator import GoogleTranslator
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import io

# ==== 設定 ====
DEFAULT_LANGUAGES = ["en", "vi", "ne", "my", "zh-CN", "zh-TW"]  # デフォルト翻訳言語

# スライドサイズ（EMU）
SLIDE_WIDTH = 914400 * 10
SLIDE_HEIGHT = 914400 * 7.5


# ==== テキストボックス関数 ====
def add_textbox(slide, text, y_percent, font_size):
    textbox = slide.shapes.add_textbox(
        left=0,
        top=int(SLIDE_HEIGHT * y_percent),
        width=SLIDE_WIDTH,
        height=int(SLIDE_HEIGHT * 0.15),
    )
    tf = textbox.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    p.alignment = PP_ALIGN.CENTER


# ==== 横線 ====
def add_center_line(slide):
    slide.shapes.add_shape(
        1,  # msoShapeRectangleを細長く
        left=0,
        top=int(SLIDE_HEIGHT * 0.5),
        width=SLIDE_WIDTH,
        height=Pt(2),
    )


# ==== 翻訳 ====
def translate_word(word, lang):
    try:
        return GoogleTranslator(source="ja", target=lang).translate(word)
    except Exception:
        return f"[Error:{lang}]"


# ==== PPT作成関数 ====
def create_ppt_from_vocab(df, target_languages, font_sizes):
    prs = Presentation()
    for _, row in df.iterrows():
        word = str(row.iloc[0]).strip()
        ruby = str(row.iloc[1]).strip()

        # --- Slide1 ---
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        add_textbox(slide1, word, 0.15, font_sizes["kanji"])
        add_center_line(slide1)

        # --- Slide2 ---
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        add_textbox(slide2, word, 0.15, font_sizes["kanji"])
        add_textbox(slide2, ruby, 0.52, font_sizes["hiragana"])

        translations = [translate_word(word, lang) for lang in target_languages]
        add_textbox(
            slide2, "   ".join(translations), 0.68, font_sizes["translation"]
        )
        add_center_line(slide2)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io


# ================= Streamlit UI =================
st.title("📚 漢字フラッシュカード自動生成ツール")

uploaded_file = st.file_uploader("語彙リストをアップロード (CSV または Excel)", type=["csv", "xlsx"])

# 設定パネル
with st.sidebar:
    st.header("⚙️ 設定")
    langs = st.text_input("翻訳言語（カンマ区切り）", ",".join(DEFAULT_LANGUAGES))
    target_languages = [lang.strip() for lang in langs.split(",") if lang.strip()]

    font_kanji = st.number_input("フォントサイズ（漢字）", 40, 200, 84)
    font_hira = st.number_input("フォントサイズ（ひらがな）", 20, 150, 70)
    font_trans = st.number_input("フォントサイズ（翻訳語）", 15, 100, 35)

    font_sizes = {
        "kanji": font_kanji,
        "hiragana": font_hira,
        "translation": font_trans,
    }

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.write("✅ 読み込んだデータプレビュー")
    st.dataframe(df.head())

    if st.button("PPTXを生成"):
        pptx_file = create_ppt_from_vocab(df, target_languages, font_sizes)
        st.success("PPTXファイルを生成しました！")

        st.download_button(
            label="📥 ダウンロード",
            data=pptx_file,
            file_name="KanjiFlashcards.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
