"""
Streamlit GUI to aggregate and visualize tagged sentiment data from BestBuy review CSVs.

2025-06-18 highlights 🆕
------------------------------------------------
1. タブ切り替えで各モデル＆全体集計を閲覧
2. タグ別感情割合のモデル比較＋Excel「Tag_Ratios」シート
3. 日本語タイトルが文字化けしないようフォント自動設定
   - IPAexGothic / Noto Sans CJK JP / Yu Gothic / MS Gothic の順で検出
4. ★モデルごとの総レビュー件数をテーブル＆グラフに表示
5. ★件数ラベルを “棒の最上端（≒100%）” に必ず配置   ← New!
"""

# ---------------------------------------------------------------------------
# Matplotlib 日本語フォント自動セットアップ
# ---------------------------------------------------------------------------
import matplotlib as mpl
import matplotlib.font_manager as fm

_FONT_CANDS = [
    "IPAexGothic",
    "Noto Sans CJK JP",
    "Yu Gothic",
    "MS Gothic",
]


def _pick_jp_font() -> str | None:
    avail = {f.name for f in fm.fontManager.ttflist}
    for cand in _FONT_CANDS:
        if cand in avail:
            return cand
    return None


_SELECTED_FONT = _pick_jp_font()
_MISSING_FONT = _SELECTED_FONT is None
if _SELECTED_FONT:
    mpl.rcParams["font.family"] = _SELECTED_FONT
    mpl.rcParams["axes.unicode_minus"] = False  # − を正しく表示

# ---------------------------------------------------------------------------
# 主要ライブラリ
# ---------------------------------------------------------------------------
import io
from datetime import datetime
from typing import Dict, List

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

# ---------------------------------------------------------------------------
# Streamlit ページ設定
# ---------------------------------------------------------------------------
st.set_page_config(page_title="BestBuy Review Analyzer", layout="centered")

st.title("📊 BestBuy Review Analyzer")
st.markdown(
    "CSV ファイルを複数選択すると、タグ別感情・評価分布などを自動集計します。<br>"
    "モデル名はファイル名から最大 5 トークンまで抽出して表示します！",
    unsafe_allow_html=True,
)

if _MISSING_FONT:
    st.warning(
        "日本語フォントがシステムに見つかりません。グラフが文字化けする場合は "
        "`sudo apt-get install fonts-noto-cjk` などで追加してください。"
    )

# ---------------------------------------------------------------------------
# ファイルアップローダ
# ---------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "解析したい CSV を選択 (複数可)",
    type="csv",
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("左サイドバーまたは上のボタンから CSV をアップロードしてください。")
    st.stop()

# ---------------------------------------------------------------------------
# モデル名抽出ヘルパ
# ---------------------------------------------------------------------------
MAX_TOKENS = 5
STOPWORDS = {
    "eval", "evaluation", "result", "results",
    "analysis", "analyze", "output"
}


def derive_model_name(filename: str) -> str:
    stem = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0].replace("-", "_")
    tokens = [t for t in stem.split("_") if t and t.lower() not in STOPWORDS]
    short = "_".join(tokens[:MAX_TOKENS]) or stem
    return short[:30]  # Excel シート名制限対策


# ---------------------------------------------------------------------------
# ファイル読込 & 前処理
# ---------------------------------------------------------------------------
TAG_COLUMNS: List[str] = [
    "SoundQuality", "Music", "Movies", "Surround",
    "Dialogue", "Bass", "App", "Setup", "Design",
]
SENTIMENTS = ["Positive", "Neutral", "Negative"]

file_dfs: Dict[str, pd.DataFrame] = {}
model_names: Dict[str, str] = {}
dup_counter: Dict[str, int] = {}

for uf in uploaded_files:
    df = pd.read_csv(uf)
    df.columns = [c.strip() for c in df.columns]
    df["__source_file"] = uf.name
    file_dfs[uf.name] = df

    base = derive_model_name(uf.name)
    if base in dup_counter:
        dup_counter[base] += 1
        base = f"{base}-{dup_counter[base]}"
    else:
        dup_counter[base] = 1
    model_names[uf.name] = base

    uf.seek(0)  # Excel 用にポインタ巻き戻し

all_data = pd.concat(file_dfs.values(), ignore_index=True)
available_tags = [c for c in TAG_COLUMNS if c in all_data.columns]

# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------
def build_tag_summary(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            tag: df[tag].value_counts().reindex(SENTIMENTS, fill_value=0)
            for tag in available_tags
        }
    ).T


def calc_ratio_df(files: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = {}
    for fname, df in files.items():
        ratios = {}
        for tag in available_tags:
            cnt = df[tag].value_counts().reindex(SENTIMENTS, fill_value=0)
            total = int(cnt.sum()) or 1
            for s in SENTIMENTS:
                ratios[f"{tag}_{s}"] = round(cnt[s] / total * 100, 2)
        ratios["TotalReviews"] = len(df)
        rows[model_names[fname]] = ratios
    return pd.DataFrame.from_dict(rows, orient="index")


ratio_df = calc_ratio_df(file_dfs)

# ---------------------------------------------------------------------------
# 個別解析表示 (タブ)
# ---------------------------------------------------------------------------
def render_single(df: pd.DataFrame):
    tbl = build_tag_summary(df)
    st.subheader("タグ別 ポジ/ネガ/ニュートラル 件数")
    with st.expander("タグ別 ポジ/ネガ/ニュートラル 件数（表）", expanded=False):
        st.dataframe(tbl)

    st.subheader("タグ別ヒストグラム")
    fig, ax = plt.subplots()
    tbl.plot(kind="bar", stacked=True, ax=ax)
    ax.set_xlabel("Tag")
    ax.set_ylabel("Count")
    ax.legend(title="Sentiment")
    st.pyplot(fig)


tabs = st.tabs(
    (["All Files"] if len(file_dfs) > 1 else [])
    + [model_names[f] for f in file_dfs.keys()]
)

if len(file_dfs) > 1:
    with tabs[0]:
        st.header("【All Files (Aggregate)】")
        render_single(all_data)

start_idx = 0 if len(file_dfs) == 1 else 1
for tab, (fname, df) in zip(tabs[start_idx:], file_dfs.items()):
    with tab:
        st.header(f"【{model_names[fname]}】の解析結果")
        render_single(df)

# ---------------------------------------------------------------------------
# モデル比較 (感情割合 + 件数)
# ---------------------------------------------------------------------------
st.header("モデル比較: タグ別スコア割合 (Pos/Neu/Neg) + 件数")
chosen_tag = st.selectbox("比較したいタグ", available_tags)

view = ratio_df[[f"{chosen_tag}_{s}" for s in SENTIMENTS] + ["TotalReviews"]].copy()
view.columns = SENTIMENTS + ["Reviews"]
st.dataframe(view)

fig_cmp, ax_cmp = plt.subplots()
view[SENTIMENTS].plot(kind="bar", stacked=True, ax=ax_cmp)
ax_cmp.set_ylabel("Percentage (%)")
ax_cmp.set_title(f"{chosen_tag}")
ax_cmp.legend(title="Sentiment", bbox_to_anchor=(1.05, 1), loc="upper left")

# ------------------------------------------------------------------
# ★ 件数ラベルを “棒の最上端” に描画
# ------------------------------------------------------------------
# Positive 部分の Rect で x 位置を取得
pos_rects = ax_cmp.containers[0]
for rect, total in zip(pos_rects, view["Reviews"]):
    x_center = rect.get_x() + rect.get_width() / 2
    ax_cmp.text(
        x_center,
        100 + 1,                  # 100% の少し上に固定配置
        f"{int(total)}",
        ha="center",
        va="bottom",
        fontsize=8,
    )
# 余白確保
ax_cmp.set_ylim(0, 105)

st.pyplot(fig_cmp)

# ---------------------------------------------------------------------------
# キーワード検索
# ---------------------------------------------------------------------------
st.header("キーワード検索 / フィルタ (All Files)")
kw = st.text_input("含めたいキーワード (複数語はスペース区切り)")
if kw:
    mask = (
        all_data["review_text"].astype(str).str.contains(kw, case=False, na=False)
        | all_data.get("translated_text", pd.Series("", index=all_data.index))
        .astype(str)
        .str.contains(kw, case=False, na=False)
    )
    filtered = all_data[mask]
    st.write(f"該当レビュー数: {len(filtered)} 件")
    st.dataframe(filtered)
else:
    filtered = all_data

# ---------------------------------------------------------------------------
# Excel 出力
# ---------------------------------------------------------------------------
st.header("Excel エクスポート")

tag_summary_all = build_tag_summary(all_data)
sentiment_all = (
    all_data["overall_sentiment_score"].dropna().astype(float)
    if "overall_sentiment_score" in all_data.columns
    else pd.Series(dtype=float)
)
rating_all = (
    all_data["rating"].astype(int).value_counts().sort_index()
    if "rating" in all_data.columns
    else pd.Series(dtype=int)
)

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    tag_summary_all.to_excel(writer, sheet_name="Tag_Summary_All")
    ratio_df.to_excel(writer, sheet_name="Tag_Ratios")
    if not sentiment_all.empty:
        sentiment_all.to_frame("overall_sentiment_score").to_excel(
            writer, sheet_name="Sentiment_Score_All"
        )
    if not rating_all.empty:
        rating_all.to_frame("count").to_excel(
            writer, sheet_name="Rating_Distribution_All"
        )
    filtered.to_excel(writer, sheet_name="Filtered_Reviews", index=False)

    for fname, df in file_dfs.items():
        sheet_name = f"{model_names[fname][:28]}_集計"
        build_tag_summary(df).to_excel(writer, sheet_name=sheet_name)

# グラフ貼り付け
buf.seek(0)
wb = load_workbook(buf)
for fname, df in file_dfs.items():
    sheet = f"{model_names[fname][:28]}_集計"
    if sheet not in wb.sheetnames:
        continue
    ws = wb[sheet]

    fig, ax = plt.subplots()
    build_tag_summary(df).plot(kind="bar", stacked=True, ax=ax)
    ax.set_xlabel("Tag")
    ax.set_ylabel("Count")
    ax.legend(title="Sentiment")

    img_data = io.BytesIO()
    fig.savefig(img_data, format="png", bbox_inches="tight")
    plt.close(fig)
    img_data.seek(0)

    img = XLImage(img_data)
    img.width = 480
    img.height = 320
    ws.add_image(img, "H2")

# ダウンロード
out = io.BytesIO()
wb.save(out)
out.seek(0)

st.download_button(
    "解析結果を Excel でダウンロード",
    data=out,
    file_name=f"bestbuy_review_analysis_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("集計完了!")
