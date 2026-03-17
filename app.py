"""
PILOT Webカタログ JANコード一括データ取得ツール
"""

import math
import re
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────
# ページ設定（最初に呼ぶ必要あり）
# ─────────────────────────────────────────────────
st.set_page_config(
    page_title="商品情報取得ツール",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────
# カスタムCSS
# ─────────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', 'Noto Sans JP', sans-serif;
        background-color: #FFFFFF;
        color: #333333;
    }

    /* 全体背景 */
    .stApp {
        background-color: #FFFFFF;
        min-height: 100vh;
    }

    /* ヘッダーバナー (モダンでフラットな赤白デザイン) */
    .hero-banner {
        background: #D00000;
        border-radius: 16px;
        padding: 40px 48px;
        margin-bottom: 32px;
        box-shadow: 0 10px 30px rgba(208, 0, 0, 0.15);
        position: relative;
        overflow: hidden;
    }
    .hero-banner::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 350px;
        height: 350px;
        background: rgba(255,255,255,0.1);
        border-radius: 50%;
    }
    .hero-title {
        font-size: 2.4rem;
        font-weight: 700;
        color: #ffffff;
        margin: 0 0 8px 0;
        letter-spacing: -0.5px;
        position: relative;
        z-index: 1;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        color: rgba(255,255,255,0.9);
        margin: 0;
        position: relative;
        z-index: 1;
    }

    /* ステップカード */
    .steps-container {
        display: flex;
        gap: 16px;
        margin-bottom: 32px;
    }
    .step-card {
        flex: 1;
        background: #ffffff;
        border: 1px solid #f0f0f0;
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        transition: all 0.3s ease;
    }
    .step-card:hover {
        border-color: #D00000;
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(208, 0, 0, 0.1);
    }
    .step-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 32px;
        height: 32px;
        background: #D00000;
        border-radius: 50%;
        font-size: 0.85rem;
        font-weight: 700;
        color: white;
        margin-bottom: 10px;
    }
    .step-title {
        font-size: 0.95rem;
        font-weight: 600;
        color: #333333;
        margin: 0 0 6px 0;
    }
    .step-desc {
        font-size: 0.82rem;
        color: #666666;
        margin: 0;
        line-height: 1.5;
    }

    /* 入力セクション */
    .input-section {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 28px 32px;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.02);
    }
    .section-label {
        font-size: 0.95rem;
        font-weight: 700;
        color: #D00000;
        letter-spacing: 0.05em;
        margin-bottom: 12px;
        border-bottom: 2px solid #ffeeee;
        padding-bottom: 6px;
        display: inline-block;
    }

    /* テキストエリア・セレクトボックスの視認性改善（強制白背景・黒文字） */
    .stTextArea textarea, .stSelectbox div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 1px solid #d1d5db !important;
        border-radius: 8px !important;
        color: #000000 !important;
        font-size: 0.95rem !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
    }
    .stTextArea textarea {
        font-family: 'Courier New', monospace !important;
    }
    .stTextArea textarea:focus, .stSelectbox div[data-baseweb="select"] > div:focus-within {
        border-color: #D00000 !important;
        box-shadow: 0 0 0 3px rgba(208, 0, 0, 0.15) !important;
        outline: none !important;
    }
    /* プレースホルダーの文字色調整 */
    .stTextArea textarea::placeholder {
        color: #9ca3af !important;
    }
    /* セレクトボックスのテキスト色対応 */
    .stSelectbox div[data-baseweb="select"] span {
        color: #000000 !important;
        font-weight: 500 !important;
    }
    .stSelectbox div[role="listbox"] {
        background-color: #ffffff !important;
    }
    .stSelectbox div[role="option"] {
        color: #000000 !important;
    }
    .stSelectbox div[role="option"]:hover, .stSelectbox div[role="option"][aria-selected="true"] {
        background-color: #ffeeee !important;
        color: #D00000 !important;
    }

    /* ボタン */
    .stButton > button {
        background: #D00000 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 32px !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.02em !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 12px rgba(208, 0, 0, 0.2) !important;
        width: 100% !important;
    }
    .stButton > button:hover {
        background: #a30000 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(208, 0, 0, 0.3) !important;
    }
    .stButton > button:active {
        transform: translateY(0) !important;
    }

    /* プログレス */
    .stProgress > div > div {
        background-color: #D00000 !important;
        border-radius: 8px !important;
    }
    .stProgress > div {
        background-color: #f3f4f6 !important;
        border-radius: 8px !important;
    }

    /* ステータステキスト */
    .status-box {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-left: 4px solid #D00000;
        border-radius: 8px;
        padding: 12px 18px;
        color: #374151;
        font-size: 0.9rem;
        margin-top: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    }
    .status-highlight {
        color: #D00000;
        font-weight: 600;
    }

    /* 結果セクション */
    .result-section {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 28px 32px;
        margin-top: 24px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.02);
    }

    /* データフレーム */
    .stDataFrame {
        border-radius: 8px !important;
        border: 1px solid #e5e7eb !important;
    }
    /* データフレーム内のヘッダー等のテキスト色対応 */
    [data-testid="stDataFrame"] div, [data-testid="stDataFrame"] span {
        color: #333333 !important;
    }

    /* ダウンロードボタン */
    .stDownloadButton > button {
        background: #ffffff !important;
        color: #D00000 !important;
        border: 2px solid #D00000 !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        font-size: 0.95rem !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        background: #fff5f5 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(208,0,0,0.1) !important;
    }

    /* 警告・エラー */
    .stWarning {
        background: #fffbeb !important;
        border: 1px solid #fcd34d !important;
        border-radius: 8px !important;
        color: #b45309 !important;
    }

    /* メトリクスカード */
    .metric-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 16px 20px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #333333;
        line-height: 1;
    }
    .metric-value.success { color: #10b981; }
    .metric-value.error { color: #ef4444; }
    
    .metric-label {
        font-size: 0.8rem;
        color: #6b7280;
        margin-top: 6px;
        font-weight: 500;
    }

    /* フッター */
    .footer {
        text-align: center;
        color: #9ca3af;
        font-size: 0.8rem;
        margin-top: 48px;
        padding-bottom: 24px;
        border-top: 1px solid #f3f4f6;
        padding-top: 24px;
    }

    /* Streamlit デフォルト要素の調整 */
    .stMarkdown p { color: #4b5563; }
    h1, h2, h3 { color: #111827 !important; }
    label { color: #4b5563 !important; font-weight: 500 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────
# 解析・計算用ユーティリティ関数
# ─────────────────────────────────────────────────

def extract_weight(weight_str: str) -> float | None:
    """重量の文字列から数値を抽出する"""
    if weight_str == "取得不可" or not weight_str:
        return None
    # \d+(?:\.\d+)? で整数または小数を抽出
    match = re.search(r"(\d+(?:\.\d+)?)", weight_str)
    if match:
        return float(match.group(1))
    return None

def calculate_volume(size_str: str) -> float | str:
    """サイズの文字列を解析し、体積(cm³)を計算する（小数点以下1位で四捨五入）"""
    if size_str == "取得不可" or size_str == "-" or not size_str:
        return "計算不可"

    # パターンA（円柱 - PILOT用）: "最大径φ12.8mm 全長144.9mm"
    match_a = re.search(r"径[^\d]*(\d+(?:\.\d+)?).*?長[^\d]*(\d+(?:\.\d+)?)", size_str)
    if match_a:
        diameter = float(match_a.group(1))
        length = float(match_a.group(2))
        radius_cm = (diameter / 2) / 10
        length_cm = length / 10
        vol = (radius_cm ** 2) * math.pi * length_cm
        return round(vol, 1)

    # パターンC（1999.co.jp等の cm 記述）を先に判定する（mmと誤認されるのを防ぐため）
    match_c = re.search(r"(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*(?:cm|㎝)", size_str, re.IGNORECASE)
    if match_c:
        l = float(match_c.group(1))
        w = float(match_c.group(2))
        h = float(match_c.group(3))
        vol = l * w * h
        return round(vol, 1)

    # パターンB（直方体 - PILOT/タカラトミー等の mm 記述）: "185×50×15mm" または "W225×H165×D97" または "82㎜×86㎜×36㎜"
    match_b = re.search(r"(?:W\s*)?(\d+(?:\.\d+)?)(?:mm|㎜)?\s*[×x*]\s*(?:H\s*)?(\d+(?:\.\d+)?)(?:mm|㎜)?\s*[×x*]\s*(?:D\s*)?(\d+(?:\.\d+)?)(?:mm|㎜)?", size_str, re.IGNORECASE)
    if match_b:
        l = float(match_b.group(1)) / 10
        w = float(match_b.group(2)) / 10
        h = float(match_b.group(3)) / 10
        vol = l * w * h
        return round(vol, 1)

    return "計算不可"


# ─────────────────────────────────────────────────
# バックエンド: スクレイピング関数
# ─────────────────────────────────────────────────

BASE_URL = "https://webcatalog.pilot.co.jp"
CATALOG_VOLUME = "00004"
SEARCH_URL = f"{BASE_URL}/products/DispSearch.do?volumeName={CATALOG_VOLUME}"
DETAIL_URL = f"{BASE_URL}/products/DispDetail.do"

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.7,en;q=0.3",
}

UNAVAILABLE = "取得不可"


def get_item_id_by_jan(jan_code: str, session: requests.Session) -> str | None:
    """PILOT: JANコードをPOSTし、フォームのaction属性からitemIDを取得する"""
    try:
        payload = {
            "searchTypeParam": "janSearch",
            "searchValue": jan_code.strip(),
        }
        resp = session.post(
            SEARCH_URL,
            data=payload,
            headers=DEFAULT_HEADERS,
            allow_redirects=True,
            timeout=20,
        )
        resp.raise_for_status()

        # フォームのhidden inputからitemIDを正規表現で抽出
        # html内に <input type="hidden" name="itemID" value="t000100002290"> がある
        match = re.search(r'name="itemID"\s+value="([^"]+)"', resp.text)
        if match:
            return match.group(1)
        
        # 従来のURLパラメータ型（念のため残す）
        match_url = re.search(r"itemID=([A-Za-z0-9]+)", resp.text)
        if match_url:
            return match_url.group(1)
        return None
    except Exception:
        return None


def get_pilot_product_detail(item_id: str, session: requests.Session) -> dict:
    """PILOT: 商品詳細ページから商品名・サイズ・重量などを取得する"""
    result = {
        "商品名": UNAVAILABLE,
        "サイズ(元データ)": UNAVAILABLE,
        "重量(元データ)": UNAVAILABLE,
        "情報元URL": UNAVAILABLE,
    }
    try:
        params = {"volumeName": CATALOG_VOLUME, "itemID": item_id}
        resp = session.get(
            DETAIL_URL,
            params=params,
            headers=DEFAULT_HEADERS,
            timeout=20,
        )
        resp.raise_for_status()

        page_url = resp.url
        result["情報元URL"] = page_url

        soup = BeautifulSoup(resp.text, "lxml")

        # 商品名 (class="p-detail-title" の h1 タグ)
        title_elem = soup.find("h1", class_="p-detail-title")
        if title_elem:
            result["商品名"] = title_elem.get_text(strip=True)
        else:
            # フォールバック: titleタグからパース（「|」区切りの最初の要素）
            page_title = soup.find("title")
            if page_title:
                parts = page_title.get_text().split("｜")
                if len(parts) >= 2:
                    result["商品名"] = parts[1].strip()

        # スペックテーブルからサイズ・重量を抽出
        tables = soup.find_all("table")
        for table in tables:
            rows = table.find_all("tr")
            for row in rows:
                th = row.find("th")
                td = row.find("td")
                if not (th and td):
                    continue
                key = th.get_text(strip=True)
                value = td.get_text(separator=" ", strip=True)
                value = re.sub(r"\s+", " ", value).strip()
                if "サイズ" in key:
                    result["サイズ(元データ)"] = value
                elif "重量" in key:
                    result["重量(元データ)"] = value

        return result
    except Exception:
        return result


def fetch_takaratomy_product(jan_code: str, _session: requests.Session) -> dict:
    """タカラトミー: JANコードから商品情報を取得する (curl_cffi でブロック回避)"""
    from curl_cffi import requests as cffi_requests
    jan_str = jan_code.strip()
    url = f"https://takaratomymall.jp/shop/g/g{jan_str}/"
    result = {
        "メーカー": "タカラトミー",
        "商品名": UNAVAILABLE,
        "サイズ(元データ)": UNAVAILABLE,
        "重量(元データ)": "-",  # 重量取得不要
        "情報元URL": url,
    }
    try:
        # User-Agent偽装ではなくTLSフィンガープリント全体をChromeに偽装してCloudflare回避
        resp = cffi_requests.get(url, impersonate="chrome110", timeout=15)
        # 404などのエラーは例外にする
        if resp.status_code != 200:
            return result

        soup = BeautifulSoup(resp.text, "lxml")

        # 商品名
        title_elem = soup.find("h1", class_="tt_block17__titleMainLabel")
        if title_elem:
            result["商品名"] = title_elem.get_text(strip=True)

        # パッケージサイズ
        spec_items = soup.find_all("li", class_="tt_block17__specItem")
        for item in spec_items:
            label_elem = item.find("span", class_="tt_block17__specLabel")
            if label_elem and "パッケージサイズ" in label_elem.get_text():
                val_elem = item.find("span", class_="tt_block17__specValue")
                if val_elem:
                    result["サイズ(元データ)"] = val_elem.get_text(strip=True)
                    break
        
        # もし `<li>` で見つからない場合は `<th>` パターンも捜索 (フォールバック)
        if result["サイズ(元データ)"] == UNAVAILABLE:
            for th in soup.find_all("th"):
                if "パッケージサイズ" in th.get_text():
                    td = th.find_next_sibling("td")
                    if td:
                        result["サイズ(元データ)"] = td.get_text(strip=True)
                        break

        return result
    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return result


def fetch_1999_product(jan_code: str, _session: requests.Session) -> dict:
    """その他おもちゃ: 1999.co.jp からJANコードで検索し商品情報を取得する"""
    import urllib.request
    from urllib.error import HTTPError, URLError

    jan_str = jan_code.strip()
    search_url = (
        f"https://www.1999.co.jp/search?typ1_c=101&cat=&target=JanCode&searchkey={jan_str}"
    )
    result = {
        "メーカー": "その他おもちゃ",
        "商品名": UNAVAILABLE,
        "サイズ(元データ)": UNAVAILABLE,
        "重量(元データ)": UNAVAILABLE,
        "情報元URL": search_url,
    }
    try:
        req = urllib.request.Request(
            search_url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ja",
            }
        )
        with urllib.request.urlopen(req, timeout=15) as response:
            html = response.read().decode("utf-8", errors="ignore")
            result["情報元URL"] = response.geturl()  # リダイレクト後の実際のURL
            
        soup = BeautifulSoup(html, "lxml")

        # 商品名: 正しいクラス名は c-product-detail__info-title
        title_elem = soup.find("h1", class_="c-product-detail__info-title")
        if not title_elem:
            title_elem = soup.find("title")
        if title_elem:
            raw_title = title_elem.get_text(strip=True)
            result["商品名"] = re.split(r"[\|\-]", raw_title)[0].strip()

        # ページ全文から「●パッケージサイズ/重さ」等を正規表現で抽出
        page_text = soup.get_text(separator="\n").replace("\r\n", "\n").replace("\r", "\n")
        
        # 1. 両方記載パターン 「●パッケージサイズ/重さ : 12.2 x 8.4 x 1 cm / 28g」
        size_weight_match = re.search(
            r"パッケージサイズ[/／]重[さ量][^:：]*[:：]\s*([\d.]+\s*[xX×]\s*[\d.]+\s*[xX×]\s*[\d.]+\s*cm)\s*[/／]\s*([\d.]+\s*[gGkK]+)",
            page_text,
            re.IGNORECASE,
        )
        if size_weight_match:
            result["サイズ(元データ)"] = size_weight_match.group(1).strip()
            result["重量(元データ)"] = size_weight_match.group(2).strip()
        else:
            # 2. サイズのみ記載パターン
            size_only = re.search(
                r"パッケージサイズ[^:：]*[:：]\s*([\d.]+\s*[xX×]\s*[\d.]+\s*[xX×]\s*[\d.]+\s*cm)",
                page_text, re.IGNORECASE
            )
            if size_only:
                result["サイズ(元データ)"] = size_only.group(1).strip()

            # 3. 重量のみ記載パターン 「●重さ : 50g」など
            weight_only = re.search(
                r"重[さ量][^:：]*[:：]\s*([\d.]+\s*[gGkK]+)",
                page_text, re.IGNORECASE
            )
            if weight_only:
                result["重量(元データ)"] = weight_only.group(1).strip()

        return result
    except HTTPError as e:
        print(f"HTTPError fetching 1999.co.jp for {jan_code}: {e.code}")
        # UIにエラー原因を表示させるため商品名に代入
        result["商品名"] = f"エラー: {e.code} (アクセス拒否)"
        return result
    except URLError as e:
        print(f"URLError fetching 1999.co.jp for {jan_code}: {e.reason}")
        result["商品名"] = f"エラー: 接続失敗"
        return result
    except Exception as e:
        print(f"Error fetching 1999.co.jp for {jan_code}: {e}")
        return result


def fetch_product_data(jan_code: str, maker: str, session: requests.Session) -> dict:
    """JANコードとメーカーから商品情報を取得し、共通の解析フォーマットで返す"""
    record = {
        "JANコード": jan_code.strip(),
        "メーカー": maker,
        "商品名": UNAVAILABLE,
        "サイズ(元データ)": UNAVAILABLE,
        "重量(元データ)": UNAVAILABLE,
        "体積(cm³)": "計算不可",
        "重量(数値)": None,
        "情報元URL": UNAVAILABLE,
        "ステータス": "❌ 取得失敗",
    }

    if maker == "PILOT":
        item_id = get_item_id_by_jan(jan_code, session)
        if not item_id:
            record["ステータス"] = "❌ 商品が見つかりません"
            return record
        detail = get_pilot_product_detail(item_id, session)
        record.update(detail)
    elif maker == "タカラトミー":
        detail = fetch_takaratomy_product(jan_code, session)
        # 取得失敗時は商品名がUNAVAILABLEのまま
        if detail["商品名"] == UNAVAILABLE:
            record["ステータス"] = "❌ 商品が見つかりません"
            record.update(detail)
            return record
        record.update(detail)
    elif maker == "その他おもちゃ":
        detail = fetch_1999_product(jan_code, session)
        if detail["商品名"] == UNAVAILABLE:
            record["ステータス"] = "❌ 商品が見つかりません"
            record.update(detail)
            return record
        record.update(detail)

    # 共通処理: 重量の数値化と体積の計算
    record["重量(数値)"] = extract_weight(record.get("重量(元データ)"))
    record["体積(cm³)"] = calculate_volume(record.get("サイズ(元データ)"))

    if record["商品名"] != UNAVAILABLE:
        record["ステータス"] = "✅ 取得成功"
    else:
        record["ステータス"] = "⚠️ 一部取得不可"

    return record


# ─────────────────────────────────────────────────
# UI: ヘッダー
# ─────────────────────────────────────────────────

st.markdown(
    """
    <div class="hero-banner">
        <div class="hero-title">📦 商品情報取得ツール</div>
        <div class="hero-subtitle">
            JANコードとメーカーを選択するだけで、Web上のカタログサイトから商品のサイズや重量などの情報を自動収集し、整理された形式で表示・保存できます。
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────
# UI: 使い方ステップ
# ─────────────────────────────────────────────────

st.markdown(
    """
    <div class="steps-container">
        <div class="step-card">
            <div class="step-number">1</div>
            <div class="step-title">📋 JANコードを入力</div>
            <div class="step-desc">取得したい商品のJANコード（13桁）を改行区切りで貼り付けます</div>
        </div>
        <div class="step-card">
            <div class="step-number">2</div>
            <div class="step-title">🚀 取得開始ボタンを押す</div>
            <div class="step-desc">選択したメーカーのWebサイトへ自動でアクセスし、商品情報を順次取得します</div>
        </div>
        <div class="step-card">
            <div class="step-number">3</div>
            <div class="step-title">📊 結果を確認・ダウンロード</div>
            <div class="step-desc">サイズ・重量データの数値化と計算結果も自動で表示されます</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────
# UI: 入力フォーム
# ─────────────────────────────────────────────────

col_input, col_options = st.columns([3, 1])

with col_input:
    st.markdown('<div class="section-label">📥 取得対象の指定</div>', unsafe_allow_html=True)
    
    selected_maker = st.selectbox(
        "メーカー",
        ["PILOT", "タカラトミー", "その他おもちゃ"],
        help="情報を取得するメーカー（サイト）を選択してください"
    )
    
    st.markdown('<div style="font-size: 0.9rem; font-weight: 500; color: #4b5563; margin-top: 12px; margin-bottom: 4px;">JANコード (改行区切り)</div>', unsafe_allow_html=True)
    jan_input = st.text_area(
        label="jan_input_area",
        label_visibility="collapsed",
        placeholder=(
            "ここに取得したいJANコードを改行区切りで貼り付けてください\n\n"
            "例:\n"
            "4902505607110\n"
            "4902505516177\n"
            "4902505398896"
        ),
        height=200,
        key="jan_input",
    )

with col_options:
    st.markdown('<div class="section-label">⚙️ 設定</div>', unsafe_allow_html=True)
    delay_sec = st.slider(
        "リクエスト間隔（秒）",
        min_value=0.5,
        max_value=3.0,
        value=1.0,
        step=0.5,
        help="サーバーへの負荷軽減のため、各リクエストの間に待機時間を設けます",
    )
    st.markdown(
        "<div style='color: rgba(255,255,255,0.4); font-size:0.75rem; margin-top:8px;'>"
        "※ アクセス間隔を短くすると接続先サーバーに負荷がかかるためご注意ください"
        "</div>",
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)
run_button = st.button("🔍　取得開始", key="run_btn", use_container_width=True)

# ─────────────────────────────────────────────────
# UI: 実行処理
# ─────────────────────────────────────────────────

if run_button:
    # バリデーション
    raw_lines = [line.strip() for line in jan_input.splitlines() if line.strip()]
    jan_list = [l for l in raw_lines if l]

    if not jan_list:
        st.warning("⚠️ JANコードが入力されていません。テキストエリアに取得対象のJANコードを改行区切りで入力してから実行してください。")
    else:
        total = len(jan_list)
        results = []

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f'<div class="section-label">⏳ データの取得・解析状況 '
            f'<span style="color:#6b7280; font-size:0.85rem; margin-left: 8px;">(対象: {total} 件)</span></div>',
            unsafe_allow_html=True,
        )

        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-label">📊 データ取得結果（リアルタイム更新）</div>', unsafe_allow_html=True)
        
        metric_placeholder = st.empty()
        table_placeholder = st.empty()

        display_cols = [
            "JANコード",
            "メーカー",
            "商品名",
            "サイズ(元データ)",
            "重量(元データ)",
            "体積(cm³)",
            "重量(数値)",
            "情報元URL"
        ]

        with requests.Session() as session:
            for idx, jan in enumerate(jan_list):
                # ステータス表示更新
                status_placeholder.markdown(
                    f'<div class="status-box">'
                    f'🔄 現在 <span class="status-highlight">{idx + 1} 件目</span> / {total} 件'
                    f' &nbsp;|&nbsp; JAN: <span class="status-highlight">{jan}</span> のデータを取得中...'
                    f"</div>",
                    unsafe_allow_html=True,
                )

                # データ取得
                record = fetch_product_data(jan, selected_maker, session)
                results.append(record)

                # ---------------------------------------------
                # リアルタイム UI 更新 (メトリクス & テーブル)
                # ---------------------------------------------
                success_count = sum(1 for r in results if "✅" in r.get("ステータス", ""))
                fail_count = len(results) - success_count

                with metric_placeholder.container():
                    mc1, mc2, mc3 = st.columns(3)
                    with mc1:
                        st.markdown(
                            f'<div class="metric-card">'
                            f'<div class="metric-value">{total}</div>'
                            f'<div class="metric-label">合計設定件数</div>'
                            f"</div>",
                            unsafe_allow_html=True,
                        )
                    with mc2:
                        st.markdown(
                            f'<div class="metric-card">'
                            f'<div class="metric-value success">{success_count}</div>'
                            f'<div class="metric-label">✅ 取得成功</div>'
                            f"</div>",
                            unsafe_allow_html=True,
                        )
                    with mc3:
                        st.markdown(
                            f'<div class="metric-card">'
                            f'<div class="metric-value error">{fail_count}</div>'
                            f'<div class="metric-label">❌ 取得失敗</div>'
                            f"</div>",
                            unsafe_allow_html=True,
                        )

                current_df = pd.DataFrame(results)
                current_df = current_df[[c for c in display_cols if c in current_df.columns]]

                table_placeholder.dataframe(
                    current_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "JANコード": st.column_config.TextColumn("JANコード", width="medium"),
                        "メーカー": st.column_config.TextColumn("メーカー", width="small"),
                        "商品名": st.column_config.TextColumn("商品名", width="large"),
                        "サイズ(元データ)": st.column_config.TextColumn("サイズ(元データ)", width="large"),
                        "重量(元データ)": st.column_config.TextColumn("重量(元データ)", width="small"),
                        "体積(cm³)": st.column_config.TextColumn("体積(cm³)", width="medium"),
                        "重量(数値)": st.column_config.NumberColumn("重量(数値)", format="%.1f", width="small"),
                        "情報元URL": st.column_config.LinkColumn(
                            "情報元URL",
                            display_text="カタログページを開く",
                            width="medium",
                        ),
                    },
                    height=min(400, 80 + 35 * len(current_df)),
                )

                # プログレスバー更新
                progress_bar.progress((idx + 1) / total)

                # サーバー負荷軽減
                if idx < total - 1:
                    time.sleep(delay_sec)

        # 完了通知
        status_placeholder.markdown(
            '<div class="status-box" style="border-color: #10b981; '
            'background: #ecfdf5; color: #047857;">'
            "✅ すべてのJANコードの取得処理が完了しました"
            "</div>",
            unsafe_allow_html=True,
        )

        st.toast(
            f"🎉 取得完了！ 成功: {success_count}件 / 失敗: {fail_count}件",
            icon="✅",
        )

        df = pd.DataFrame(results)
        df = df[[c for c in display_cols if c in df.columns]]

        # ─────────────────────────────────────────────────
        # UI: Excel転記用コピー
        # ─────────────────────────────────────────────────

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-label">📋 Excel転記用データ</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="color: #6b7280; font-size: 0.9rem; margin-bottom: 8px;">'
            "以下のテキスト枠の右上にある「コピー」アイコンをクリックしてコピーし、そのままExcelに貼り付けてください。"
            "</div>",
            unsafe_allow_html=True,
        )

        # 抽出する列: JANコード, メーカー, 体積(cm³), 重量(数値)
        copy_df = df[["JANコード", "メーカー", "体積(cm³)", "重量(数値)"]].copy()
        copy_df.fillna("", inplace=True) # NoneやNaNを空文字に変換
        
        # ヘッダーを含めてTSV形式で文字列化
        tsv_str = copy_df.to_csv(sep="\t", index=False)
        
        # st.code で表示（右上に標準のコピーボタンが表示される）
        st.code(tsv_str, language="text")

        # ─────────────────────────────────────────────────
        # UI: CSVダウンロード
        # ─────────────────────────────────────────────────

        now_str = datetime.now().strftime("%Y%m%d_%H%M")
        csv_filename = f"pilot_data_{now_str}.csv"

        # CSV（BOM付きUTF-8 でExcelでも文字化けしない）
        csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️　結果をCSVとしてダウンロードする　（{csv_filename}）",
            data=csv_bytes,
            file_name=csv_filename,
            mime="text/csv",
            use_container_width=True,
        )

        st.markdown(
            f'<div style="color:rgba(255,255,255,0.3); font-size:0.75rem; text-align:center; margin-top:8px;">'
            f"出力ファイル名: {csv_filename}　｜　文字コード: UTF-8 BOM（Excel対応）"
            f"</div>",
            unsafe_allow_html=True,
        )

# ─────────────────────────────────────────────────
# フッター
# ─────────────────────────────────────────────────

st.markdown(
    '<div class="footer">'
    "PILOT Webカタログ データ取得ツール　｜　社内利用専用　｜　"
    "データソース: webcatalog.pilot.co.jp"
    "</div>",
    unsafe_allow_html=True,
)
