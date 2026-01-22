# -*- coding: utf-8 -*-
import re
import time
from urllib.parse import urlencode, urlparse

import pandas as pd
from playwright.sync_api import sync_playwright


# ====== 設定 ======
CSV_NAME = "EP案件一覧 - EP (2).csv"
OUT_NAME = "EP案件一覧 - EP (2)_ranked.csv"

# 列位置（0-based）
COL_F_URL = 5     # F
COL_G_FLAG = 6    # G
COL_J_KEYWORD = 9 # J
COL_AZ = 51       # AZ
COL_BA_RESULT = 52  # BA（AZの次）

TODAY_MMDD = "1/9"
MAX_ORGANIC_RANK = 50

MOBILE_USER_AGENT = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1"
)
VIEWPORT = {"width": 390, "height": 844}


# ====== df列数確保 ======
def ensure_col_count(df: pd.DataFrame, n_cols: int) -> pd.DataFrame:
    if df.shape[1] < n_cols:
        for i in range(df.shape[1], n_cols):
            df[i] = ""
    return df


# ====== URL正規化（末尾スラッシュ除去込み） ======
def normalize_url_for_match(u: str) -> str:
    """
    比較用にURLを正規化：
    - スキーム除去
    - www. は除去（wwwあり/なしで一致しない事故を防ぐ）
    - 末尾スラッシュ除去（追加条件）
    - クエリ/フラグメント除去
    - path は末尾 / を削る
    戻り値例: example.com/path/to/page
    """
    if not u:
        return ""
    u = u.strip()
    if not u:
        return ""

    # 末尾スラッシュがあるなら除外して一致検索（=削る）
    # （追加条件）
    if u.endswith("/"):
        u = u[:-1]

    # スキームがなければ仮で付けて parse する
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u

    p = urlparse(u)
    netloc = (p.netloc or "").lower()
    if netloc.startswith("www."):
        netloc = netloc[4:]

    path = (p.path or "").rstrip("/")
    return f"{netloc}{path}"


def extract_result_href(href: str) -> str:
    """
    Google結果の href を実URLへ寄せて、比較用正規化した文字列を返す。
    - /url?q=... も https://www.google.com/url?q=... も対応
    - 相対URLやgoogle内部URLは除外（ただし url?q= を展開した後）
    """
    if not href:
        return ""

    href = href.strip()

    # 1) https://www.google.com/url?q=... を先に展開
    # 2) /url?q=... も展開
    if ("/url?" in href) and ("q=" in href):
        m = re.search(r"[?&]q=([^&]+)", href)
        if m:
            href = m.group(1)

    # 相対URLは捨てる
    if href.startswith("/"):
        return ""

    # google内部URLは捨てる（※url?q展開後に判定するのが重要）
    low = href.lower()
    if low.startswith("https://www.google.") or low.startswith("https://google."):
        return ""
    if "googleadservices.com" in low:
        return ""

    return normalize_url_for_match(href)


def looks_like_ad(link_locator) -> bool:
    """
    広告を除外したいので、リンク周辺のラベルを軽くチェック
    （Googleは構造が変わるので“強すぎない”判定にする）
    """
    try:
        container = link_locator.locator("xpath=ancestor::*[self::div or self::article][1]")
        txt = container.inner_text(timeout=300).strip()
    except Exception:
        txt = ""

    # 広告ラベル（日本語/英語）
    if "広告" in txt or "スポンサー" in txt or "Sponsored" in txt:
        return True
    return False


def get_rank(page, keyword: str, target_url: str) -> str:
    target_norm = normalize_url_for_match(target_url)
    if not keyword.strip() or not target_norm:
        return ""

    search_url = "https://www.google.com/search?" + urlencode({
        "hl": "ja",
        "gl": "jp",
        "q": keyword.strip()
    })

    page.goto(search_url, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1200)

    organic_rank = 0
    seen = set()

    for _ in range(30):
        # ★ここが変更点：h3ではなく a.rTyHce を使う
        links = page.locator("a.rTyHce[href]")
        cnt = links.count()

        for i in range(cnt):
            a = links.nth(i)

            href = (a.get_attribute("href") or "").strip()
            if not href:
                continue

            # （任意）広告判定：強すぎると自然検索も弾くので基本は弱め
            # まずは確実に順位が取れるか優先で、広告判定は軽くする
            # if looks_like_ad(a):
            #     continue

            norm = extract_result_href(href)
            if not norm:
                continue

            # 重複排除
            if norm in seen:
                continue
            seen.add(norm)

            organic_rank += 1

            # 一致判定（末尾スラッシュ除去済み同士で比較）
            if norm == target_norm:
                return str(organic_rank)

            if organic_rank >= MAX_ORGANIC_RANK:
                return "50~"

        # 追加読み込み
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(1500)

        # モバイルで「もっと見る」が出たら押す
        more = page.locator("text=もっと見る").first
        if more.count() > 0:
            try:
                more.click(timeout=1500)
                page.wait_for_timeout(1200)
            except Exception:
                pass

        if organic_rank >= MAX_ORGANIC_RANK:
            return "50~"

    return "50~"



def main():
    df = pd.read_csv(CSV_NAME, header=None, dtype=str, keep_default_na=False, encoding="utf-8-sig")

    # AZ/BA列まで確保
    df = ensure_col_count(df, COL_BA_RESULT + 1)

    # AZ4に日付
    if df.shape[0] >= 4:
        df.iat[3, COL_AZ] = TODAY_MMDD

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            user_agent=MOBILE_USER_AGENT,
            viewport=VIEWPORT,
            locale="ja-JP",
            timezone_id="Asia/Tokyo",
        )
        page = context.new_page()

        # 5行目以降：G列にEP含む行だけ処理
        for r in range(4, len(df)):
            g_val = (df.iat[r, COL_G_FLAG] or "").strip()
            if "EP" not in g_val:
                continue

            keyword = (df.iat[r, COL_J_KEYWORD] or "").strip()
            target_url = (df.iat[r, COL_F_URL] or "").strip()

            if not keyword or not target_url:
                df.iat[r, COL_BA_RESULT] = ""
                continue

            print(f"[{r+1}行目] keyword={keyword} / url={target_url}")

            try:
                rank = get_rank(page, keyword, target_url)
            except Exception as e:
                print(f"  -> ERROR: {e}")
                rank = ""

            df.iat[r, COL_BA_RESULT] = rank

            # ブロック回避（少し長め推奨）
            time.sleep(5)

        context.close()
        browser.close()

    # ====== 出力条件：G列にEP含まない行は出力しない ======
    # 1〜4行目（0〜3）は残して、それ以降はEP行だけ残す
    head = df.iloc[:4].copy()
    body = df.iloc[4:].copy()
    body_ep = body[body[COL_G_FLAG].astype(str).str.contains("EP", na=False)]
    df_out = pd.concat([head, body_ep], axis=0)

    # ※もし「1〜4行目も不要でEP行のみ出力」なら、上の3行を消してこれにする：
    # df_out = df[df[COL_G_FLAG].astype(str).str.contains("EP", na=False)].copy()

    df_out.to_csv(OUT_NAME, index=False, header=False, encoding="utf-8-sig")
    print(f"\n完了: {OUT_NAME} を出力しました")


if __name__ == "__main__":
    main()
