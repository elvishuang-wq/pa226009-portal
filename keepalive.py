"""
真正喚醒 Streamlit Cloud App 的腳本（不是單純 ping，而是用無頭瀏覽器打開頁面，
偵測「Yes, get this app back up!」按鈕並點擊）。

需求套件:
    pip install playwright
    playwright install chromium

使用方式：
    python keepalive.py https://your-app.streamlit.app
"""

import sys
from playwright.sync_api import sync_playwright


def wake_app(url: str):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        print(f"造訪 {url} ...")
        page.goto(url, wait_until="networkidle", timeout=30000)

        # Streamlit 睡眠中的畫面會有這個喚醒按鈕
        wake_button = page.get_by_text("get this app back up", exact=False)
        try:
            wake_button.wait_for(timeout=5000)
            print("偵測到 App 正在睡眠，點擊喚醒按鈕...")
            wake_button.click()
            page.wait_for_timeout(15000)  # 等待 App 重新啟動
            print("已送出喚醒請求")
        except Exception:
            print("App 目前是醒著的，不需要喚醒")

        browser.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python keepalive.py <streamlit_app_url>")
        sys.exit(1)
    wake_app(sys.argv[1])
