from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

def google_search(keyword):
    # Chromeドライバーの設定
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        # Googleを開く
        driver.get("https://www.google.com")

        # 検索ボックスを見つけて検索語を入力
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(keyword)
        search_box.send_keys(Keys.RETURN)

        # ページの読み込みを待つ
        time.sleep(2)

        # 検索結果を取得
        valid_results = []
        seen_urls = set()

        # スクロールして追加の結果を読み込む
        for _ in range(3):  # 必要に応じてスクロール回数を増やす
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

        search_results = driver.find_elements(By.CSS_SELECTOR, ".tF2Cxc")  # Google検索結果のセレクター


        # 検索結果から情報を抽出
        for result in search_results:
            try:
                title_element = result.find_element(By.CSS_SELECTOR, "h3")
                url_element = result.find_element(By.CSS_SELECTOR, "a")

                title = title_element.text
                url = url_element.get_attribute("href")

                if (title.strip() and
                    not url.startswith("https://www.google.com") and
                    url not in seen_urls):
                    valid_results.append((title, url))
                    seen_urls.add(url)

                    if len(valid_results) == 11:  # 11件で停止
                        break

            except:
                continue

        # 結果を表示
        for i, (title, url) in enumerate(valid_results, 1):
            print(f"{i}. タイトル: {title}")
            print(f"   URL: {url}\n")

        print(f"取得した総検索結果: {len(valid_results)}件")

    finally:
        # ブラウザを閉じる
        driver.quit()

# プログラムを実行
keyword = input("検索キーワードを入力してください: ")
google_search(keyword)
