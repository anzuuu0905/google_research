from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

def analyze_multiple_sites(urls):
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    # Excelワークブックを作成
    wb = Workbook()
    ws = wb.active
    ws.title = "サイト解析結果"

    try:
        for site_index, url in enumerate(urls):
            # 列のオフセットを計算（4列ごと）
            col_offset = site_index * 4

            # サイトごとのヘッダーを設定
            headers = [f'タグ_{site_index+1}', f'レベル_{site_index+1}',
                      f'テキスト_{site_index+1}', f'URL_{site_index+1}']

            # ヘッダーを書き込み
            for i, header in enumerate(headers):
                cell = ws.cell(row=1, column=col_offset + i + 1)
                cell.value = header
                cell.font = Font(bold=True)

            try:
                driver.get(url)
                time.sleep(2)

                elements = driver.find_elements(By.XPATH, "//*[self::h1 or self::h2 or self::h3 or self::a]")
                current_level = 0
                seen_urls = set()
                row = 2  # データは2行目から開始

                for element in elements:
                    try:
                        tag_name = element.tag_name
                        text = ' '.join(element.text.strip().split())

                        if tag_name.startswith('h'):
                            level = int(tag_name[1])
                            current_level = level

                            if not text:
                                img_elements = element.find_elements(By.TAG_NAME, "img")
                                for img in img_elements:
                                    alt_text = img.get_attribute("alt")
                                    if alt_text:
                                        text = alt_text
                                        break

                            if text:
                                # 4列に分けて書き込み
                                ws.cell(row=row, column=col_offset + 1, value=f'H{level}')
                                ws.cell(row=row, column=col_offset + 2, value=level)
                                ws.cell(row=row, column=col_offset + 3, value=text)
                                ws.cell(row=row, column=col_offset + 4, value='')
                                row += 1

                        elif tag_name == 'a':
                            href = element.get_attribute("href")
                            if not text:
                                img_elements = element.find_elements(By.TAG_NAME, "img")
                                for img in img_elements:
                                    alt_text = img.get_attribute("alt")
                                    if alt_text:
                                        text = alt_text
                                        break

                            if href and text and href not in seen_urls:
                                seen_urls.add(href)
                                ws.cell(row=row, column=col_offset + 1, value='a')
                                ws.cell(row=row, column=col_offset + 2, value=current_level)
                                ws.cell(row=row, column=col_offset + 3, value=text)
                                ws.cell(row=row, column=col_offset + 4, value=href)
                                row += 1

                    except Exception as e:
                        continue

            except Exception as e:
                print(f"Error processing URL {url}: {str(e)}")
                continue

        # 列幅の自動調整
        for column in ws.columns:
            max_length = 0
            column = list(column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # 最大幅を50に制限
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        # ファイル保存
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'multiple_sites_analysis_{timestamp}.xlsx'
        wb.save(filename)
        print(f"Excelファイル '{filename}' を保存しました。")

    finally:
        driver.quit()

# メイン実行部分
if __name__ == "__main__":
    # Google検索結果のURLリストを使用
    from googe_research import google_search

    keyword = input("検索キーワードを入力してください: ")

    # Google検索を実行してURLを取得
    google_search_results = []
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)

        # Googleを開く
        driver.get("https://www.google.com")

        # 検索実行
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(keyword)
        search_box.send_keys(Keys.RETURN)
        time.sleep(2)

        # 検索結果を取得
        search_results = driver.find_elements(By.CSS_SELECTOR, ".tF2Cxc")

        for result in search_results[:10]:  # 最初の10件を取得
            try:
                url = result.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
                if url and not url.startswith("https://www.google.com"):
                    google_search_results.append(url)
            except:
                continue

    finally:
        driver.quit()

    # 取得したURLを解析
    analyze_multiple_sites(google_search_results)
