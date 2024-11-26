from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime

def analyze_page_content(url, filename):
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    # Excelワークブックを作成
    wb = Workbook()
    ws = wb.active
    ws.title = "解析結果"

    # ヘッダー行を追加して太字に設定
    headers = ['タグ', 'レベル', 'テキスト', 'URL']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    try:
        driver.get(url)
        time.sleep(2)

        elements = driver.find_elements(By.XPATH, "//*[self::h1 or self::h2 or self::h3 or self::a]")

        current_level = 0
        seen_urls = set()

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
                        ws.append([f'H{level}', level, text, ''])

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
                        ws.append(['a', current_level, text, href])

            except Exception as e:
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
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        # Excelファイルを保存
        wb.save(filename)
        print(f"Excelファイル '{filename}' を保存しました。")

    finally:
        driver.quit()

if __name__ == "__main__":
    url = input("URLを入力してください: ")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'website_analysis_{timestamp}.xlsx'
    analyze_page_content(url, filename)
