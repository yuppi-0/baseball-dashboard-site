# %%
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# 保存先フォルダの準備
output_dir = os.path.join("/Users/zenkigenuser/野球/data/html")
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, "スタッツページ.html")

# Selenium設定
options = Options()

# 👇 まずは原因を調べるため、headless（画面非表示）をコメントアウトして画面を出してみる
# options.add_argument("--headless") 

options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--ignore-certificate-errors") # 証明書エラーを無視

# 一般的なブラウザに見せかける（User-Agentの偽装）
options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

# ChromeDriver起動
driver = webdriver.Chrome(options=options)

try:
    # 対象ページ
    url = "https://baseball.yahoo.co.jp/npb/game/2021040236/stats"
    print(f"アクセス中: {url}")
    
    # アクセス時に少し待機時間を持たせるオプション（タイムアウト対策）
    driver.set_page_load_timeout(30)
    driver.get(url)
    
    # ページが完全に読み込まれるまで少し待つ
    time.sleep(3)

    # HTML取得
    html = driver.page_source

    # 保存
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"保存完了: {output_path}")

except Exception as e:
    print(f"エラーが発生しました:\n{e}")

finally:
    driver.quit()
# %%