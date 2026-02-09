import os
import time
import requests
import pandas as pd
from pathlib import Path
from dotenv import load_dotenv
import re


load_dotenv()
HEYGEN_API_KEY = os.getenv("HEYGEN_API_KEY")

API_HOST = "https://api.heygen.com"
GENERATE_URL_V2 = f"{API_HOST}/v2/video/generate"
UPLOAD_URL_V1 = "https://upload.heygen.com/v1/asset"

TALKING_PHOTO_ID = "735b502f20074c68a13d51fe84aa0767"
VOICE_ID_EN = "cef3bc4e0a84424cafcde6f2cf466c97"
VOICE_ID_ZH = "4158cf2ef85d4ccc856aacb1c47dbb0c"

ROOT = Path(__file__).parent if "__file__" in globals() else Path(os.getcwd())
PNG_DIR = ROOT / "outputs" / "slides_png"
SCRIPT_PATH = ROOT / "scripts" / "script.xlsx"

def load_scripts(script_path: Path):
    try:
        df = pd.read_excel(script_path)
        df.columns = [c.lower() for c in df.columns]
        df["text"] = df["text"].astype(str)
        return dict(zip(df["slide"], df["text"]))
    except FileNotFoundError:
        print(f"找不到腳本檔案: {script_path}")
        return {}


def detect_voice_and_locale(text):
    if re.search(r"[\u4e00-\u9fff]", text):
        return (VOICE_ID_ZH, "zh-TW")
    return (VOICE_ID_EN, "en-US")


def upload_image(img_path: Path):
    headers = {
        "x-api-key": HEYGEN_API_KEY,
        "Content-Type": "image/png",
    }
    with open(img_path, "rb") as f:
        file_data = f.read()

    r = requests.post(UPLOAD_URL_V1, headers=headers, data=file_data, params={"type": "image"})
    if r.status_code != 200:
        print(f"Upload failed: {r.text}")
        return None
    print(f" Uploaded: {img_path.name}")
    return r.json()["data"]["id"]


def create_video(asset_id, text, slide_number):
    voice_id, locale = detect_voice_and_locale(text)
    headers = {
        "x-api-key": HEYGEN_API_KEY,
        "Content-Type": "application/json"
    }

    payload = {
        "caption": True,
        "video_inputs": [
            {
                "character": {
                    "type": "talking_photo",
                    "talking_photo_id": TALKING_PHOTO_ID,
                    "scale": 0.3,
                    "offset": {"x": 0.4, "y": 0.4},
                    "talking_style": "stable",
                    "fit": "cover"
                },
                "voice": {
                    "type": "text",
                    "voice_id": voice_id,
                    "input_text": text,
                    "speed": 1.0,
                    "locale": locale
                },
                "background": {
                    "type": "image",
                    "image_asset_id": asset_id,
                    "fit": "contain"
                }
            }
        ],
        "dimension": {
            "width": 1280,
            "height": 720
        },
        "test": False
    }

    try:
        r = requests.post(GENERATE_URL_V2, headers=headers, json=payload)
        if r.status_code == 200:
            data = r.json()
            vid = data.get("data", {}).get("video_id")
            if vid:
                print(f"Job Submitted! ID: {vid}")
                return vid

        print(f"Job failed: Status {r.status_code}. Response: {r.text[:100]}...")
    except Exception as e:
        print(f"Network Error: {e}")

    return None


def main():
    print("HeyGen 小廢片產生器")

    scripts = load_scripts(SCRIPT_PATH)
    pngs = sorted(PNG_DIR.glob("*.png"))

    if not scripts or not pngs:
        print("缺少腳本或投影片圖片，請檢查 'scripts/script.xlsx' 和 'outputs/slides_png/'")
        return

    print(f"找到 {len(pngs)} 張投影片，準備發送 {len(scripts)} 個任務...")

    for i, img_path in enumerate(pngs, 1):
        text = scripts.get(i, "")
        if not text: continue

        print(f"\n--- Processing Slide {i} ---")


        asset_id = upload_image(img_path)
        if not asset_id: continue

        create_video(asset_id, text, i)

        time.sleep(1)

    print("\n 所有任務已發送完畢！")
    print("請手動至 HeyGen 後台 (Projects) 查看進度並下載影片。")


if __name__ == "__main__":
    main()