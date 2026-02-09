import os
import time
import subprocess
import platform
import requests
import re
from google import genai  # 新版 SDK import
from google.genai import types
from pdf2image import convert_from_path
from dotenv import load_dotenv
from datetime import datetime, timedelta

# ================= 1. 環境設定 =================

load_dotenv()

# --- API Keys ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
HEYGEN_API_KEY = os.getenv("HEYGEN_API_KEY")

if not GEMINI_API_KEY or not HEYGEN_API_KEY:
    raise ValueError(" 錯誤：請確認 .env 檔案中包含 API Key")

# --- HeyGen API Endpoints ---
API_HOST = "https://api.heygen.com"
GENERATE_URL_V2 = f"{API_HOST}/v2/video/generate"
UPLOAD_URL_V1 = "https://upload.heygen.com/v1/asset"
VIDEO_STATUS_URL_V1 = f"{API_HOST}/v1/video_status.get"

# --- 專案參數 ---
GEMINI_MODEL_NAME = "gemini-2.5-flash-lite" 

TALKING_PHOTO_ID = "8c6187262e744939bb335949024e3ec5"
VOICE_ID_EN = "cef3bc4e0a84424cafcde6f2cf466c97"
VOICE_ID_ZH = "4158cf2ef85d4ccc856aacb1c47dbb0c"

INPUT_PPTX = "test.pptx"
OUTPUT_VIDEO = "final_output.mp4"
OUTPUT_DIR = "outputs"

# Windows LibreOffice 路徑 (請確認此路徑是否正確)
WINDOWS_SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

# ================= 2. 功能函數 =================

def detect_voice_id(text):
    if re.search(r"[\u4e00-\u9fff]", text):
        return VOICE_ID_ZH
    return VOICE_ID_EN

def get_soffice_command():
    system_os = platform.system()
    if system_os == "Windows":
        return WINDOWS_SOFFICE_PATH if os.path.exists(WINDOWS_SOFFICE_PATH) else "soffice"
    elif system_os == "Darwin":
        return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    return "libreoffice"

def convert_pptx_to_images(pptx_path):
    print(f" [1/5] 正在處理投影片: {pptx_path}")
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"找不到檔案: {pptx_path}")

    # === 確保輸出資料夾存在 ===
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # 1. PPTX -> PDF
    pdf_path = os.path.join(
        OUTPUT_DIR,
        os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    )

    cmd = [
        get_soffice_command(),
        "--headless",
        "--convert-to", "pdf",
        pptx_path,
        "--outdir", OUTPUT_DIR
    ]

    # 簡單判斷是否需要重新轉檔 (若 PDF 已存在且比 PPTX 新則跳過)
    if not (os.path.exists(pdf_path) and os.path.getmtime(pdf_path) > os.path.getmtime(pptx_path)):
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except Exception as e:
            raise Exception(f"LibreOffice 轉檔失敗: {e}\n請確認已安裝 LibreOffice 並設定正確路徑。")

    # 2. PDF -> Images
    print("   正在將 PDF 轉換為圖片...")
    try:
        images = convert_from_path(pdf_path, dpi=200)
    except Exception as e:
        raise Exception(f"Poppler 錯誤: {e}\n請確認已安裝 Poppler 並加入系統環境變數。")

    image_paths = []
    for i, image in enumerate(images):
        path = os.path.join(OUTPUT_DIR, f"slide_{i+1}.png")
        image.save(path, "PNG")
        image_paths.append(path)

    return image_paths

def generate_scripts(image_paths):
    """
    【更新版】使用 google-genai (v1.0+) 新版 SDK
    """
    print(f" [2/5] Gemini ({GEMINI_MODEL_NAME}) 正在看圖說故事...")
    
    # 初始化 Client
    client = genai.Client(api_key=GEMINI_API_KEY)
    
    scripts = []
    for i, img_path in enumerate(image_paths):
        print(f"   分析第 {i+1} 頁...")
        
        try:
            # 1. 上傳檔案 (新版 SDK 直接上傳，通常無需長時間等待處理)
            file_ref = client.files.upload(
                file=img_path,
                config={'display_name': f"Slide_{i}"}
            )
            
            # 2. 生成內容
            prompt = "你是專業講師。請用繁體中文(台灣)，針對這張簡報生成約 25 秒的口語講稿。直接輸出文字，不要有Markdown格式。"
            
            response = client.models.generate_content(
                model=GEMINI_MODEL_NAME,
                contents=[file_ref, prompt]
            )
            
            if response.text:
                text_content = response.text.strip()
                scripts.append(text_content)
                # print(f"   -> 生成內容: {text_content[:20]}...")
            else:
                scripts.append("（無法生成文字內容）")
                
        except Exception as e:
            print(f"   Gemini 生成錯誤: {e}")
            scripts.append(f"第 {i+1} 頁內容生成失敗，請手動補充。")
            
    return scripts

def upload_to_heygen(file_path):
    headers = {
        "X-Api-Key": HEYGEN_API_KEY,
        "Content-Type": "image/png" 
    }
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到檔案: {file_path}")

    print(f"    正在上傳背景: {file_path}")

    with open(file_path, "rb") as f:
        file_data = f.read()
        
    response = requests.post(
        UPLOAD_URL_V1, 
        headers=headers, 
        data=file_data, 
        params={"type": "image"}
    )
    
    if response.status_code == 200:
        return response.json()["data"]["id"]
    else:
        raise Exception(f"HeyGen 上傳失敗 ({response.status_code}): {response.text}")

def create_full_video(image_paths, scripts):
    print(" [3/5] 建立影片任務 (Talking Photo 模式)...")
    
    scenes = []
    for i, (img_path, script) in enumerate(zip(image_paths, scripts)):
        print(f"   正在上傳第 {i+1} 頁背景...")
        bg_asset_id = upload_to_heygen(img_path)
        
        voice_id = detect_voice_id(script)
        
        scene = {
            "character": {
                "type": "talking_photo",
                "talking_photo_id": TALKING_PHOTO_ID,
                "scale": 0.25,                  
                "offset": {"x": 0.4, "y": 0.4}, 
                "talking_style": "stable",
                "fit": "cover"
            },
            "voice": {
                "type": "text",
                "voice_id": voice_id,
                "input_text": script,
                "speed": 1.0
            },
            "background": {
                "type": "image",
                "image_asset_id": bg_asset_id,
                "fit": "contain"
            }
        }
        scenes.append(scene)

    headers = {
        "X-Api-Key": HEYGEN_API_KEY,
        "Content-Type": "application/json"
    }
    
    payload = {
        "video_inputs": scenes,
        "test": False,
        "aspect_ratio": "16:9",
        "caption": True,  # 開啟字幕生成
        "dimension": {"width": 1280, "height": 720}
    }

    response = requests.post(GENERATE_URL_V2, json=payload, headers=headers)
    data = response.json()
    
    if response.status_code == 200 and not data.get("error"):
        vid = data["data"]["video_id"]
        print(f"    任務建立成功！ID: {vid}")
        return vid
    else:
        raise Exception(f"HeyGen 任務失敗: {data}")

# === ASS 時間解析工具 ===
def parse_ass_time(time_str):
    try:
        parts = time_str.split(':')
        h = int(parts[0])
        m = int(parts[1])
        s_parts = parts[2].split('.')
        s = int(s_parts[0])
        cs = int(s_parts[1]) # centiseconds
        return datetime(1900, 1, 1, h, m, s) + timedelta(milliseconds=cs*10)
    except Exception:
        return datetime(1900, 1, 1, 0, 0, 0)

# === ASS 轉 SRT 核心函數 ===
def convert_ass_to_srt(input_path, offset_seconds=-1.0):
    print(f"    正在偵測格式並執行轉換 (時間位移 {offset_seconds} 秒)...")
    
    content = ""
    try:
        with open(input_path, "r", encoding="utf-8-sig") as f:
            content = f.read()
    except:
        with open(input_path, "r", encoding="utf-8") as f:
            content = f.read()

    lines = content.splitlines()
    srt_events = []
    
    # 解析 ASS 內容
    for line in lines:
        if line.startswith("Dialogue:"):
            try:
                parts = line.split(',', 9)
                if len(parts) < 10: continue

                start_str = parts[1].strip()
                end_str = parts[2].strip()
                text = parts[9].strip()

                start_dt = parse_ass_time(start_str)
                end_dt = parse_ass_time(end_str)
                
                delta = timedelta(seconds=offset_seconds)
                new_start = start_dt + delta
                new_end = end_dt + delta

                base_time = datetime(1900, 1, 1, 0, 0, 0)
                if new_start < base_time: new_start = base_time
                if new_end < base_time: new_end = base_time

                text = text.replace(r'\n', '\n').replace(r'\N', '\n')

                srt_events.append({
                    "start": new_start,
                    "end": new_end,
                    "text": text
                })
            except Exception as e:
                print(f"    解析跳過: {line[:30]}... ({e})")

    if not srt_events:
        print("    未偵測到 ASS 格式，無法轉換。")
        return 

    # 覆蓋原檔案為標準 SRT
    with open(input_path, "w", encoding="utf-8-sig") as f:
        for idx, event in enumerate(srt_events, 1):
            start_fmt = event['start'].strftime("%H:%M:%S,%f")[:-3]
            end_fmt = event['end'].strftime("%H:%M:%S,%f")[:-3]
            
            f.write(f"{idx}\r\n")
            f.write(f"{start_fmt} --> {end_fmt}\r\n")
            f.write(f"{event['text']}\r\n\r\n")

    print(f"    格式轉換完成！已將 ASS 轉為標準 SRT (共 {len(srt_events)} 行)。")


def download_video(video_id, output_filename):
    print(f" [4/5] 等待渲染中...")
    headers = {"X-Api-Key": HEYGEN_API_KEY}
    status_url = f"{VIDEO_STATUS_URL_V1}?video_id={video_id}"
    
    start_time = time.time()
    while True:
        try:
            resp = requests.get(status_url, headers=headers)
            data = resp.json()["data"]
            status = data["status"]
            
            if status == "completed":
                video_url = data.get("video_url")
                caption_url = data.get("caption_url")

                print("    下載影片中...")
                if video_url:
                    content = requests.get(video_url).content
                    with open(output_filename, "wb") as f:
                        f.write(content)
                    print(f"    影片已儲存: {output_filename}")

                # === 字幕下載與處理區塊 ===
                if caption_url:
                    print("    發現字幕，正在下載...")
                    try:
                        sub_resp = requests.get(caption_url)
                        sub_resp.encoding = 'utf-8'
                        content_text = sub_resp.text.strip()

                        if content_text.startswith("WEBVTT"):
                            print("    字幕為 VTT 格式，存為 .vtt")
                            vtt_filename = os.path.splitext(output_filename)[0] + ".vtt"
                            with open(vtt_filename, "w", encoding="utf-8") as f:
                                f.write(content_text)
                        else:
                            # 預設存為 SRT
                            srt_filename = os.path.splitext(output_filename)[0] + ".srt"
                            # 先存一次原始檔
                            with open(srt_filename, "w", encoding="utf-8-sig") as f:
                                f.write(content_text)
                            
                            # 呼叫 ASS 轉換與時間修正 (修正 -1.3 秒)
                            convert_ass_to_srt(srt_filename, offset_seconds=-1.3)
                            
                    except Exception as e:
                        print(f"    字幕下載或處理失敗: {e}")
                else:
                    print("    此次生成未包含字幕連結")
                # ==========================
                break

            elif status == "failed":
                print(f"    渲染失敗: {data.get('error')}")
                break
            
            elapsed = int(time.time() - start_time)
            print(f"   狀態: {status} (已耗時 {elapsed}s)...")
            time.sleep(10)
            
        except Exception as e:
            print(f"   連線錯誤: {e}")
            time.sleep(10)

# ================= 3. 主程式執行 =================
if __name__ == "__main__":
    try:
        # 確保有輸出目錄
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            
        # 開始流程
        # 1. 轉檔 (PPTX -> PNG)
        slides = convert_pptx_to_images(INPUT_PPTX)
        
        # 2. 生成講稿 (Gemini V1.0 SDK)
        generated_scripts = generate_scripts(slides)
        
        # 3. 建立影片 (HeyGen)
        vid_id = create_full_video(slides, generated_scripts)
        
        # 4. 下載與處理字幕
        output_video_path = os.path.join(OUTPUT_DIR, OUTPUT_VIDEO)
        download_video(vid_id, output_video_path)

        print("\n 恭喜！全流程執行完畢。")
        
    except Exception as e:
        print(f"\n 流程終止: {e}")