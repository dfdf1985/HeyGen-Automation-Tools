import glob
import subprocess
import re

def natural_sort_key(s):
    return [int(t) if t.isdigit() else t for t in re.findall(r'\d+|\D+', s)]

videos = glob.glob(r"C:\Users\User\Downloads\video*.mp4")
videos = sorted(videos, key=natural_sort_key)

if not videos:
    raise ValueError("找不到任何 video*.mp4")

with open("list.txt", "w", encoding="utf-8") as f:
    for v in videos:
        f.write(f"file '{v}'\n")

cmd = [
    "ffmpeg","-y"
    "-f", "concat",
    "-safe", "0",
    "-i", "list.txt",
    "-c", "copy",
    "output.mp4"
]

subprocess.run(cmd, check=True)
