import os
import io
import time
import threading
import requests
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}

MIME_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
}


def collect_image_files(input_path: str) -> list[str]:
    """フォルダまたは単一ファイルから対象画像のパスリストを返す"""
    if os.path.isfile(input_path):
        ext = os.path.splitext(input_path)[1].lower()
        if ext in SUPPORTED_EXTENSIONS:
            return [input_path]
        return []
    files = []
    for entry in sorted(os.scandir(input_path), key=lambda e: e.name):
        if entry.is_file():
            ext = os.path.splitext(entry.name)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                files.append(entry.path)
    return files


def build_output_filename(original_path: str, naming_mode: str, naming_text: str, ext: str) -> str:
    """命名規則に従って出力ファイル名を生成する"""
    stem = os.path.splitext(os.path.basename(original_path))[0]
    if naming_mode == "prefix":
        return f"{naming_text}{stem}{ext}"
    elif naming_mode == "suffix":
        return f"{stem}{naming_text}{ext}"
    return f"{stem}{ext}"



TARGET_BYTES = 2 * 1024 * 1024  # 2MB


def _ext_from_image_bytes(data: bytes) -> str:
    """画像バイナリから拡張子を判定する。判定できない場合は .png を返す。"""
    if data[:3] == b'\xff\xd8\xff':
        return ".jpg"
    if data[:4] == b'\x89PNG':
        return ".png"
    if len(data) >= 12 and data[:4] == b'RIFF' and data[8:12] == b'WEBP':
        return ".webp"
    return ".png"


def _output_exists(output_dir: str, original_path: str, naming_mode: str, naming_text: str) -> bool:
    """出力ファイルが何らかの拡張子で既に存在するか確認する（レジューム用）"""
    for ext in SUPPORTED_EXTENSIONS:
        name = build_output_filename(original_path, naming_mode, naming_text, ext)
        if os.path.exists(os.path.join(output_dir, name)):
            return True
    return False


def _get_resized_b64(filepath: str) -> tuple[str, str]:
    """
    ファイルサイズが TARGET_BYTES を超える場合、メモリ上でリサイズしてBase64を返す。
    超えない場合はそのまま読み込んでBase64を返す。
    戻り値: (base64文字列, MIMEタイプ)
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".jpg", ".jpeg"):
        save_format, save_mime = "JPEG", "image/jpeg"
    elif ext == ".webp":
        save_format, save_mime = "WEBP", "image/webp"
    else:
        save_format, save_mime = "PNG", "image/png"

    file_size = os.path.getsize(filepath)

    if file_size <= TARGET_BYTES:
        with open(filepath, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8"), MIME_TYPES.get(ext, "image/jpeg")

    img = Image.open(filepath)
    if save_format == "JPEG" and img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")

    scale = (TARGET_BYTES / file_size) ** 0.5
    new_w = max(1, int(img.width * scale))
    new_h = max(1, int(img.height * scale))
    resized = img.resize((new_w, new_h), Image.LANCZOS)

    save_kwargs = {"quality": 85} if save_format in ("JPEG", "WEBP") else {}
    buf = io.BytesIO()
    resized.save(buf, format=save_format, **save_kwargs)

    # 計算が外れてまだオーバーしていたら 0.9 倍ずつ縮小して再試行
    while buf.tell() > TARGET_BYTES and new_w > 1 and new_h > 1:
        new_w = max(1, int(new_w * 0.9))
        new_h = max(1, int(new_h * 0.9))
        resized = img.resize((new_w, new_h), Image.LANCZOS)
        buf = io.BytesIO()
        resized.save(buf, format=save_format, **save_kwargs)

    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8"), save_mime


def call_api(api_key: str, image_path: str, prompt: str,
             api_provider: str = "xai") -> tuple[bool, bytes | None, str]:
    """API呼び出しのディスパッチャー。戻り値: (成功フラグ, 画像バイナリ or None, エラーメッセージ)"""
    if api_provider == "venice":
        return _call_api_venice(api_key, image_path, prompt)
    return _call_api_xai(api_key, image_path, prompt)


def _call_api_xai(api_key: str, image_path: str, prompt: str) -> tuple[bool, bytes | None, str]:
    url = "https://api.x.ai/v1/images/edits"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        b64_input, mime = _get_resized_b64(image_path)

        payload = {
            "model": "grok-imagine-image",
            "prompt": prompt,
            "image": {
                "url": f"data:{mime};base64,{b64_input}",
                "type": "image_url",
            },
            "response_format": "b64_json",
        }

        response = requests.post(url, headers=headers, json=payload, timeout=60)

        if response.status_code == 200:
            result = response.json()
            b64_output = result["data"][0]["b64_json"]
            return True, base64.b64decode(b64_output), ""
        else:
            error_type = _classify_error(response.status_code)
            return False, None, f"{error_type}|{response.status_code} {response.text[:200]}"

    except requests.exceptions.Timeout:
        return False, None, "timeout|Request timed out"
    except Exception as e:
        return False, None, f"unknown|{str(e)}"


def _call_api_venice(api_key: str, image_path: str, prompt: str) -> tuple[bool, bytes | None, str]:
    url = "https://api.venice.ai/api/v1/image/edit"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        b64_input, mime = _get_resized_b64(image_path)

        payload = {
            "model": "grok-imagine-edit",
            "prompt": prompt,
            "image": f"data:{mime};base64,{b64_input}",
            "safe_mode": False,
        }

        response = requests.post(url, headers=headers, json=payload, timeout=60)

        if response.status_code == 200:
            return True, response.content, ""
        else:
            error_type = _classify_error(response.status_code)
            return False, None, f"{error_type}|{response.status_code} {response.text[:200]}"

    except requests.exceptions.Timeout:
        return False, None, "timeout|Request timed out"
    except Exception as e:
        return False, None, f"unknown|{str(e)}"


def _classify_error(status_code: int) -> str:
    if status_code == 429:
        return "rate_limit"
    elif status_code in (400, 451):
        return "moderation"
    elif status_code >= 500:
        return "server_error"
    return "client_error"


def run_batch(
    api_key: str,
    image_files: list[str],
    prompt: str,
    output_dir: str,
    naming_mode: str,
    naming_text: str,
    interval_sec: float,
    max_workers: int,
    on_progress,    # callback(done, total, success, fail, skip)
    stop_flag,      # callable -> bool
    api_provider: str = "xai",
):
    """並列バッチ処理のメインループ"""
    total = len(image_files)
    os.makedirs(output_dir, exist_ok=True)

    # スレッドセーフなカウンター
    counters = {"done": 0, "success": 0, "fail": 0, "skip": 0}
    counter_lock = threading.Lock()
    failed_list = []
    failed_lock = threading.Lock()

    def process_one(filepath: str):
        if stop_flag():
            return

        # 既存ファイルはスキップ（レジューム対応・拡張子不問）
        if _output_exists(output_dir, filepath, naming_mode, naming_text):
            with counter_lock:
                counters["skip"] += 1
                counters["done"] += 1
                d, s, f, sk = counters["done"], counters["success"], counters["fail"], counters["skip"]
            on_progress(d, total, s, f, sk)
            return

        if stop_flag():
            return

        # 送信時刻を記録してAPIを呼ぶ
        send_time = time.monotonic()
        ok, image_data, error_msg = call_api(api_key, filepath, prompt, api_provider)

        if ok:
            ext = _ext_from_image_bytes(image_data)
            out_name = build_output_filename(filepath, naming_mode, naming_text, ext)
            out_path = os.path.join(output_dir, out_name)
            with open(out_path, "wb") as f:
                f.write(image_data)
            with counter_lock:
                counters["success"] += 1
                counters["done"] += 1
                d, s, f, sk = counters["done"], counters["success"], counters["fail"], counters["skip"]
        else:
            parts = error_msg.split("|", 1)
            error_type = parts[0]
            error_detail = parts[1] if len(parts) > 1 else ""
            with failed_lock:
                failed_list.append((filepath, error_type, error_detail))
            with counter_lock:
                counters["fail"] += 1
                counters["done"] += 1
                d, s, f, sk = counters["done"], counters["success"], counters["fail"], counters["skip"]

        on_progress(d, total, s, f, sk)

        # 送信時刻から interval_sec が経過するまで待機（ワーカー個別クールダウン）
        remaining = interval_sec - (time.monotonic() - send_time)
        if remaining > 0 and not stop_flag():
            time.sleep(remaining)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_one, fp): fp for fp in image_files}
        for future in as_completed(futures):
            if stop_flag():
                # 残りのタスクをキャンセル
                for f in futures:
                    f.cancel()
                break

    return failed_list


