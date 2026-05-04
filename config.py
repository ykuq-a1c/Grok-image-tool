import json
import os
import sys


def _base_dir() -> str:
    # PyInstaller でフリーズされた exe の場合は exe のあるフォルダを使う
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


SETTINGS_FILE = os.path.join(_base_dir(), "settings.json")
STATE_FILE    = os.path.join(_base_dir(), "state.json")

DEFAULT_SETTINGS = {
    "api_provider": "xai",
    "xai_api_key": "",
    "venice_api_key": "",
    "interval_sec": 2.0,
    "max_workers": 5,
    "retry_count": 0,
    "naming_mode": "none",
    "naming_text": "",
    "last_input_dir": "",
    "last_output_dir": "",
    "last_prompt": "",
}


def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        settings = DEFAULT_SETTINGS.copy()
        settings.update(data)
        # 旧設定ファイルの api_key を xai_api_key へ移行
        if not settings["xai_api_key"] and data.get("api_key"):
            settings["xai_api_key"] = data["api_key"]
        return settings
    except Exception:
        return DEFAULT_SETTINGS.copy()


def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def load_state() -> dict:
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state: dict):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
