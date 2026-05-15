"""配置文件读写。总表路径、学校名、学期等持久化到 config.json。"""

import json
import os

_CONFIG_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(_CONFIG_DIR, "config.json")

DEFAULT_CONFIG = {
    "总表路径": "",
    "学校名": "广州市番禺区番广附万博学校",
    "学期": "2025学年第二学期"
}

ASK_FOR_DOCX_MESSAGE = (
    "老师，我还不知道政治学习总文档在哪。"
    "请发送总文档，或告诉我总文档的完整路径，我会先记录下来再归档本次材料。"
)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return dict(DEFAULT_CONFIG)
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    for key, val in DEFAULT_CONFIG.items():
        if key not in config:
            config[key] = val
    return config


def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def is_configured():
    config = load_config()
    path = config.get("总表路径", "")
    return bool(path) and os.path.exists(path)


def set_docx_path(path):
    config = load_config()
    config["总表路径"] = path
    save_config(config)


def get_docx_path():
    return load_config().get("总表路径", "")


def resolve_docx_path(override_path=None):
    """返回可用总表路径；不可用时返回给老师的追问提示。"""
    if override_path:
        if os.path.exists(override_path):
            return override_path, ""
        return None, (
            f"老师，您指定的总文档路径不存在：{override_path}。"
            "请重新发送总文档，或告诉我正确的完整路径。"
        )

    path = get_docx_path()
    if not path:
        return None, ASK_FOR_DOCX_MESSAGE
    if not os.path.exists(path):
        return None, (
            f"老师，原先记录的总文档路径已失效：{path}。"
            "请重新发送总文档，或告诉我新的完整路径。"
        )
    return path, ""


def get_school():
    return load_config().get("学校名", DEFAULT_CONFIG["学校名"])


def get_semester():
    return load_config().get("学期", DEFAULT_CONFIG["学期"])
