import importlib
import json
import sys
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))


def load_config_with_path(monkeypatch, config_path):
    config = importlib.import_module("config")
    monkeypatch.setattr(config, "CONFIG_PATH", str(config_path))
    return config


def test_resolve_docx_path_asks_teacher_when_config_is_missing(monkeypatch, tmp_path):
    config = load_config_with_path(monkeypatch, tmp_path / "missing-config.json")

    path, message = config.resolve_docx_path()

    assert path is None
    assert "还不知道政治学习总文档在哪" in message
    assert "请发送总文档" in message


def test_resolve_docx_path_asks_teacher_when_config_path_is_empty(monkeypatch, tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(json.dumps({"总表路径": ""}, ensure_ascii=False), encoding="utf-8")
    config = load_config_with_path(monkeypatch, config_path)

    path, message = config.resolve_docx_path()

    assert path is None
    assert "还不知道政治学习总文档在哪" in message


def test_resolve_docx_path_asks_teacher_when_config_path_no_longer_exists(monkeypatch, tmp_path):
    config_path = tmp_path / "config.json"
    missing_docx = tmp_path / "missing.docx"
    config_path.write_text(json.dumps({"总表路径": str(missing_docx)}, ensure_ascii=False), encoding="utf-8")
    config = load_config_with_path(monkeypatch, config_path)

    path, message = config.resolve_docx_path()

    assert path is None
    assert "原先记录的总文档路径已失效" in message
    assert str(missing_docx) in message


def test_resolve_docx_path_returns_existing_override(monkeypatch, tmp_path):
    config = load_config_with_path(monkeypatch, tmp_path / "missing-config.json")
    docx = tmp_path / "summary.docx"
    docx.write_bytes(b"placeholder")

    path, message = config.resolve_docx_path(str(docx))

    assert path == str(docx)
    assert message == ""
