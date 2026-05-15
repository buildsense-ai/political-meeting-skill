import importlib
import json
import sys
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))


class FakeOfficeCli:
    def __init__(self):
        self.docs = {}
        self.images = {}
        self.calls = []
        self.reject_past_end_index = False
        self.fail_format_set = False

    def set_doc(self, path, paragraphs):
        self.docs[path] = list(paragraphs)
        self.images[path] = set()

    def __call__(self, args, timeout=60):
        self.calls.append(args)
        command = args[0]
        path = args[1]

        if command == "view":
            return "\n".join(self.docs.get(path, []))

        if command == "query" and args[2] == "picture":
            results = []
            for index in sorted(self.images.get(path, set())):
                results.append({
                    "path": f"/body/p[{index + 1}]/r[1]",
                    "type": "picture",
                    "text": self.docs[path][index],
                    "format": {"alt": self.docs[path][index]},
                })
            return json.dumps({"success": True, "data": {"matches": len(results), "results": results}})

        if command == "add":
            doc = self.docs.setdefault(path, [])
            item_type = args[args.index("--type") + 1]
            index = None
            if "--index" in args:
                index = int(args[args.index("--index") + 1])
                if self.reject_past_end_index and index >= len(doc):
                    raise RuntimeError(f"index {index} is past end of document with {len(doc)} paragraphs")

            if item_type == "paragraph":
                prop = next(arg for arg in args if arg.startswith("text="))
                value = prop[len("text="):]
            elif item_type in ("image", "picture"):
                prop = next(arg for arg in args if arg.startswith("file=") or arg.startswith("src="))
                value = f"[IMAGE:{Path(prop.split('=', 1)[1]).name}]"
            else:
                raise AssertionError(f"unexpected add type: {item_type}")

            if index is None or index >= len(doc):
                doc.append(value)
                inserted_at = len(doc) - 1
            else:
                doc.insert(index, value)
                inserted_at = index
                self.images[path] = {i + 1 if i >= inserted_at else i for i in self.images.get(path, set())}
            if item_type in ("image", "picture"):
                self.images.setdefault(path, set()).add(inserted_at)
            return ""

        if command == "set":
            if "--prop" not in args:
                return ""
            props = [args[i + 1] for i, arg in enumerate(args) if arg == "--prop"]
            text_props = [prop for prop in props if prop.startswith("text=")]
            if self.fail_format_set and not text_props and "/r[1]" in args[2]:
                raise RuntimeError("OfficeCLI failed: set\nError: Path not found: /body/p[999]/r[1]. Available at /body: p[1]..p[10]")
            if text_props:
                selector = args[2]
                marker = "/body/p["
                start = selector.index(marker) + len(marker)
                end = selector.index("]", start)
                index = int(selector[start:end]) - 1
                self.docs[path][index] = text_props[-1][len("text="):]
            return ""

        raise AssertionError(f"unexpected command: {args}")


def load_modules(monkeypatch):
    officecli_helper = importlib.import_module("officecli_helper")
    insert_material = importlib.import_module("insert_material")
    fake = FakeOfficeCli()
    monkeypatch.setattr(officecli_helper, "_run", fake)
    monkeypatch.setattr(
        insert_material,
        "extract_first_title",
        lambda filepath, fmt=None: fake.docs[filepath][0],
    )
    monkeypatch.setattr(
        insert_material,
        "extract_docx_paragraphs",
        lambda filepath: fake.docs[filepath],
    )
    return officecli_helper, insert_material, fake


def test_learning_content_first_creates_week_scaffold_before_content(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")

    fake.set_doc(summary, ["2025学年第二学期第七、八周政治学习", "旧内容"])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    doc = fake.docs[summary]
    title_idx = doc.index("2025学年第二学期第九、十周政治学习")
    brief_idx = next(i for i, text in enumerate(doc) if text.startswith("广州市番禺区番广附万博学校2025学年第二学期第九、十周政治学习："))
    photo_idx = doc.index("会议照片：")
    content_idx = doc.index("习近平总书记会见中国国民党主席郑丽文学习材料")
    signin_idx = doc.index("会议签到表")

    assert doc[signin_idx - 1] == "会议记录"
    assert signin_idx < title_idx < brief_idx < photo_idx < content_idx


def test_appending_learning_content_to_last_existing_week_does_not_use_past_end_index(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    fake.reject_past_end_index = True
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week11.docx")

    fake.set_doc(summary, [
        "会议记录",
        "会议签到表",
        "2025学年第二学期第十一、十二周政治学习",
        "广州市番禺区番广附万博学校2025学年第二学期第十一、十二周政治学习：已有材料",
        "会议照片：",
        "已有正文",
    ])
    fake.set_doc(source, ["补充学习材料标题", "补充学习材料正文"])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第十一、十二周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    assert fake.docs[summary][-2:] == ["补充学习材料标题", "补充学习材料正文"]


def test_photo_arriving_after_learning_content_is_inserted_before_content(monkeypatch, tmp_path):
    officecli_helper, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")
    image = str(tmp_path / "photo.jpg")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )
    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "会议照片",
        image,
    )

    assert success, message
    doc = fake.docs[summary]
    photo_marker_idx = doc.index("会议照片：")
    image_idx = doc.index("[IMAGE:photo.jpg]")
    content_idx = doc.index("习近平总书记会见中国国民党主席郑丽文学习材料")

    assert photo_marker_idx < image_idx < content_idx
    status = officecli_helper.detect_week_status(summary, "第九、十周政治学习")
    assert status["会议照片"] is True


def test_photo_placeholder_does_not_count_as_uploaded_photo(monkeypatch, tmp_path):
    officecli_helper, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    status = officecli_helper.detect_week_status(summary, "第九、十周政治学习")

    assert status["会议照片"] is False
    assert status["学习内容"] is True
    assert status["签到表"] is False


def test_meeting_photo_and_brief_do_not_count_as_learning_content(monkeypatch, tmp_path):
    officecli_helper, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    image = str(tmp_path / "meeting-photo-with-long-alt-text.jpg")
    week_title = officecli_helper.DEFAULT_SEMESTER + "绗節銆佸崄鍛ㄦ斂娌诲涔?"
    week_keyword = week_title.replace(officecli_helper.DEFAULT_SEMESTER, "").strip()
    material_order = officecli_helper.MATERIAL_ORDER

    fake.set_doc(summary, [
        material_order[0],
        material_order[1],
        week_title,
        officecli_helper.DEFAULT_SCHOOL + "x" * 40,
    ])

    success, message = insert_material.insert_material(summary, week_title, material_order[4], image)

    status = officecli_helper.detect_week_status(summary, week_keyword)

    assert success, message
    assert material_order[1] in message
    assert material_order[5] in message
    assert status[material_order[4]] is True
    assert status[material_order[5]] is False
    assert status[material_order[1]] is False


def test_signin_sheet_arriving_after_learning_content_is_inserted_before_title(monkeypatch, tmp_path):
    officecli_helper, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")
    signin = str(tmp_path / "signin.jpg")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )
    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "签到表",
        signin,
    )

    assert success, message
    doc = fake.docs[summary]
    meeting_record_idx = doc.index("会议记录")
    signin_marker_idx = doc.index("会议签到表")
    signin_image_idx = doc.index("[IMAGE:signin.jpg]")
    title_idx = doc.index("2025学年第二学期第九、十周政治学习")

    assert meeting_record_idx < signin_marker_idx < signin_image_idx < title_idx
    status = officecli_helper.detect_week_status(summary, "第九、十周政治学习")
    assert status["签到表"] is True


def test_signin_placeholder_does_not_count_when_only_previous_picture_exists(monkeypatch, tmp_path):
    officecli_helper, _, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")

    fake.set_doc(summary, [
        "2025学年第二学期第七、八周政治学习",
        "[IMAGE:old-photo.jpg]",
        "会议记录",
        "会议签到表",
        "2025学年第二学期第九、十周政治学习",
        "广州市番禺区番广附万博学校2025学年第二学期第九、十周政治学习：习近平总书记会见中国国民党主席郑丽文学习材料",
        "会议照片：",
        "习近平总书记会见中国国民党主席郑丽文学习材料",
    ])
    fake.images[summary] = {1}

    status = officecli_helper.detect_week_status(summary, "第九、十周政治学习")

    assert status["签到表"] is False
    assert status["学习内容"] is True


def test_docx_learning_content_does_not_use_officecli_to_read_source(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    def fail_if_called(*args, **kwargs):
        raise AssertionError("source docx should be read without OfficeCLI copy_docx_content")

    monkeypatch.setattr(insert_material, "copy_docx_content", fail_if_called)
    monkeypatch.setattr(
        insert_material,
        "extract_docx_paragraphs",
        lambda filepath: fake.docs[filepath],
        raising=False,
    )

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    assert "习近平总书记会见中国国民党主席郑丽文学习材料" in fake.docs[summary]


def test_retried_learning_content_is_not_appended_twice(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week11.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, [
        "教育部关于全面推进健康学校建设的指导意见",
        "学生身心健康从软要求走向硬约束",
    ])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第十一周政治学习",
        "学习内容",
        source,
    )
    assert success, message

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第十一周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    assert "未重复追加" in message
    assert fake.docs[summary].count("教育部关于全面推进健康学校建设的指导意见") == 1
    assert fake.docs[summary].count("学生身心健康从软要求走向硬约束") == 1


def test_duplicate_paragraphs_in_source_are_inserted_once(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, [
        "习近平总书记会见中国国民党主席郑丽文",
        "2026-04-10来源：“学习强国”学习平台",
        "习近平总书记会见中国国民党主席郑丽文",
        "4月10日上午，中共中央总书记习近平在北京会见郑丽文主席率领的中国国民党访问团。",
        "4月10日上午，中共中央总书记习近平在北京会见郑丽文主席率领的中国国民党访问团。",
    ])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    content = fake.docs[summary][fake.docs[summary].index("会议照片：") + 1:]
    assert content.count("习近平总书记会见中国国民党主席郑丽文") == 1
    assert content.count("4月10日上午，中共中央总书记习近平在北京会见郑丽文主席率领的中国国民党访问团。") == 1


def test_doc_learning_content_auto_converts_to_docx_before_binary_fallback(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week11.doc")
    converted = str(tmp_path / "week11-converted.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["乱码源文件不应直接使用"])
    fake.set_doc(converted, ["自动转换后的标题", "自动转换后的正文"])

    monkeypatch.setattr(insert_material, "convert_doc_to_docx", lambda filepath: converted)

    def fail_if_called(*args, **kwargs):
        raise AssertionError(".doc should try auto conversion before binary text extraction")

    monkeypatch.setattr(insert_material, "extract_doc_text", fail_if_called)

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第十一、十二周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    assert "自动转换后的标题" in fake.docs[summary]
    assert "自动转换后的正文" in fake.docs[summary]


def test_brief_intro_contains_only_material_titles_not_body(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week9.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, [
        "《关于开展基础教育规范管理提升年行动的通知》",
        "作为规范管理的深化举措，文件聚焦四大重点任务，推动基础教育治理提质增效。",
        "《义务教育、普通高中课程方案和课程标准日常修订版》是正文中提到的延伸材料，不应进入简要介绍。",
        "教师关注点：教学行为规范将更加精细化、智能化。",
    ])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    brief = next(text for text in fake.docs[summary] if text.startswith("广州市番禺区番广附万博学校"))
    assert brief == "广州市番禺区番广附万博学校2025学年第二学期第九、十周政治学习：《关于开展基础教育规范管理提升年行动的通知》"
    assert "作为规范管理的深化举措" not in brief
    assert "义务教育" not in brief
    assert "教师关注点" not in brief


def test_inserted_learning_content_uses_body_font_size(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week10.docx")

    fake.set_doc(summary, [])
    fake.set_doc(source, ["习近平总书记会见中国国民党主席郑丽文学习材料", "学习材料正文"])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第九、十周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    content_idx = fake.docs[summary].index("习近平总书记会见中国国民党主席郑丽文学习材料")
    set_calls = [
        call for call in fake.calls
        if call[:3] == ["set", summary, f"/body/p[{content_idx + 1}]/r[1]"]
    ]
    props = [call[i + 1] for call in set_calls for i, arg in enumerate(call) if arg == "--prop"]
    assert "font=宋体" in props
    assert "size=14pt" in props


def test_learning_content_still_succeeds_when_format_path_is_unavailable(monkeypatch, tmp_path):
    _, insert_material, fake = load_modules(monkeypatch)
    fake.fail_format_set = True
    summary = str(tmp_path / "summary.docx")
    source = str(tmp_path / "week11.docx")

    fake.set_doc(summary, [
        "会议记录",
        "会议签到表",
        "2025学年第二学期第十一周政治学习",
        "广州市番禺区番广附万博学校2025学年第二学期第十一周政治学习：已有材料",
        "会议照片：",
        "已有正文",
    ])
    fake.set_doc(source, ["教育部关于全面推进健康学校建设的指导意见", "补充正文"])

    success, message = insert_material.insert_material(
        summary,
        "2025学年第二学期第十一周政治学习",
        "学习内容",
        source,
    )

    assert success, message
    assert "教育部关于全面推进健康学校建设的指导意见" in fake.docs[summary]
    assert "补充正文" in fake.docs[summary]


def test_oc_add_image_uses_current_officecli_picture_src_properties(monkeypatch):
    officecli_helper = importlib.import_module("officecli_helper")
    calls = []

    monkeypatch.setattr(officecli_helper, "_run", lambda args: calls.append(args) or "")

    officecli_helper.oc_add_image("summary.docx", "photo.jpg", index=3)

    args = calls[0]
    assert args[args.index("--type") + 1] == "picture"
    assert "src=photo.jpg" in [args[i + 1] for i, arg in enumerate(args) if arg == "--prop"]
