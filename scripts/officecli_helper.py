"""OfficeCLI 命令封装 + 文档缓存层。提供 Pythonic 接口操作 docx。"""

import json
import os
import re
import shutil
import subprocess
import time


def _find_officecli():
    candidates = []
    if os.environ.get("OFFICECLI_PATH"):
        candidates.append(os.environ["OFFICECLI_PATH"])
    if os.environ.get("LOCALAPPDATA"):
        candidates.append(os.path.join(os.environ["LOCALAPPDATA"], "OfficeCLI", "officecli.exe"))
    if os.environ.get("USERPROFILE"):
        candidates.append(os.path.join(os.environ["USERPROFILE"], "AppData", "Local", "OfficeCLI", "officecli.exe"))
    which = shutil.which("officecli")
    if which:
        candidates.append(which)

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return "officecli"


OFFICECLI = _find_officecli()

MATERIAL_ORDER = ["会议记录", "签到表", "标题", "简要介绍", "会议照片", "学习内容"]
DEFAULT_SCHOOL = "广州市番禺区番广附万博学校"
DEFAULT_SEMESTER = "2025学年第二学期"
BODY_FONT = "宋体"
BODY_SIZE = "14pt"


class DocumentLockedError(RuntimeError):
    """目标 Word 文档正被 Word 或其他程序占用。"""

    def __init__(self, path, detail):
        self.path = path
        self.detail = detail
        super().__init__(
            f"老师，总文档现在正被 Word 或其他程序打开，暂时不能归档：{path}\n"
            "请先关闭这份总文档，再重新发送材料或让我继续归档。"
        )


def _run(args, timeout=60):
    """执行 officecli 命令。"""
    cmd = [OFFICECLI] + args
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout,
            encoding="utf-8", errors="replace"
        )
    except FileNotFoundError as exc:
        raise RuntimeError(
            "OfficeCLI 执行文件未找到。请确认已安装 OfficeCLI，或设置 OFFICECLI_PATH 指向 officecli.exe。"
        ) from exc
    if result.returncode != 0 and result.stderr:
        if "being used by another process" in result.stderr or "cannot access the file" in result.stderr:
            docx_path = args[1] if len(args) > 1 else ""
            raise DocumentLockedError(docx_path, result.stderr)
        raise RuntimeError(f"OfficeCLI failed: {' '.join(cmd)}\n{result.stderr}")
    return result.stdout


def release_docx(path):
    """尽量释放 OfficeCLI resident/watch 句柄；没有打开时忽略错误。"""
    for command in ("close", "unwatch"):
        try:
            subprocess.run(
                [OFFICECLI, command, path],
                capture_output=True,
                text=True,
                timeout=10,
                encoding="utf-8",
                errors="replace",
            )
        except Exception:
            pass


def wait_until_file_available(path, timeout_seconds=5):
    """等待文件可被当前进程独占打开，用于避开 OfficeCLI 刚写完的短暂锁。"""
    deadline = time.time() + timeout_seconds
    while time.time() <= deadline:
        try:
            with open(path, "ab"):
                return True
        except OSError:
            time.sleep(0.2)
    return False


class DocxCache:
    """缓存文档内容，避免反复调用 OfficeCLI 读取。"""

    def __init__(self, path):
        self.path = path
        self._lines = None  # 完整文本行列表
        self._texts = None  # 纯文本列表（每段）
        self._picture_indices = None

    def _load(self):
        if self._lines is None:
            raw = _run(["view", self.path, "text"]).strip()
            self._lines = raw.split("\n") if raw else []
            self._texts = []
            for line in self._lines:
                idx = line.find("]] ")
                self._texts.append(line[idx + 3:] if idx > 0 else line)

    def invalidate(self):
        """文档被修改后，标记缓存过期。"""
        self._lines = None
        self._texts = None
        self._picture_indices = None

    @property
    def lines(self):
        self._load()
        return self._lines

    @property
    def texts(self):
        self._load()
        return self._texts

    def __len__(self):
        self._load()
        return len(self._texts)

    def text_at(self, i):
        self._load()
        if 0 <= i < len(self._texts):
            return self._texts[i].strip()
        return ""

    def find(self, text, exclude_if_contains=None):
        """查找第一个包含 text 的段落索引。"""
        self._load()
        for i, t in enumerate(self._texts):
            if exclude_if_contains and exclude_if_contains in t:
                continue
            if text in t:
                return i
        return -1

    def find_all_weeks(self):
        """返回所有周标题的 (索引, 文本) 列表。"""
        self._load()
        results = []
        for i, t in enumerate(self._texts):
            s = t.strip()
            if DEFAULT_SCHOOL in s:
                continue
            if "政治学习" in s and ("学期" in s or "周" in s):
                results.append((i, s))
        return results

    def week_boundary(self, week_keyword):
        """返回 (标题索引, 下一周起始索引)。"""
        anchor = self.find(week_keyword, exclude_if_contains=DEFAULT_SCHOOL)
        if anchor < 0:
            return -1, -1
        all_titles = self.find_all_weeks()
        total = len(self._texts)
        next_start = total
        for idx, _ in all_titles:
            if idx > anchor:
                next_start = idx
                break
        return anchor, next_start

    def last_content_para(self, start, end):
        """在 [start, end) 范围内最后一个非空段落索引。"""
        last = start
        for i in range(start, end):
            if self._texts[i].strip():
                last = i
        return last

    def has_marker(self, text, start, end):
        """检查 [start, end) 范围内是否存在包含 text 的段落。"""
        for i in range(max(0, start), min(end, len(self._texts))):
            if text in self._texts[i]:
                return True
        return False

    def find_marker(self, text, start, end):
        """在 [start, end) 范围内查找包含 text 的段落索引。"""
        for i in range(max(0, start), min(end, len(self._texts))):
            if text in self._texts[i]:
                return i
        return -1

    def picture_indices(self):
        """返回文档中包含图片的段落索引集合。"""
        self._load()
        if self._picture_indices is not None:
            return self._picture_indices

        raw = _run(["query", self.path, "picture", "--json"]).strip()
        try:
            data = json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            self._picture_indices = set()
            return self._picture_indices

        results = data.get("data", {}).get("results", [])
        para_id_to_index = {}
        for i, line in enumerate(self._lines):
            match = re.search(r"paraId=([A-Za-z0-9]+)", line)
            if match:
                para_id_to_index[match.group(1)] = i

        indices = set()
        for result in results:
            path = result.get("path", "")
            id_match = re.search(r"/body/p\[@paraId=([A-Za-z0-9]+)\]", path)
            if id_match and id_match.group(1) in para_id_to_index:
                indices.add(para_id_to_index[id_match.group(1)])
                continue

            index_match = re.search(r"/body/p\[(\d+)\]", path)
            if index_match:
                indices.add(int(index_match.group(1)) - 1)

        self._picture_indices = indices
        return self._picture_indices


def oc_add_para(docx_path, text, index=None):
    """在文档末尾或指定位置插入段落。index=None 追加到末尾。"""
    args = ["add", docx_path, "/body", "--type", "paragraph", "--prop", f"text={text}"]
    if index is not None:
        args += ["--index", str(index)]
    _run(args)


def oc_add_image(docx_path, image_path, index=None, width_inches=5.5):
    """在末尾或指定位置插入图片。"""
    width_cm = width_inches * 2.54
    args = ["add", docx_path, "/body", "--type", "picture",
            "--prop", f"src={image_path}", "--prop", f"width={width_cm}cm"]
    if index is not None:
        args += ["--index", str(index)]
    _run(args)


def oc_set_format(docx_path, index, font=None, size=None, bold=None):
    """设置段落中 run 的格式。index 为 0-based。"""
    args = ["set", docx_path, f"/body/p[{index + 1}]/r[1]"]
    if font:
        args += ["--prop", f"font={font}"]
    if size:
        args += ["--prop", f"size={size}"]
    if bold is not None:
        args += ["--prop", f"bold={str(bold).lower()}"]
    try:
        _run(args)
        return True
    except RuntimeError as exc:
        if "Path not found" in str(exc) and "/body/p[" in str(exc):
            return False
        raise


def oc_replace_text(docx_path, index, text):
    """替换段落的全部文本。"""
    _run(["set", docx_path, f"/body/p[{index + 1}]", "--prop", f"text={text}"])


def _normalize_add_index(index, current_length):
    """OfficeCLI does not accept an index at or beyond the current paragraph count."""
    if index is None or index >= current_length:
        return None, current_length
    return index, index


def oc_copy_paras(src_path, dst_path, start_idx, end_idx, after_idx):
    """从 src 复制 [start, end) 段落到 dst 的 after_idx 之后。"""
    from lxml import etree
    import tempfile

    # 先处理 dst 文档
    # OfficeCLI 不支持跨文档复制，改用读取文本再插入
    src_doc = DocxCache(src_path)
    insert_pos = after_idx
    for i in range(start_idx, end_idx):
        t = src_doc.text_at(i)
        if t:
            insert_pos += 1
            oc_add_para(dst_path, t, insert_pos)


# ─── 业务逻辑函数（使用缓存）─────────────────────────────


def _is_learning_content_text(text):
    """判断段落是否是学习材料正文。"""
    if not text or len(text) <= 20:
        return False
    if DEFAULT_SCHOOL in text:
        return False
    if text.startswith("[IMAGE:"):
        return False
    if MATERIAL_ORDER[0] in text or MATERIAL_ORDER[1] in text or MATERIAL_ORDER[4] in text:
        return False
    return True


def detect_week_status(path, week_keyword):
    """扫描某周区域，返回各材料是否已插入。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        return {item: False for item in MATERIAL_ORDER}

    status = {item: False for item in MATERIAL_ORDER}
    status["标题"] = True
    status["会议记录"] = True

    prev_start = 0
    for idx, _ in doc.find_all_weeks():
        if idx < anchor:
            prev_start = idx

    picture_indices = doc.picture_indices()

    meeting_record = -1
    for i in range(prev_start, anchor):
        if doc.text_at(i) == "会议记录":
            meeting_record = i

    signin_marker = -1
    for i in range(max(prev_start, meeting_record + 1), anchor):
        if "会议签到表" in doc.text_at(i):
            signin_marker = i

    if signin_marker >= 0 and any(signin_marker < pic < anchor for pic in picture_indices):
        status["签到表"] = True

    photo_marker = doc.find_marker("会议照片：", anchor + 1, next_start)
    if photo_marker >= 0 and any(photo_marker < pic < next_start for pic in picture_indices):
        status["会议照片"] = True

    for i in range(anchor + 1, next_start):
        t = doc.text_at(i)
        if DEFAULT_SCHOOL in t and len(t) > 30:
            status["简要介绍"] = True

    learning_start = photo_marker + 1 if photo_marker >= 0 else anchor + 1
    for i in range(learning_start, next_start):
        if i in picture_indices:
            continue
        if _is_learning_content_text(doc.text_at(i)):
            status["学习内容"] = True
            break

    return status


def get_missing_items(status):
    return [item for item in MATERIAL_ORDER if not status[item]]


def get_inserted_items(status):
    return [item for item in MATERIAL_ORDER if status[item]]


def get_or_create_week(path, week_title):
    """获取或创建某周标题索引。"""
    doc = DocxCache(path)
    anchor = doc.find(week_title, exclude_if_contains=DEFAULT_SCHOOL)
    if anchor >= 0:
        return anchor

    keyword = week_title.replace(DEFAULT_SEMESTER, "").strip() or week_title[-20:]
    anchor = doc.find(keyword, exclude_if_contains=DEFAULT_SCHOOL)
    if anchor >= 0:
        return anchor

    # 创建新周
    total = len(doc)
    if total > 0 and doc.text_at(total - 1):
        oc_add_para(path, "", None)
        total += 1

    oc_add_para(path, "会议记录", None)
    oc_add_para(path, "会议签到表", None)
    title_idx = total + 2
    oc_add_para(path, week_title, None)
    oc_set_format(path, title_idx, font="宋体", size="20pt", bold=True)
    return title_idx


def insert_learning_content_text(path, week_keyword, paragraphs):
    """以纯文本方式插入学习内容。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        anchor = get_or_create_week(path, week_keyword)
        doc.invalidate()
        doc._load()
        next_start = len(doc)

    last_pos = doc.last_content_para(anchor + 1, next_start)
    insert_pos = max(last_pos, anchor)
    current_length = len(doc)

    def add_para(text, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_para(path, text, office_index)
        current_length += 1
        return actual_index

    insert_pos = add_para("", insert_pos + 1)

    for p in paragraphs:
        t = p.strip()
        if t:
            insert_pos = add_para(t, insert_pos + 1)
            oc_set_format(path, insert_pos, font=BODY_FONT, size=BODY_SIZE)

    doc.invalidate()
    return detect_week_status(path, week_keyword)


def insert_photo(path, week_keyword, image_path):
    """插入会议照片。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        anchor = get_or_create_week(path, week_keyword)
        doc.invalidate()
        doc._load()
        next_start = len(doc)

    current_length = len(doc)

    def add_para(text, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_para(path, text, office_index)
        current_length += 1
        return actual_index

    def add_image(image, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_image(path, image, office_index)
        current_length += 1
        return actual_index

    photo_marker = doc.find_marker("会议照片：", anchor + 1, next_start)
    if photo_marker < 0:
        brief_idx = doc.find_marker(DEFAULT_SCHOOL, anchor + 1, min(anchor + 10, next_start))
        photo_marker = brief_idx if brief_idx >= 0 else anchor
        photo_marker = add_para("会议照片：", photo_marker + 1)

    add_image(image_path, photo_marker + 1)
    doc.invalidate()
    return detect_week_status(path, week_keyword)


def insert_signin_sheet(path, week_keyword, image_path):
    """插入签到表。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        anchor = get_or_create_week(path, week_keyword)
        doc.invalidate()
        doc._load()

    current_length = len(doc)

    def add_para(text, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_para(path, text, office_index)
        current_length += 1
        return actual_index

    def add_image(image, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_image(path, image, office_index)
        current_length += 1
        return actual_index

    existing = doc.find_marker("会议签到表", max(0, anchor - 20), anchor)
    if existing >= 0:
        add_image(image_path, existing + 1)
        doc.invalidate()
        return detect_week_status(path, week_keyword)

    meeting_record = -1
    for i in range(max(0, anchor - 20), anchor):
        if doc.text_at(i) == "会议记录":
            meeting_record = i
            break

    insert_pos = meeting_record if meeting_record >= 0 else anchor - 1
    if insert_pos < 0:
        insert_pos = add_para("会议记录", 0)

    insert_pos = add_para("会议签到表", insert_pos + 1)
    add_image(image_path, insert_pos + 1)
    doc.invalidate()
    return detect_week_status(path, week_keyword)


def update_brief_intro(path, week_keyword, content_titles):
    """更新简要介绍（四号宋体）。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        return False

    title_text = doc.text_at(anchor)
    week_part = title_text.replace(DEFAULT_SEMESTER, "").strip()
    titles_str = "、".join(content_titles)
    brief_text = f"{DEFAULT_SCHOOL}{DEFAULT_SEMESTER}{week_part}：{titles_str}"

    brief_idx = doc.find_marker(DEFAULT_SCHOOL, anchor + 1, min(anchor + 10, next_start))
    if brief_idx >= 0:
        oc_replace_text(path, brief_idx, brief_text)
        oc_set_format(path, brief_idx, font="宋体", size="14pt")
    else:
        office_index, actual_index = _normalize_add_index(anchor + 1, len(doc))
        oc_add_para(path, brief_text, office_index)
        oc_set_format(path, actual_index, font="宋体", size="14pt")

    return True


def ensure_signin_marker(path, week_keyword):
    """确保某周标题前有“会议签到表”占位。"""
    doc = DocxCache(path)
    anchor, _ = doc.week_boundary(week_keyword)
    if anchor < 0:
        return False

    prev_start = 0
    for idx, _ in doc.find_all_weeks():
        if idx < anchor:
            prev_start = idx

    if doc.find_marker("会议签到表", prev_start, anchor) >= 0:
        return True

    meeting_record = -1
    for i in range(max(0, anchor - 20), anchor):
        if doc.text_at(i) == "会议记录":
            meeting_record = i
            break

    insert_pos = meeting_record + 1 if meeting_record >= 0 else anchor
    office_index, _ = _normalize_add_index(insert_pos, len(doc))
    oc_add_para(path, "会议签到表", office_index)
    return True


def ensure_photo_marker(path, week_keyword):
    """确保某周有“会议照片：”占位，位置紧跟简要介绍之后。"""
    doc = DocxCache(path)
    anchor, next_start = doc.week_boundary(week_keyword)
    if anchor < 0:
        return False

    if doc.find_marker("会议照片：", anchor + 1, next_start) >= 0:
        return True

    brief_idx = doc.find_marker(DEFAULT_SCHOOL, anchor + 1, min(anchor + 10, next_start))
    insert_after = brief_idx if brief_idx >= 0 else anchor
    office_index, _ = _normalize_add_index(insert_after + 1, len(doc))
    oc_add_para(path, "会议照片：", office_index)
    return True


def copy_docx_content(path, week_keyword, source_path):
    """将源 docx 内容段落追加到指定周。"""
    src_doc = DocxCache(source_path)
    dest_doc = DocxCache(path)

    paragraphs = [t.strip() for t in src_doc.texts if t.strip()]

    anchor, next_start = dest_doc.week_boundary(week_keyword)
    if anchor < 0:
        anchor = get_or_create_week(path, week_keyword)
        dest_doc.invalidate()
        dest_doc._load()
        next_start = len(dest_doc)

    last_pos = dest_doc.last_content_para(anchor + 1, next_start)
    insert_pos = max(last_pos, anchor)
    current_length = len(dest_doc)

    def add_para(text, preferred_index):
        nonlocal current_length
        office_index, actual_index = _normalize_add_index(preferred_index, current_length)
        oc_add_para(path, text, office_index)
        current_length += 1
        return actual_index

    insert_pos = add_para("", insert_pos + 1)
    for p in paragraphs:
        insert_pos = add_para(p, insert_pos + 1)
