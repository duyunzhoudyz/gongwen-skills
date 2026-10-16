# -*- coding: utf-8 -*-
"""gongwen-profile 脚本测试的公共工具：临时目录、调用 main()、用标准库生成带修订痕迹的 docx。"""
from __future__ import annotations

import io
import shutil
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = REPO_ROOT / "skills" / "gongwen-profile" / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '{comments}'
    '</Types>'
)
COMMENTS_OVERRIDE = ('<Override PartName="/word/comments.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>')
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '{comments}'
    '</Relationships>'
)
COMMENTS_REL = ('<Relationship Id="rId9" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
                'Target="comments.xml"/>')

# 段落内的片段：
#   ("t", 文字)                       普通文字
#   ("ins", 文字, 修订人, 日期)        插入
#   ("del", 文字, 修订人, 日期)        删除
#   ("moveFrom", 文字, 修订人, 日期)   移出
#   ("moveTo", 文字, 修订人, 日期)     移入
#   ("cstart", 批注编号) / ("cend", 批注编号)
#   ("insdel", 文字, 插入人, 删除人, 日期)  嵌套修订：插入人加上、删除人又删掉
Run = Tuple


def _run(text: str, deleted: bool = False) -> str:
    tag = "w:delText" if deleted else "w:t"
    return f'<w:r><{tag} xml:space="preserve">{escape(text)}</{tag}></w:r>'


def paragraph_xml(runs: Sequence[Run], rev_id: List[int]) -> str:
    parts = []
    for item in runs:
        kind = item[0]
        if kind == "t":
            parts.append(_run(item[1]))
        elif kind in ("ins", "del", "moveFrom", "moveTo"):
            rev_id[0] += 1
            text, author, date = item[1], item[2], item[3]
            inner = _run(text, deleted=(kind == "del"))
            parts.append(f'<w:{kind} w:id="{rev_id[0]}" w:author="{escape(author)}" w:date="{date}">{inner}</w:{kind}>')
        elif kind == "insdel":
            text, ins_author, del_author, date = item[1], item[2], item[3], item[4]
            rev_id[0] += 2
            inner = _run(text, deleted=True)
            parts.append(f'<w:ins w:id="{rev_id[0] - 1}" w:author="{escape(ins_author)}" w:date="{date}">'
                         f'<w:del w:id="{rev_id[0]}" w:author="{escape(del_author)}" w:date="{date}">{inner}</w:del>'
                         f'</w:ins>')
        elif kind == "cstart":
            parts.append(f'<w:commentRangeStart w:id="{item[1]}"/>')
        elif kind == "cend":
            parts.append(f'<w:commentRangeEnd w:id="{item[1]}"/>'
                         f'<w:r><w:commentReference w:id="{item[1]}"/></w:r>')
        else:
            raise ValueError(kind)
    return "<w:p>" + "".join(parts) + "</w:p>"


def make_docx(path: Path, paragraphs: Sequence[Sequence[Run]],
              comments: Optional[Sequence[Tuple[str, str, str, str]]] = None,
              table_rows: Optional[Sequence[Sequence[str]]] = None,
              header: Optional[str] = None, footer: Optional[str] = None) -> Path:
    """生成最小可用的 docx。comments 为 [(编号, 作者, 日期, 内容)]；header、footer 为页眉、页脚文字。"""
    rev_id = [0]
    body = "".join(paragraph_xml(p, rev_id) for p in paragraphs)
    if table_rows:
        rows = "".join("<w:tr>" + "".join(f"<w:tc><w:p>{_run(c)}</w:p></w:tc>" for c in row) + "</w:tr>"
                       for row in table_rows)
        body += f"<w:tbl>{rows}</w:tbl>"
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>')
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES.format(comments=COMMENTS_OVERRIDE if comments else ""))
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS.format(comments=COMMENTS_REL if comments else ""))
        zf.writestr("word/document.xml", document)
        for part, text in (("header1", header), ("footer1", footer)):
            if text is not None:
                tag = "hdr" if part.startswith("header") else "ftr"
                zf.writestr(f"word/{part}.xml",
                            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                            f'<w:{tag} xmlns:w="{W_NS}"><w:p>{_run(text)}</w:p></w:{tag}>')
        if comments:
            items = "".join(
                f'<w:comment w:id="{cid}" w:author="{escape(author)}" w:date="{date}">'
                f'<w:p>{_run(text)}</w:p></w:comment>'
                for cid, author, date, text in comments)
            zf.writestr("word/comments.xml",
                        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w:comments xmlns:w="{W_NS}">{items}</w:comments>')
    return path


class ScriptTestCase(unittest.TestCase):
    """带临时目录的测试基类。子类设置 module 为被测脚本模块。"""

    module = None

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="gongwen_profile_test_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write(self, name: str, text: str, encoding: str = "utf-8") -> Path:
        path = self.tmp / name
        path.write_bytes(text.encode(encoding))
        return path

    def run_main(self, *args: str):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = self.module.main(list(args))  # type: ignore[union-attr]
        return code, out.getvalue(), err.getvalue()
