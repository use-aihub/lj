import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "review_artifact_guard.py"
SPEC = importlib.util.spec_from_file_location("review_artifact_guard", SCRIPT)
guard = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(guard)


class ReviewArtifactGuardTests(unittest.TestCase):
    def state(self, formats=("docx",)):
        return {
            "version": guard.STATE_VERSION,
            "status": "prepared",
            "files": [
                {
                    "file_id": "f%03d" % index,
                    "format": file_format,
                    "baseline_comments": [],
                }
                for index, file_format in enumerate(formats, start=1)
            ],
        }

    def issue(self, issue_id="f001-i001", **overrides):
        value = {
            "issue_id": issue_id,
            "kind": "point",
            "severity": "一般",
            "location": "第2章，表2-1，第1行容量",
            "original_kind": "value",
            "original": "50 MVA",
            "reason": "设计批复第3页明确为63 MVA",
            "correct": "63 MVA",
            "action": "直接将本处50 MVA改为63 MVA",
            "anchor": {
                "format": "docx",
                "paragraph_path": "/body/p[@paraId=ABCDEF01]",
                "anchor_text": "50 MVA",
            },
        }
        value.update(overrides)
        return value

    def payload(self, *issues):
        return {"files": [{"file_id": "f001", "issues": list(issues)}]}

    def test_incrementing_name_and_paths_with_spaces(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary) / "含 空格"
            directory.mkdir()
            source = directory / "项目报告.DOCX"
            source.touch()
            first = directory / "项目报告-审核批注版.DOCX"
            first.touch()
            second = guard.next_output_path(source)
            self.assertEqual(directory.resolve() / "项目报告-审核批注版-v2.DOCX", second)
            occupied = {second}
            self.assertEqual(
                directory.resolve() / "项目报告-审核批注版-v3.DOCX",
                guard.next_output_path(source, occupied=occupied),
            )

    def test_required_fields_forbidden_phrases_and_duplicates(self):
        missing = self.issue()
        del missing["action"]
        with self.assertRaisesRegex(guard.ReviewGuardError, "action.*不能为空"):
            guard.validate_issues(self.state(), self.payload(missing))

        forbidden = self.issue(action="请核对设计批复后修改本处")
        with self.assertRaisesRegex(guard.ReviewGuardError, "命中禁用措辞"):
            guard.validate_issues(self.state(), self.payload(forbidden))

        quoted_original = self.issue(original="原报告写有‘请核对后填写’")
        self.assertEqual(
            1,
            len(guard.validate_issues(self.state(), self.payload(quoted_original))["f001"]),
        )

        first = self.issue()
        second = self.issue(issue_id="f001-i002")
        with self.assertRaisesRegex(guard.ReviewGuardError, "批注全文重复"):
            guard.validate_issues(self.state(), self.payload(first, second))

    def test_unit_superscript_gate(self):
        rejected = (
            "100m2",
            "100 m3",
            "0.25km2",
            "1 km3",
            "0.0750hm2",
            "10dm3",
            "2dam2",
            "10cm3",
            "20mm2",
            "25nm2",
            "1um3",
            "1e3 m2",
            "单位：m2",
            "单位（m2）",
            "面积/m2",
            "排放强度单位为kg/m2",
            "面积单位采用m2",
            "按m^3计",
            "m^2/m^3",
            "将hm2改成hm²",
            "10-20m2",
        )
        for text in rejected:
            with self.subTest(normalized=text):
                self.assertEqual(
                    [],
                    guard.ascii_unit_power_hits(guard.normalize_unit_superscripts(text)),
                )

        allowed = (
            "100 m²",
            "1 km³",
            "0.0750 hm²",
            "HJ 2.3-2018",
            "GB/T 12345m2-2020",
            "GB/T 12345m2.1-2020",
            "标准号为ISO m2:2020",
            "标准号为DB33/T m2-2020",
            "变量m2、km3_value与M2",
            "将变量值改为m2",
            "参数名称为m2",
            "变量（m2）",
            "字段名m2",
            "公式中m2表示系数",
            "变量`m2`",
            "路径/tmp/m2/report",
            "路径为m2/report",
            "/data/kg/m2/report",
            "/data/m2/m3/report",
            "路径为100m2.docx",
            "文件report-10m2.docx",
            "m2.pptx",
            "m2.md",
            "m2.zip",
            "m2.tar.gz",
        )
        for text in allowed:
            with self.subTest(allowed=text):
                self.assertEqual([], guard.ascii_unit_power_hits(text))

        bad_issue = self.issue(reason="批复面积为100m2，本处抄录不一致")
        normalized_bad = guard.validate_issues(self.state(), self.payload(bad_issue))["f001"]
        self.assertIn("100m²", normalized_bad[0]["comment"])

        good_issue = self.issue(reason="批复面积为100 m²，本处抄录不一致")
        self.assertEqual(1, len(guard.validate_issues(self.state(), self.payload(good_issue))["f001"]))

        raw_original = self.issue(
            original="100m2",
            anchor={"format": "docx", "paragraph_path": "/body/p[1]", "anchor_text": "100m2"},
        )
        normalized_raw = guard.validate_issues(self.state(), self.payload(raw_original))["f001"]
        self.assertIn("原值“100m²”", normalized_raw[0]["comment"])
        normalized_original = self.issue(
            original="100m²",
            anchor={"format": "docx", "paragraph_path": "/body/p[1]", "anchor_text": "100m2"},
        )
        self.assertEqual(
            1,
            len(guard.validate_issues(self.state(), self.payload(normalized_original))["f001"]),
        )

        literal_issue = self.issue(correct="变量m2")
        self.assertEqual(1, len(guard.validate_issues(self.state(), self.payload(literal_issue))["f001"]))

    def test_readback_rejects_nfkc_unit_degradation(self):
        self.assertEqual(guard.normalize_text("100 m²"), guard.normalize_text("100 m2"))
        state = self.state(("xlsx",))
        issue = self.issue(
            original="100 m²",
            reason="批复面积为120 m²，本处抄录不一致",
            correct="120 m²",
            action="直接将本单元格100 m²改为120 m²",
            anchor={"format": "xlsx", "sheet": "验收", "cell": "A1"},
        )
        expected = guard.validate_issues(state, self.payload(issue))["f001"]
        readback = {
            "text": expected[0]["comment"].replace("m²", "m2"),
            "path": "/验收/A1",
            "format": {
                "author": guard.AUTHOR,
                "anchoredTo": "/验收/A1",
                "ref": "A1",
                "id": "1",
            },
        }
        item = {
            "stage": "/tmp/unit-readback.xlsx",
            "source": "/tmp/source.xlsx",
            "format": "xlsx",
            "baseline_comment_count": 0,
            "baseline_comments": [],
        }
        with mock.patch.object(guard, "query_comments", return_value=[readback]):
            with self.assertRaisesRegex(guard.ReviewGuardError, "单位角标字形与写入前不一致"):
                guard.verify_comments_for_file("officecli", item, expected)

        readback["text"] = expected[0]["comment"]
        with mock.patch.object(guard, "query_comments", return_value=[readback]):
            self.assertEqual(
                {"total": 1, "original": 0, "new": 1},
                guard.verify_comments_for_file("officecli", item, expected),
            )

    def test_section_summary_is_explicitly_scoped(self):
        summary = self.issue(
            kind="section_summary",
            location="8.2附件",
            original_kind="text",
            original="附件目录列有8项材料，实际未提供实体",
            reason="许可、质量评定和土方去向均无法由附件验证",
            correct="附件目录、实体和正文引用逐项闭合",
            action="补入附件2、附件4和附件7实体，并同步修正对应正文引用",
            section_id="8.2附件",
            affected_items=["附件2", "附件4", "附件7"],
        )
        self.assertEqual(1, len(guard.validate_issues(self.state(), self.payload(summary))["f001"]))

        del summary["affected_items"]
        with self.assertRaisesRegex(guard.ReviewGuardError, "affected_items必须是非空数组"):
            guard.validate_issues(self.state(), self.payload(summary))

    def test_table_summary_closes_all_numeric_rows(self):
        point_one = self.issue(
            table_id="表4-2",
            row_label="第3行",
            anchor={"format": "docx", "paragraph_path": "/body/p[1]", "anchor_text": "50.0"},
        )
        point_two = self.issue(
            issue_id="f001-i002",
            location="表4-2，第4行，夜间叠加值",
            original="49.0 dB(A)",
            reason="背景值47.0 dB(A)、贡献值48.0 dB(A)，能量叠加结果为50.6 dB(A)",
            correct="50.6 dB(A)",
            action="直接将本单元格49.0改为50.6",
            table_id="表4-2",
            row_label="第4行",
            anchor={"format": "docx", "paragraph_path": "/body/p[2]", "anchor_text": "49.0"},
        )
        summary = self.issue(
            issue_id="f001-i003",
            kind="table_summary",
            location="表4-2，数值错误汇总",
            original_kind="text",
            original="第3行和第4行夜间叠加值",
            reason="上述两行未按噪声能量公式形成叠加值",
            correct="第3行改为51.6 dB(A)，第4行改为50.6 dB(A)",
            action="依次修改第3行和第4行；只修改上述所列行，未列行无需复核",
            table_id="表4-2",
            affected_rows=["第3行", "第4行"],
            anchor={"format": "docx", "paragraph_path": "/body/p[3]", "anchor_text": "表4-2"},
        )
        self.assertEqual(
            2,
            len(guard.validate_issues(self.state(), self.payload(point_one, point_two))["f001"]),
        )
        validated = guard.validate_issues(self.state(), self.payload(point_one, point_two, summary))
        self.assertEqual(3, len(validated["f001"]))
        self.assertEqual([], guard.forbidden_hits(guard.SUMMARY_ACTION))

        summary["affected_rows"] = ["第3行"]
        with self.assertRaisesRegex(guard.ReviewGuardError, "汇总行.*不一致"):
            guard.validate_issues(self.state(), self.payload(point_one, point_two, summary))

    def test_docx_comment_anchor_text_must_be_nonempty(self):
        document_xml = """\
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>非空锚点文字</w:t></w:r>
  <w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p></w:body>
</w:document>"""
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "anchor.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", document_xml)
            self.assertEqual("非空锚点文字", guard.docx_anchor_texts(path)["7"])


if __name__ == "__main__":
    unittest.main()
