#!/usr/bin/env python3
"""安全准备、批注、验收并发布 DOCX/XLSX 审核副本。"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET


STATE_VERSION = 1
AUTHOR = "审核助手"
DEFAULT_PROFILE = "fast"
PROFILES = ("fast", "balanced", "strict")
PROFILE_SNAPSHOT_MODES = {
    "fast": ("text",),
    "balanced": ("text", "outline"),
    "strict": ("text", "annotated", "outline"),
}
SEVERITIES = ("重大", "较大", "一般", "文字")
ORIGINAL_KINDS = ("value", "text")
ISSUE_KINDS = ("point", "table_summary", "section_summary")
SUPPORTED_EXTENSIONS = (".docx", ".xlsx")
REVIEWED_STEM_RE = re.compile(r"-审核批注版(?:-v\d+)?$", re.IGNORECASE)
CELL_RE = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")
TRANSFER_RE = re.compile(r"(?:请|建议|(?<!无)需)(?:进一步)?(?:核对|复核|核实|确认)")
FORBIDDEN_PHRASES = (
    "逐项重算",
    "建议完善",
    "可能有误",
    "全面检查",
    "全面复核",
    "自行确认",
    "确认后修改",
    "核实后修改",
)
SUMMARY_ACTION = "只修改上述所列行，未列行无需复核"
DROP_SNAPSHOT_KEYS = {"filename", "filepath"}
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_ID = "{%s}id" % WORD_NS
UNIT_PREFIX_PATTERN = r"(?:km|hm|dam|dm|cm|mm|μm|µm|um|nm|m)"
NUMBER_PATTERN = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?"
ASCII_UNIT_POWER_RE = re.compile(
    r"(?<![A-Za-z0-9_.\\+\-])"
    r"(?:(?P<quantity>" + NUMBER_PATTERN
    + r"(?:\s*(?:-|~|～|至|–|—)\s*" + NUMBER_PATTERN + r")?\s*))?"
    r"(?P<unit>" + UNIT_PREFIX_PATTERN + r")(?P<caret>\^?)(?P<power>[23])"
    r"(?![A-Za-z0-9_])"
)
UNIT_POWER_SPELLING_RE = re.compile(
    r"(?<![A-Za-z_])" + UNIT_PREFIX_PATTERN + r"(?:\^?[23]|[²³])(?![A-Za-z0-9_])"
)
UNIT_POWER_LITERAL_RE = re.compile(r"^" + UNIT_PREFIX_PATTERN + r"(?:\^?[23]|[²³])$")
UNIT_POWER_SUPERSCRIPTS = {"2": "²", "3": "³"}
FILE_SUFFIX_RE = re.compile(r"^(?:\.[A-Za-z][A-Za-z0-9_-]*)+")
STANDARD_PREFIX_RE = re.compile(r"(?:GB(?:/T)?|HJ|SL|DL/T|NB/T|Q/[A-Z0-9.-]+)\s*$", re.IGNORECASE)
STANDARD_SUFFIX_RE = re.compile(r"^(?:\.\d+(?:\.\d+)*)?\s*[-–—]\s*\d{4}\b")
VARIABLE_CONTEXT_RE = re.compile(
    r"(?:变量|参数|系数|代号|字段|标识符)(?:名称|名|值)?(?:应为|改为|写为|是|为)?\s*[（(]?\s*$"
)
VARIABLE_SUFFIX_RE = re.compile(
    r"^\s*[）)]?\s*(?:表示|代表|作为|是|为)?\s*(?:变量|参数|系数|代号|字段|标识符)(?:名称|名|值)?"
)
PATH_CONTEXT_RE = re.compile(r"(?:路径|文件(?:名)?|目录(?:名)?|URL|链接)(?:为|是|：|:)?\s*$", re.IGNORECASE)
STANDARD_CONTEXT_RE = re.compile(
    r"(?:标准|规范|导则)(?:号|编号|代码|名称)?(?:为|是|采用|引用)?[A-Za-z0-9/ ._-]*$",
    re.IGNORECASE,
)
RATE_NUMERATOR_RE = re.compile(
    r"^(?:t|kg|g|mg|μg|ug|l|ml|元|万元|人|株|次|面积|体积|容积)$",
    re.IGNORECASE,
)
RATE_DENOMINATOR_RE = re.compile(r"^(?:s|min|h|d|a|年|月|次|人|台)$", re.IGNORECASE)


class ReviewGuardError(RuntimeError):
    pass


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    return re.sub(r"\s+", " ", value).strip()


def unit_power_literal_context(text: str, start: int, end: int) -> bool:
    """判断候选串是否明确属于变量、路径、文件名或标准号。"""
    prefix = text[:start]
    suffix = text[end:]
    if prefix.count("`") % 2:
        return True
    if FILE_SUFFIX_RE.match(suffix):
        return True
    if STANDARD_PREFIX_RE.search(prefix[-48:]) and STANDARD_SUFFIX_RE.match(suffix):
        return True
    if STANDARD_CONTEXT_RE.search(prefix[-64:]):
        return True
    if PATH_CONTEXT_RE.search(prefix[-48:]):
        return True
    if VARIABLE_CONTEXT_RE.search(prefix[-32:]):
        return True
    if VARIABLE_SUFFIX_RE.match(suffix):
        return True
    slash_suffix = re.match(r"^/\s*([^/\s，。；;）)]+)", suffix)
    if slash_suffix:
        denominator = slash_suffix.group(1)
        if not RATE_DENOMINATOR_RE.fullmatch(denominator) and not UNIT_POWER_LITERAL_RE.fullmatch(denominator):
            return True
    slash_context = re.search(r"([^\s，。；;（）()]*)/\s*$", prefix)
    if slash_context:
        segments = [segment for segment in slash_context.group(1).split("/") if segment]
        for segment in segments:
            rate_token = re.search(r"([A-Za-zμµ]+|万元|面积|体积|容积|元|人|株|次)$", segment)
            is_rate = bool(rate_token and RATE_NUMERATOR_RE.fullmatch(rate_token.group(1)))
            if not is_rate and not UNIT_POWER_LITERAL_RE.fullmatch(segment):
                return True
    return False


def ascii_unit_power_hits(text: str) -> List[str]:
    """返回非字面量语境中的 ASCII 单位角标及规范写法。"""
    hits: List[str] = []
    seen: set[str] = set()
    source = text or ""
    for match in ASCII_UNIT_POWER_RE.finditer(source):
        if unit_power_literal_context(source, match.start(), match.end()):
            continue
        unit = match.group("unit")
        power = match.group("power")
        spelling = "%s%s%s" % (unit, match.group("caret"), power)
        hit = "%s→%s%s" % (spelling, unit, UNIT_POWER_SUPERSCRIPTS[power])
        if hit not in seen:
            seen.add(hit)
            hits.append(hit)
    return hits


def normalize_unit_superscripts(text: str) -> str:
    """规范明确表示面积、体积单位的 ASCII 角标，保留变量、路径和标准号。"""
    source = text or ""
    parts: List[str] = []
    cursor = 0
    for match in ASCII_UNIT_POWER_RE.finditer(source):
        if unit_power_literal_context(source, match.start(), match.end()):
            continue
        parts.append(source[cursor:match.start()])
        parts.append(
            "%s%s%s" % (
                match.group("quantity") or "",
                match.group("unit"),
                UNIT_POWER_SUPERSCRIPTS[match.group("power")],
            )
        )
        cursor = match.end()
    if not parts:
        return source
    parts.append(source[cursor:])
    return "".join(parts)


def unit_power_spellings(text: str) -> List[str]:
    """保留原始字形，避免 NFKC 将 m²/m³ 与 m2/m3 视为相同。"""
    return [match.group(0) for match in UNIT_POWER_SPELLING_RE.finditer(text or "")]


def require_nonempty(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ReviewGuardError("%s不能为空" % label)
    return value.strip()


def json_dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def atomic_write_json(path: Path, payload: Any, refuse_existing: bool = False) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    if refuse_existing and path.exists():
        raise ReviewGuardError("状态文件已存在，拒绝覆盖：%s" % path)
    fd, temporary_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, path)
    finally:
        if os.path.exists(temporary_name):
            os.unlink(temporary_name)


def load_json(path: Path, label: str) -> Dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            value = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise ReviewGuardError("无法读取%s %s：%s" % (label, path, exc)) from exc
    if not isinstance(value, dict):
        raise ReviewGuardError("%s顶层必须是JSON对象：%s" % (label, path))
    return value


def resolve_officecli(explicit: Optional[str]) -> str:
    if explicit:
        candidate = shutil.which(explicit) or (explicit if Path(explicit).is_file() else None)
    else:
        candidate = shutil.which("officecli")
    if not candidate:
        raise ReviewGuardError("未找到officecli；先安装并确认其位于PATH")
    return str(candidate)


def parse_json_output(stdout: str, stderr: str, command: Sequence[str]) -> Dict[str, Any]:
    for text in (stdout, stderr):
        stripped = text.strip().lstrip("\ufeff")
        if not stripped:
            continue
        try:
            value = json.loads(stripped)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    detail = (stderr or stdout).strip()
    raise ReviewGuardError("OfficeCLI未返回有效JSON：%s；%s" % (" ".join(command), detail[:500]))


def officecli_json(
    officecli: str,
    args: Sequence[str],
    allow_failure: bool = False,
) -> Tuple[Dict[str, Any], subprocess.CompletedProcess[str]]:
    command = [officecli] + [str(arg) for arg in args]
    env = os.environ.copy()
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    process = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        check=False,
    )
    payload = parse_json_output(process.stdout, process.stderr, command)
    if not allow_failure:
        if process.returncode != 0 or payload.get("success") is False:
            error = payload.get("error") or process.stderr.strip() or process.stdout.strip()
            raise ReviewGuardError("OfficeCLI命令失败：%s；%s" % (" ".join(command), error))
    return payload, process


def batch_items(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    containers: List[Dict[str, Any]] = [payload]
    if isinstance(payload.get("data"), dict):
        containers.insert(0, payload["data"])
    for container in containers:
        for key in ("results", "Results", "items", "Items", "commands", "Commands"):
            value = container.get(key)
            if isinstance(value, list) and all(isinstance(item, dict) for item in value):
                return value
    return []


def batch_item_results(item: Dict[str, Any]) -> List[Dict[str, Any]]:
    candidates = [item]
    for key in ("result", "Result", "output", "Output", "response", "Response"):
        nested = item.get(key)
        if isinstance(nested, dict):
            candidates.append(nested)
    for candidate in candidates:
        results = data_results(candidate)
        if results:
            return results
        for key in ("results", "Results"):
            value = candidate.get(key)
            if isinstance(value, list):
                return [result for result in value if isinstance(result, dict)]
    if any(key in item for key in ("path", "text", "children")):
        return [item]
    return []


def officecli_batch(
    officecli: str,
    document: Path,
    commands: List[Dict[str, Any]],
    command_file: Path,
    require_results: bool = True,
) -> List[Dict[str, Any]]:
    if not commands:
        return []
    atomic_write_json(command_file, commands)
    payload, _ = officecli_json(
        officecli,
        ["batch", str(document), "--input", str(command_file), "--stop-on-error", "--json"],
    )
    items = batch_items(payload)
    if require_results and len(items) != len(commands):
        raise ReviewGuardError(
            "OfficeCLI批处理结果数不闭合：预期%d，实际%d" % (len(commands), len(items))
        )
    for index, item in enumerate(items, start=1):
        success = item.get("success", item.get("Success"))
        if success is False:
            error = item.get("error") or item.get("Error") or "未知错误"
            raise ReviewGuardError("OfficeCLI批处理第%d项失败：%s" % (index, error))
    return items


def data_results(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    data = payload.get("data")
    if not isinstance(data, dict):
        return []
    results = data.get("results")
    if results is None:
        results = data.get("Results")
    if not isinstance(results, list):
        return []
    return [item for item in results if isinstance(item, dict)]


def query_comments(officecli: str, path: Path) -> List[Dict[str, Any]]:
    payload, _ = officecli_json(officecli, ["query", str(path), "comment", "--json"])
    return data_results(payload)


def comment_record(result: Dict[str, Any]) -> Dict[str, str]:
    formatting = result.get("format") if isinstance(result.get("format"), dict) else {}
    return {
        "text": str(result.get("text") or ""),
        "author": str(formatting.get("author") or ""),
        "initials": str(formatting.get("initials") or ""),
        "anchored_to": str(formatting.get("anchoredTo") or ""),
        "ref": str(formatting.get("ref") or ""),
        "id": str(formatting.get("id") or ""),
        "path": str(result.get("path") or ""),
    }


def canonical_comments(results: Iterable[Dict[str, Any]]) -> List[Dict[str, str]]:
    records = [comment_record(item) for item in results]
    return sorted(records, key=json_dump)


def canonicalize_snapshot(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: canonicalize_snapshot(item)
            for key, item in value.items()
            if key.lower() not in DROP_SNAPSHOT_KEYS
        }
    if isinstance(value, list):
        return [canonicalize_snapshot(item) for item in value]
    return value


def snapshot_hash(officecli: str, path: Path, mode: str) -> str:
    payload, _ = officecli_json(officecli, ["view", str(path), mode, "--json"])
    canonical = canonicalize_snapshot(payload.get("data"))
    if mode == "outline":
        canonical = scrub_error_value(canonical, path)
    return hashlib.sha256(json_dump(canonical).encode("utf-8")).hexdigest()


def scrub_error_value(value: Any, file_path: Path) -> Any:
    if isinstance(value, dict):
        return {
            key: scrub_error_value(item, file_path)
            for key, item in value.items()
            if key.lower() not in DROP_SNAPSHOT_KEYS
        }
    if isinstance(value, list):
        return [scrub_error_value(item, file_path) for item in value]
    if isinstance(value, str):
        return value.replace(str(file_path), "<FILE>").replace(file_path.name, "<FILE>")
    return value


def validate_document(officecli: str, path: Path) -> Dict[str, Any]:
    payload, _ = officecli_json(officecli, ["validate", str(path), "--json"], allow_failure=True)
    if payload.get("success") is False:
        raise ReviewGuardError("OfficeCLI结构检查失败：%s；%s" % (path, payload.get("error")))
    data = payload.get("data") if isinstance(payload.get("data"), dict) else {}
    errors = data.get("errors") if isinstance(data.get("errors"), list) else []
    signatures = Counter(json_dump(scrub_error_value(error, path)) for error in errors)
    return {
        "count": int(data.get("count", len(errors))),
        "signatures": [
            {"signature": signature, "count": count}
            for signature, count in sorted(signatures.items())
        ],
    }


def signature_counter(validation: Dict[str, Any]) -> Counter[str]:
    counter: Counter[str] = Counter()
    for item in validation.get("signatures", []):
        if isinstance(item, dict):
            counter[str(item.get("signature", ""))] += int(item.get("count", 0))
    return counter


def find_page_count(value: Any) -> Optional[int]:
    if isinstance(value, dict):
        for key, item in value.items():
            if key.lower() in {"pages", "pagecount", "totalpages"} and isinstance(item, int):
                return item
        for item in value.values():
            found = find_page_count(item)
            if found is not None:
                return found
    elif isinstance(value, list):
        for item in value:
            found = find_page_count(item)
            if found is not None:
                return found
    return None


def page_count_status(officecli: str, path: Path, file_format: str) -> Dict[str, Any]:
    if file_format != "docx":
        return {"status": "not_applicable", "value": None}
    payload, process = officecli_json(
        officecli,
        ["view", str(path), "stats", "--page-count", "--json"],
        allow_failure=True,
    )
    if process.returncode != 0 or payload.get("success") is False:
        error = payload.get("error")
        if isinstance(error, dict):
            error = error.get("error") or error.get("code")
        return {"status": "not_performed", "value": None, "reason": str(error or "页数接口不可用")}
    count = find_page_count(payload.get("data"))
    if count is None:
        return {"status": "not_performed", "value": None, "reason": "OfficeCLI未返回真实页数"}
    return {"status": "available", "value": count}


def is_reviewed_output(path: Path) -> bool:
    return bool(REVIEWED_STEM_RE.search(path.stem)) or path.name.startswith("~$")


def next_output_path(source: Path, output_dir: Optional[Path] = None, occupied: Optional[set[Path]] = None) -> Path:
    directory = (output_dir or source.parent).resolve()
    occupied = occupied or set()
    base = "%s-审核批注版" % source.stem
    candidate = directory / (base + source.suffix)
    version = 2
    while candidate.exists() or candidate in occupied:
        candidate = directory / ("%s-v%d%s" % (base, version, source.suffix))
        version += 1
    return candidate


def state_file_map(state: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    files = state.get("files")
    if not isinstance(files, list):
        raise ReviewGuardError("状态文件缺少files数组")
    result: Dict[str, Dict[str, Any]] = {}
    for item in files:
        if not isinstance(item, dict):
            raise ReviewGuardError("状态文件files成员必须是对象")
        file_id = require_nonempty(item.get("file_id"), "状态文件file_id")
        if file_id in result:
            raise ReviewGuardError("状态文件file_id重复：%s" % file_id)
        result[file_id] = item
    return result


def state_profile(state: Dict[str, Any]) -> str:
    # 旧状态没有 profile，按旧版完整验收语义处理。
    profile = str(state.get("profile") or "strict").lower()
    if profile not in PROFILES:
        raise ReviewGuardError("不支持的验收档位：%s" % profile)
    return profile


def check_state(state: Dict[str, Any], expected_status: Optional[str] = None) -> None:
    if state.get("version") != STATE_VERSION:
        raise ReviewGuardError("不支持的状态文件版本：%s" % state.get("version"))
    if expected_status and state.get("status") != expected_status:
        raise ReviewGuardError("状态应为%s，实际为%s" % (expected_status, state.get("status")))
    state_profile(state)
    state_file_map(state)


def trim_terminal_punctuation(text: str) -> str:
    return text.rstrip().rstrip("。；;！!")


def compose_comment(issue: Dict[str, Any]) -> str:
    label = "原值" if issue["original_kind"] == "value" else "原文"
    return "【%s｜%s】%s“%s”；%s。正确值/判断应为%s。修改动作：%s。" % (
        issue["severity"],
        trim_terminal_punctuation(issue["location"]),
        label,
        issue["original"],
        trim_terminal_punctuation(issue["reason"]),
        trim_terminal_punctuation(issue["correct"]),
        trim_terminal_punctuation(issue["action"]),
    )


def forbidden_hits(text: str) -> List[str]:
    hits = list(FORBIDDEN_PHRASES[i] for i in range(len(FORBIDDEN_PHRASES)) if FORBIDDEN_PHRASES[i] in text)
    hits.extend(match.group(0) for match in TRANSFER_RE.finditer(text))
    return sorted(set(hits))


def validate_anchor(anchor: Any, file_format: str, issue_id: str) -> Dict[str, str]:
    if not isinstance(anchor, dict):
        raise ReviewGuardError("%s.anchor必须是对象" % issue_id)
    anchor_format = require_nonempty(anchor.get("format"), "%s.anchor.format" % issue_id).lower()
    if anchor_format != file_format:
        raise ReviewGuardError("%s锚点格式%s与文件格式%s不一致" % (issue_id, anchor_format, file_format))
    if file_format == "docx":
        return {
            "format": "docx",
            "paragraph_path": require_nonempty(anchor.get("paragraph_path"), "%s.anchor.paragraph_path" % issue_id),
            "anchor_text": require_nonempty(anchor.get("anchor_text"), "%s.anchor.anchor_text" % issue_id),
        }
    sheet = require_nonempty(anchor.get("sheet"), "%s.anchor.sheet" % issue_id)
    cell = require_nonempty(anchor.get("cell"), "%s.anchor.cell" % issue_id).upper()
    if not CELL_RE.fullmatch(cell):
        raise ReviewGuardError("%s.anchor.cell不是有效单元格：%s" % (issue_id, cell))
    return {"format": "xlsx", "sheet": sheet, "cell": cell}


def validate_issues(state: Dict[str, Any], issues_payload: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    state_files = state_file_map(state)
    entries = issues_payload.get("files")
    if not isinstance(entries, list):
        raise ReviewGuardError("问题清单顶层必须包含files数组")
    seen_files: set[str] = set()
    seen_issue_ids: set[str] = set()
    validated: Dict[str, List[Dict[str, Any]]] = {}

    for entry in entries:
        if not isinstance(entry, dict):
            raise ReviewGuardError("问题清单files成员必须是对象")
        file_id = require_nonempty(entry.get("file_id"), "问题清单file_id")
        if file_id not in state_files:
            raise ReviewGuardError("问题清单包含未知file_id：%s" % file_id)
        if file_id in seen_files:
            raise ReviewGuardError("问题清单file_id重复：%s" % file_id)
        seen_files.add(file_id)
        issues = entry.get("issues")
        if not isinstance(issues, list):
            raise ReviewGuardError("%s.issues必须是数组" % file_id)
        file_format = state_files[file_id]["format"]
        baseline_texts = {
            normalize_text(comment.get("text", ""))
            for comment in state_files[file_id].get("baseline_comments", [])
            if normalize_text(comment.get("text", ""))
        }
        seen_comments: Dict[str, str] = {}
        validated[file_id] = []

        for raw_issue in issues:
            if not isinstance(raw_issue, dict):
                raise ReviewGuardError("%s的issue必须是对象" % file_id)
            issue_id = require_nonempty(raw_issue.get("issue_id"), "%s.issue_id" % file_id)
            if issue_id in seen_issue_ids:
                raise ReviewGuardError("issue_id重复：%s" % issue_id)
            seen_issue_ids.add(issue_id)
            issue = dict(raw_issue)
            issue["kind"] = require_nonempty(issue.get("kind"), "%s.kind" % issue_id)
            issue["severity"] = require_nonempty(issue.get("severity"), "%s.severity" % issue_id)
            issue["location"] = require_nonempty(issue.get("location"), "%s.location" % issue_id)
            issue["original_kind"] = require_nonempty(issue.get("original_kind"), "%s.original_kind" % issue_id)
            issue["original"] = require_nonempty(issue.get("original"), "%s.original" % issue_id)
            issue["reason"] = require_nonempty(issue.get("reason"), "%s.reason" % issue_id)
            issue["correct"] = require_nonempty(issue.get("correct"), "%s.correct" % issue_id)
            issue["action"] = require_nonempty(issue.get("action"), "%s.action" % issue_id)
            if issue["kind"] not in ISSUE_KINDS:
                raise ReviewGuardError("%s.kind仅允许%s" % (issue_id, "/".join(ISSUE_KINDS)))
            if issue["severity"] not in SEVERITIES:
                raise ReviewGuardError("%s.severity仅允许%s" % (issue_id, "/".join(SEVERITIES)))
            if issue["original_kind"] not in ORIGINAL_KINDS:
                raise ReviewGuardError("%s.original_kind仅允许value/text" % issue_id)
            if issue["kind"] == "section_summary":
                if issue["original_kind"] != "text":
                    raise ReviewGuardError("%s.section_summary的original_kind必须为text" % issue_id)
                issue["section_id"] = require_nonempty(issue.get("section_id"), "%s.section_id" % issue_id)
                affected_items = issue.get("affected_items")
                if not isinstance(affected_items, list) or not affected_items:
                    raise ReviewGuardError("%s.affected_items必须是非空数组" % issue_id)
                issue["affected_items"] = [
                    require_nonempty(item, "%s.affected_items" % issue_id) for item in affected_items
                ]
                if len(set(issue["affected_items"])) != len(issue["affected_items"]):
                    raise ReviewGuardError("%s.affected_items存在重复项" % issue_id)
            issue["anchor"] = validate_anchor(issue.get("anchor"), file_format, issue_id)
            issue["issue_id"] = issue_id
            for field in ("location", "original", "reason", "correct", "action"):
                issue[field] = normalize_unit_superscripts(issue[field])
            comment = compose_comment(issue)
            hits = forbidden_hits("\n".join(issue[field] for field in ("reason", "correct", "action")))
            if hits:
                raise ReviewGuardError("%s命中禁用措辞：%s" % (issue_id, "、".join(hits)))
            normalized = normalize_text(comment)
            if normalized in baseline_texts:
                raise ReviewGuardError("%s与原有批注重复" % issue_id)
            if normalized in seen_comments:
                raise ReviewGuardError("%s与%s批注全文重复" % (issue_id, seen_comments[normalized]))
            seen_comments[normalized] = issue_id
            issue["comment"] = comment
            validated[file_id].append(issue)

    missing = sorted(set(state_files) - seen_files)
    if missing:
        raise ReviewGuardError("问题清单缺少待审核文件：%s" % "、".join(missing))

    validate_table_summaries(validated)
    return validated


def validate_table_summaries(validated: Dict[str, List[Dict[str, Any]]]) -> None:
    for file_id, issues in validated.items():
        points: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        summaries: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for issue in issues:
            table_id = str(issue.get("table_id") or "").strip()
            if issue["kind"] == "table_summary":
                if not table_id:
                    raise ReviewGuardError("%s.table_summary缺少table_id" % issue["issue_id"])
                affected = issue.get("affected_rows")
                if not isinstance(affected, list) or not affected:
                    raise ReviewGuardError("%s.affected_rows必须是非空数组" % issue["issue_id"])
                affected_rows = [require_nonempty(row, "%s.affected_rows" % issue["issue_id"]) for row in affected]
                if len(set(affected_rows)) != len(affected_rows):
                    raise ReviewGuardError("%s.affected_rows存在重复行" % issue["issue_id"])
                issue["affected_rows"] = affected_rows
                if SUMMARY_ACTION not in issue["action"]:
                    raise ReviewGuardError("%s修改动作必须包含“%s”" % (issue["issue_id"], SUMMARY_ACTION))
                summaries[table_id].append(issue)
            elif issue["kind"] == "point" and issue["original_kind"] == "value" and table_id:
                issue["row_label"] = require_nonempty(issue.get("row_label"), "%s.row_label" % issue["issue_id"])
                points[table_id].append(issue)

        for table_id, table_points in points.items():
            anchor_keys = [json_dump(issue["anchor"]) for issue in table_points]
            if len(anchor_keys) != len(set(anchor_keys)):
                raise ReviewGuardError("%s的%s存在多个数值问题共用同一单元格或文字锚点" % (file_id, table_id))
            if len(table_points) < 2:
                continue
            table_summaries = summaries.get(table_id, [])
            if not table_summaries:
                continue
            if len(table_summaries) != 1:
                raise ReviewGuardError("%s的%s最多允许1条table_summary" % (file_id, table_id))
            expected_rows = {issue["row_label"] for issue in table_points}
            actual_rows = set(table_summaries[0]["affected_rows"])
            if expected_rows != actual_rows:
                raise ReviewGuardError("%s的%s汇总行与逐单元格问题不一致" % (file_id, table_id))
        for table_id, table_summaries in summaries.items():
            if len(points.get(table_id, [])) < 2:
                raise ReviewGuardError("%s的%s没有至少2个数值point，不应建立table_summary" % (file_id, table_id))


def walk_nodes(value: Any) -> Iterable[Dict[str, Any]]:
    if isinstance(value, dict):
        yield value
        for item in value.values():
            yield from walk_nodes(item)
    elif isinstance(value, list):
        for item in value:
            yield from walk_nodes(item)


def resolve_docx_anchor(results: List[Dict[str, Any]], anchor: Dict[str, str]) -> str:
    paragraph_path = anchor["paragraph_path"]
    anchor_text = anchor["anchor_text"]
    if len(results) != 1:
        raise ReviewGuardError("Word段落锚点必须唯一命中：%s" % paragraph_path)
    paragraph = results[0]
    paragraph_text = str(paragraph.get("text") or "")
    if not paragraph_text.strip() or anchor_text not in paragraph_text:
        raise ReviewGuardError("Word段落不含非空锚点文字：%s -> %s" % (paragraph_path, anchor_text))
    runs = []
    for node in walk_nodes(paragraph.get("children", [])):
        if str(node.get("type") or "").lower() != "run":
            continue
        text = str(node.get("text") or "")
        path = str(node.get("path") or "")
        if path and text.strip() and anchor_text in text:
            runs.append(path)
    return runs[0] if len(runs) == 1 else paragraph_path


def preflight_anchors_for_file(
    officecli: str,
    item: Dict[str, Any],
    issues: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    if not issues:
        return []
    baseline = Path(item["baseline"])
    work_dir = Path(item["work_dir"])
    preflight = work_dir / ("anchor-preflight." + item["format"])
    command_file = work_dir / "anchor-preflight-commands.json"
    shutil.copy2(baseline, preflight)
    if sha256_file(preflight) != item["source_sha256"]:
        raise ReviewGuardError("锚点预检副本与不可变基线不一致：%s" % item["source"])

    keys: List[str] = []
    commands: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for issue in issues:
        if item["format"] == "docx":
            key = issue["anchor"]["paragraph_path"]
            command = {"command": "get", "path": key, "depth": 3}
        else:
            key = "/%s/%s" % (issue["anchor"]["sheet"], issue["anchor"]["cell"])
            command = {"command": "get", "path": key}
        if key not in seen:
            seen.add(key)
            keys.append(key)
            commands.append(command)

    try:
        items = officecli_batch(officecli, preflight, commands, command_file)
        results_by_key = {
            key: batch_item_results(result)
            for key, result in zip(keys, items)
        }
        resolved: List[Dict[str, Any]] = []
        for issue in issues:
            resolved_issue = dict(issue)
            if item["format"] == "docx":
                key = issue["anchor"]["paragraph_path"]
                target = resolve_docx_anchor(results_by_key[key], issue["anchor"])
            else:
                target = "/%s/%s" % (issue["anchor"]["sheet"], issue["anchor"]["cell"])
            resolved_issue["resolved_anchor"] = target
            resolved.append(resolved_issue)
        return resolved
    finally:
        for path in (preflight, command_file):
            if path.exists():
                path.unlink()


def assert_sources_unchanged(state: Dict[str, Any]) -> None:
    for item in state_file_map(state).values():
        source = Path(item["source"])
        if not source.is_file():
            raise ReviewGuardError("原件已不存在：%s" % source)
        current_hash = sha256_file(source)
        if current_hash != item["source_sha256"]:
            raise ReviewGuardError("原件SHA-256已变化，停止：%s" % source)


def docx_anchor_texts(path: Path) -> Dict[str, str]:
    with zipfile.ZipFile(path, "r") as archive:
        try:
            document_xml = archive.read("word/document.xml")
        except KeyError as exc:
            raise ReviewGuardError("DOCX缺少word/document.xml：%s" % path) from exc
    captured: Dict[str, List[str]] = defaultdict(list)
    active: List[str] = []
    for event, element in ET.iterparse(io.BytesIO(document_xml), events=("start", "end")):
        local = element.tag.rsplit("}", 1)[-1]
        if event == "start" and local == "commentRangeStart":
            comment_id = element.get(WORD_ID)
            if comment_id is not None:
                active.append(comment_id)
        elif event == "start" and local == "commentRangeEnd":
            comment_id = element.get(WORD_ID)
            if comment_id in active:
                active.remove(comment_id)
        elif event == "end" and local in {"t", "delText", "instrText"} and element.text:
            for comment_id in active:
                captured[comment_id].append(element.text)
        if event == "end":
            element.clear()
    return {comment_id: "".join(parts) for comment_id, parts in captured.items()}


def zip_check(path: Path, require_unzip: bool) -> Dict[str, Any]:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            bad_member = archive.testzip()
    except (OSError, zipfile.BadZipFile) as exc:
        raise ReviewGuardError("压缩包检查失败：%s；%s" % (path, exc)) from exc
    if bad_member:
        raise ReviewGuardError("压缩包成员损坏：%s -> %s" % (path, bad_member))
    result: Dict[str, Any] = {"zipfile_test": "passed", "unzip_tq": "not_available"}
    unzip = shutil.which("unzip")
    if require_unzip and unzip:
        process = subprocess.run([unzip, "-tq", str(path)], capture_output=True, text=True, check=False)
        if process.returncode != 0:
            raise ReviewGuardError("unzip -tq失败：%s；%s" % (path, process.stderr.strip() or process.stdout.strip()))
        result["unzip_tq"] = "passed"
    elif not require_unzip:
        result["unzip_tq"] = "not_applicable"
    return result


def prepare_command(args: argparse.Namespace) -> Dict[str, Any]:
    state_path = Path(args.state).expanduser().resolve()
    if state_path.exists():
        raise ReviewGuardError("状态文件已存在，拒绝覆盖：%s" % state_path)
    officecli = resolve_officecli(args.officecli)
    profile = args.profile
    inputs: List[Path] = []
    seen: set[Path] = set()
    for raw in args.input:
        source = Path(raw).expanduser().resolve()
        if source in seen:
            raise ReviewGuardError("待审核文件重复：%s" % source)
        seen.add(source)
        if not source.is_file():
            raise ReviewGuardError("待审核文件不存在：%s" % source)
        if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ReviewGuardError("仅支持DOCX/XLSX待审核文件：%s" % source)
        if is_reviewed_output(source):
            raise ReviewGuardError("待审核清单不得包含临时文件或既有审核批注版：%s" % source)
        inputs.append(source)
    if not inputs:
        raise ReviewGuardError("至少提供一个--input")

    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else None
    if output_dir and not output_dir.is_dir():
        raise ReviewGuardError("输出目录不存在：%s" % output_dir)

    stage_dir = Path(tempfile.mkdtemp(prefix="reviewing-reports-guard-"))
    files: List[Dict[str, Any]] = []
    occupied: set[Path] = set()
    page_probe_enabled = True
    page_probe_reason = ""
    try:
        for index, source in enumerate(inputs, start=1):
            file_id = "f%03d" % index
            file_format = source.suffix.lower().lstrip(".")
            work_dir = stage_dir / file_id
            work_dir.mkdir(parents=True)
            baseline = work_dir / ("baseline" + source.suffix.lower())
            stage = work_dir / ("review" + source.suffix.lower())
            read_copy = work_dir / ("prepare-read" + source.suffix.lower())
            source_hash = sha256_file(source)
            shutil.copy2(source, baseline)
            shutil.copy2(source, read_copy)
            if sha256_file(baseline) != source_hash or sha256_file(read_copy) != source_hash:
                raise ReviewGuardError("临时副本与原件哈希不一致：%s" % source)
            baseline_zip = zip_check(
                baseline,
                require_unzip=profile == "strict" and file_format == "docx",
            )
            baseline_comments = canonical_comments(query_comments(officecli, read_copy))
            snapshots = {
                mode: snapshot_hash(officecli, read_copy, mode)
                for mode in PROFILE_SNAPSHOT_MODES[profile]
            }
            validation: Optional[Dict[str, Any]] = None
            if profile == "strict":
                validation_copy = work_dir / ("validate-baseline" + source.suffix.lower())
                shutil.copy2(baseline, validation_copy)
                validation = validate_document(officecli, validation_copy)
                validation_copy.unlink()
            if profile != "strict":
                page_count = {
                    "status": "not_performed_by_profile",
                    "value": None,
                }
            elif file_format != "docx":
                page_count = {"status": "not_applicable", "value": None}
            elif page_probe_enabled:
                page_count = page_count_status(officecli, read_copy, file_format)
                if page_count.get("status") != "available":
                    page_probe_enabled = False
                    page_probe_reason = str(page_count.get("reason") or "页数接口不可用")
            else:
                page_count = {
                    "status": "not_performed",
                    "value": None,
                    "reason": "本任务此前页数接口不可用，已跳过重复探测：%s" % page_probe_reason,
                }
            read_copy.unlink()
            proposed_output = next_output_path(source, output_dir, occupied)
            occupied.add(proposed_output)
            files.append(
                {
                    "file_id": file_id,
                    "source": str(source),
                    "format": file_format,
                    "source_sha256": source_hash,
                    "baseline_comment_count": len(baseline_comments),
                    "baseline_comments": baseline_comments,
                    "baseline_snapshots": snapshots,
                    "baseline_validation": validation,
                    "baseline_page_count": page_count,
                    "baseline_zip": baseline_zip,
                    "work_dir": str(work_dir),
                    "baseline": str(baseline),
                    "stage": str(stage),
                    "proposed_output": str(proposed_output),
                }
            )
        for source, item in zip(inputs, files):
            if sha256_file(source) != item["source_sha256"]:
                raise ReviewGuardError("预检期间原件发生变化：%s" % source)
        state = {
            "version": STATE_VERSION,
            "status": "prepared",
            "profile": profile,
            "created_at": utc_now(),
            "officecli": officecli,
            "stage_dir": str(stage_dir),
            "state_path": str(state_path),
            "output_dir": str(output_dir) if output_dir else None,
            "files": files,
        }
        atomic_write_json(state_path, state, refuse_existing=True)
    except Exception:
        shutil.rmtree(stage_dir, ignore_errors=True)
        raise

    return {
        "success": True,
        "status": "prepared",
        "profile": profile,
        "state": str(state_path),
        "stage_dir": str(stage_dir),
        "files": [
            {
                "file_id": item["file_id"],
                "source": item["source"],
                "sha256": item["source_sha256"],
                "original_comments": item["baseline_comment_count"],
                "proposed_output": item["proposed_output"],
                "page_count": item["baseline_page_count"],
            }
            for item in files
        ],
        "issues_template": {
            "files": [{"file_id": item["file_id"], "issues": []} for item in files]
        },
    }


def write_command(args: argparse.Namespace) -> Dict[str, Any]:
    state_path = Path(args.state).expanduser().resolve()
    issues_path = Path(args.issues).expanduser().resolve()
    state = load_json(state_path, "状态文件")
    check_state(state, "prepared")
    officecli = resolve_officecli(args.officecli or state.get("officecli"))
    issues_payload = load_json(issues_path, "问题清单")
    validated = validate_issues(state, issues_payload)
    assert_sources_unchanged(state)

    for item in state_file_map(state).values():
        baseline = Path(item["baseline"])
        if sha256_file(baseline) != item["source_sha256"]:
            raise ReviewGuardError("临时基线与预检哈希不一致：%s" % baseline)

    resolved_by_file: Dict[str, List[Dict[str, Any]]] = {}
    for file_id, item in state_file_map(state).items():
        resolved_by_file[file_id] = preflight_anchors_for_file(
            officecli,
            item,
            validated[file_id],
        )

    for file_id, item in state_file_map(state).items():
        baseline = Path(item["baseline"])
        stage = Path(item["stage"])
        # 中断后仍可重试：每次均从不可变基线重建，不保留半写入状态。
        shutil.copy2(baseline, stage)
        if sha256_file(stage) != item["source_sha256"]:
            raise ReviewGuardError("重建临时批注版失败：%s" % stage)
        commands: List[Dict[str, Any]] = []
        for issue in resolved_by_file[file_id]:
            props = {"author": AUTHOR, "text": issue["comment"]}
            if item["format"] == "docx":
                props["range"] = "true"
            commands.append(
                {
                    "command": "add",
                    "parent": issue["resolved_anchor"],
                    "type": "comment",
                    "props": props,
                }
            )
        command_file = Path(item["work_dir"]) / "write-comments-commands.json"
        try:
            officecli_batch(
                officecli,
                stage,
                commands,
                command_file,
                require_results=False,
            )
        finally:
            if command_file.exists():
                command_file.unlink()

    assert_sources_unchanged(state)
    state["status"] = "written"
    state["written_at"] = utc_now()
    state["issues_path"] = str(issues_path)
    state["issues_sha256"] = sha256_file(issues_path)
    state["written_issues"] = resolved_by_file
    state["officecli"] = officecli
    atomic_write_json(state_path, state)
    return {
        "success": True,
        "status": "written",
        "profile": state_profile(state),
        "state": str(state_path),
        "files": [
            {
                "file_id": file_id,
                "source": item["source"],
                "new_comments": len(resolved_by_file[file_id]),
            }
            for file_id, item in state_file_map(state).items()
        ],
    }


def verify_comments_for_file(
    officecli: str,
    item: Dict[str, Any],
    expected: List[Dict[str, Any]],
) -> Dict[str, Any]:
    stage = Path(item["stage"])
    results = query_comments(officecli, stage)
    records = [comment_record(result) for result in results]
    baseline_count = int(item["baseline_comment_count"])
    if len(records) != baseline_count + len(expected):
        raise ReviewGuardError("%s批注数不闭合：原有%d，预期新增%d，实际%d" % (
            item["source"], baseline_count, len(expected), len(records)
        ))
    by_text: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    for record in records:
        by_text[normalize_text(record["text"])].append(record)

    baseline_identities = Counter(
        json_dump({key: value for key, value in comment.items() if key != "id"})
        for comment in item["baseline_comments"]
    )
    actual_identities = Counter(
        json_dump({key: value for key, value in record.items() if key != "id"})
        for record in records
    )
    missing_originals = baseline_identities - actual_identities
    if missing_originals:
        raise ReviewGuardError("%s的原有批注未完整保留" % item["source"])

    anchor_text_by_id = docx_anchor_texts(stage) if item["format"] == "docx" else {}
    for issue in expected:
        normalized = normalize_text(issue["comment"])
        matches = by_text.get(normalized, [])
        if len(matches) != 1:
            raise ReviewGuardError("%s未能逐条唯一回读完整批注：%s" % (item["source"], issue["issue_id"]))
        record = matches[0]
        if record["author"] != AUTHOR:
            raise ReviewGuardError("%s批注作者不是“%s”：%s" % (issue["issue_id"], AUTHOR, record["author"]))
        if not record["text"].strip():
            raise ReviewGuardError("%s批注文本为空" % issue["issue_id"])
        if unit_power_spellings(record["text"]) != unit_power_spellings(issue["comment"]):
            raise ReviewGuardError("%s回读单位角标字形与写入前不一致" % issue["issue_id"])
        if item["format"] == "docx":
            comment_id = record["id"]
            anchor_text = anchor_text_by_id.get(comment_id, "")
            if not anchor_text.strip():
                raise ReviewGuardError("%s的Word批注锚点为空" % issue["issue_id"])
            if issue["anchor"]["anchor_text"] not in anchor_text:
                raise ReviewGuardError("%s的Word批注锚点未包含预期文字" % issue["issue_id"])
        else:
            expected_path = "/%s/%s" % (issue["anchor"]["sheet"], issue["anchor"]["cell"])
            if record["anchored_to"] != expected_path and record["ref"].upper() != issue["anchor"]["cell"]:
                raise ReviewGuardError("%s的Excel批注未锚定目标单元格%s" % (issue["issue_id"], expected_path))

    expected_texts = {normalize_text(issue["comment"]) for issue in expected}
    baseline_texts = Counter(normalize_text(comment["text"]) for comment in item["baseline_comments"])
    actual_texts = Counter(normalize_text(record["text"]) for record in records)
    unexpected = actual_texts - baseline_texts - Counter(expected_texts)
    if unexpected:
        raise ReviewGuardError("%s存在未登记新增批注" % item["source"])
    return {"total": len(records), "original": baseline_count, "new": len(expected)}


def severity_distribution(issues: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = Counter(issue["severity"] for issue in issues)
    return {severity: counts.get(severity, 0) for severity in SEVERITIES}


def exclusive_publish(stage: Path, source: Path, output_dir: Optional[Path], occupied: set[Path]) -> Path:
    while True:
        destination = next_output_path(source, output_dir, occupied)
        try:
            with stage.open("rb") as source_handle, destination.open("xb") as destination_handle:
                shutil.copyfileobj(source_handle, destination_handle, length=1024 * 1024)
                destination_handle.flush()
                os.fsync(destination_handle.fileno())
            shutil.copystat(stage, destination)
            occupied.add(destination)
            return destination
        except FileExistsError:
            occupied.add(destination)
            continue
        except Exception:
            if destination.exists():
                destination.unlink()
            raise


def path_is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def cleanup_internal(
    state: Dict[str, Any],
    state_path: Path,
    issues_path: Optional[Path],
    remove_issues: bool = False,
) -> Dict[str, Any]:
    stage_dir = Path(state["stage_dir"]).resolve()
    temp_root = Path(tempfile.gettempdir()).resolve()
    if not path_is_within(stage_dir, temp_root) or not stage_dir.name.startswith("reviewing-reports-guard-"):
        raise ReviewGuardError("拒绝清理非守护脚本临时目录：%s" % stage_dir)
    shutil.rmtree(stage_dir, ignore_errors=False)
    removed: List[str] = [str(stage_dir)]
    retained: List[str] = []
    for candidate in (state_path, issues_path):
        if candidate is None or not candidate.exists():
            continue
        if candidate == state_path or remove_issues or path_is_within(candidate, temp_root):
            if not candidate.is_file():
                raise ReviewGuardError("拒绝清理非文件元数据路径：%s" % candidate)
            candidate.unlink()
            removed.append(str(candidate))
        else:
            retained.append(str(candidate))
    return {"removed": removed, "retained_non_temp": retained}


def verify_command(args: argparse.Namespace) -> Dict[str, Any]:
    state_path = Path(args.state).expanduser().resolve()
    issues_path = Path(args.issues).expanduser().resolve()
    state = load_json(state_path, "状态文件")
    check_state(state, "written")
    officecli = resolve_officecli(args.officecli or state.get("officecli"))
    if str(issues_path) != state.get("issues_path") or sha256_file(issues_path) != state.get("issues_sha256"):
        raise ReviewGuardError("问题清单在写入后发生变化，停止验收")
    issues_payload = load_json(issues_path, "问题清单")
    validated = validate_issues(state, issues_payload)
    profile = state_profile(state)
    snapshot_modes = PROFILE_SNAPSHOT_MODES[profile]
    assert_sources_unchanged(state)

    file_reports: List[Dict[str, Any]] = []
    for file_id, item in state_file_map(state).items():
        stage = Path(item["stage"])
        comment_check = verify_comments_for_file(officecli, item, validated[file_id])
        snapshots = {mode: snapshot_hash(officecli, stage, mode) for mode in snapshot_modes}
        if snapshots != item["baseline_snapshots"]:
            changed = sorted(mode for mode in snapshots if snapshots[mode] != item["baseline_snapshots"].get(mode))
            raise ReviewGuardError("%s可见正文/表格或结构发生变化：%s" % (item["source"], "、".join(changed)))
        zip_result = zip_check(
            stage,
            require_unzip=profile == "strict" and item["format"] == "docx",
        )
        officecli_new_errors: Any = "not_performed_by_profile"
        if profile == "strict":
            baseline_validation = item.get("baseline_validation")
            if not isinstance(baseline_validation, dict):
                raise ReviewGuardError("%s缺少strict档结构检查基线" % item["source"])
            validation_copy = Path(item["work_dir"]) / ("validate-output." + item["format"])
            shutil.copy2(stage, validation_copy)
            try:
                output_validation = validate_document(officecli, validation_copy)
            finally:
                if validation_copy.exists():
                    validation_copy.unlink()
            new_errors = signature_counter(output_validation) - signature_counter(baseline_validation)
            if new_errors:
                raise ReviewGuardError("%s的OfficeCLI结构检查新增%d类错误" % (item["source"], len(new_errors)))
            officecli_new_errors = 0
        baseline_page_count = item["baseline_page_count"]
        if profile == "strict" and baseline_page_count.get("status") == "available":
            page_count = page_count_status(officecli, stage, item["format"])
            if page_count.get("status") != "available" or page_count.get("value") != baseline_page_count.get("value"):
                raise ReviewGuardError("%s页数与原件不一致或无法回读" % item["source"])
        else:
            page_count = dict(baseline_page_count)
        snapshot_checks = {
            mode: ("passed" if mode in snapshot_modes else "not_performed_by_profile")
            for mode in ("text", "outline", "annotated")
        }
        file_reports.append(
            {
                "file_id": file_id,
                "source": item["source"],
                "source_sha256": item["source_sha256"],
                "comments": comment_check,
                "severity": severity_distribution(validated[file_id]),
                "checks": {
                    "source_unchanged": "passed",
                    "comment_readback": "passed",
                    "anchor_nonempty": "passed",
                    "duplicates": 0,
                    "forbidden_phrases": 0,
                    "unit_superscripts": "normalized_and_verified",
                    "snapshots": snapshot_checks,
                    "zip": zip_result,
                    "officecli_new_errors": officecli_new_errors,
                    "page_count": page_count,
                    "visual_review": "not_performed_by_design",
                },
            }
        )

    assert_sources_unchanged(state)

    outputs: List[Dict[str, str]] = []
    created: List[Path] = []
    if args.publish:
        output_dir = Path(state["output_dir"]).resolve() if state.get("output_dir") else None
        occupied: set[Path] = set()
        try:
            for item in state_file_map(state).values():
                source = Path(item["source"])
                stage = Path(item["stage"])
                destination = exclusive_publish(stage, source, output_dir, occupied)
                created.append(destination)
                if sha256_file(destination) != sha256_file(stage):
                    raise ReviewGuardError("发布文件与验收副本哈希不一致：%s" % destination)
                outputs.append({"source": str(source), "output": str(destination)})
            if len(outputs) != len(state_file_map(state)):
                raise ReviewGuardError("正式输出数量与待审核文件数量不一致")
        except Exception:
            for path in created:
                if path.exists():
                    path.unlink()
            raise

    for report in file_reports:
        for output in outputs:
            if output["source"] == report["source"]:
                report["output"] = output["output"]
                break

    result: Dict[str, Any] = {
        "success": True,
        "status": "published" if args.publish else "verified",
        "profile": profile,
        "verified_at": utc_now(),
        "output_count": len(outputs) if args.publish else 0,
        "target_count": len(file_reports),
        "files": file_reports,
    }
    if args.publish:
        result["cleanup"] = cleanup_internal(state, state_path, issues_path, remove_issues=True)
    else:
        state["status"] = "verified"
        state["verified_at"] = result["verified_at"]
        atomic_write_json(state_path, state)
    return result


def cleanup_command(args: argparse.Namespace) -> Dict[str, Any]:
    state_path = Path(args.state).expanduser().resolve()
    state = load_json(state_path, "状态文件")
    check_state(state)
    issues_path = Path(state["issues_path"]).resolve() if state.get("issues_path") else None
    cleanup = cleanup_internal(state, state_path, issues_path)
    return {"success": True, "status": "cleaned", "cleanup": cleanup}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="报告原件保全、原生批注写入与验收守护脚本")
    subparsers = parser.add_subparsers(dest="command", required=True)

    prepare = subparsers.add_parser("prepare", help="按验收档位预检、记录基线并建立临时副本")
    prepare.add_argument("--state", required=True, help="临时状态JSON路径")
    prepare.add_argument("--input", action="append", required=True, help="待审核DOCX/XLSX；可重复")
    prepare.add_argument(
        "--profile",
        choices=PROFILES,
        default=DEFAULT_PROFILE,
        help="验收档位；默认fast，可选balanced或strict",
    )
    prepare.add_argument("--output-dir", help="可选统一输出目录；默认原件同目录")
    prepare.add_argument("--officecli", help="officecli可执行文件路径或命令名")
    prepare.set_defaults(handler=prepare_command)

    write = subparsers.add_parser("write", help="校验问题清单并按文件原子批量写入原生批注")
    write.add_argument("--state", required=True)
    write.add_argument("--issues", required=True)
    write.add_argument("--officecli")
    write.set_defaults(handler=write_command)

    verify = subparsers.add_parser("verify", help="统一回读并可在通过后发布")
    verify.add_argument("--state", required=True)
    verify.add_argument("--issues", required=True)
    verify.add_argument("--publish", action="store_true", help="全量通过后发布正式输出并清理临时状态")
    verify.add_argument("--officecli")
    verify.set_defaults(handler=verify_command)

    cleanup = subparsers.add_parser("cleanup", help="清理本脚本建立的临时副本")
    cleanup.add_argument("--state", required=True)
    cleanup.set_defaults(handler=cleanup_command)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        result = args.handler(args)
    except ReviewGuardError as exc:
        print(json.dumps({"success": False, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - 未预期异常仍须返回结构化结果
        print(json.dumps({"success": False, "error": "未处理错误：%s" % exc}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
