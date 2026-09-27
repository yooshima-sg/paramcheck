"""チェック: パラメータシート(config-items-yamlの出力 or Excel)と、generate-collect-scriptで
収集したアーカイブ内の実設定ファイルを突き合わせる。

- アーカイブは <任意名>/collected/<元の絶対パス> の構造。ファイル名に "rod" を含めば
  ROD設定値、"ec2" を含めばEC2設定値と比較する。
- 期待する状態は ROD設定値/EC2設定値 の値から判定する:
    "…コメントアウト"          -> コメントアウト行が存在し、有効行が無いこと
    "記載なし"                 -> その設定項目の行が(コメント含め)存在しないこと
    "未設定…"/"設定なし"/"※設定項目なし" -> 有効行が無いこと（コメント行は許容）
    "有効化"/"有効…"           -> 有効行が存在すること（値は不問）
    "無効(…)" / "※…"注記のみ   -> テキストでは判定不能。行の存在のみ確認し review 扱い
    それ以外の具体値            -> 有効行が存在し、値が一致すること（末尾の "※…" 注記は比較から除外）
    空欄・ダッシュ類            -> skipped
- 設定階層: "global" はトップレベル。入れ子は "親 > 子" の連結表記。
  Apacheは <Tag args> 形式、Dovecotは "service imap-login" のようなブロックヘッダ。
- パラメータが "(ディレクティブ)" やダッシュ類の行は確認対象外(skipped)。
- RHELバージョン差の考慮（rod=RHEL7/el7, ec2=RHEL8/el8）:
    行/ファイルが見つからない場合、reference/rhel-pkgs/<release>/configs のパッケージ既定ファイルを参照し、
    その版の既定にも存在せず、もう一方の版には存在する項目は not_in_release として区別する。
    シートの "※旧設定名:xxx <値>" 注記、または調査で確定した改名表(_KNOWN_RENAMES)により
    旧名で再照合し、見つかれば renamed（値注記があれば値も比較）とする。

- 漏れチェック（既定で有効、--no-coverage で無効化）:
    設定が格納されている /etc/httpd /etc/postfix /etc/dovecot /etc/vsftpd 配下の設定ファイルを
    走査し、有効な設定のうちシートに記載が無いものを uncovered、シートが全く言及していない
    ファイルを uncovered_file として報告する。

対応フォーマット: Apache httpd (/etc/httpd/**)、Postfix main.cf・vsftpd.conf (key=value)、
Postfix master.cf (列形式。"-o key=value" 継続行はサービス行に畳み込む)、Dovecot (/etc/dovecot/**)。
それ以外は「先頭トークン=名前」の汎用パーサで扱う。
"""

from __future__ import annotations

import argparse
import json
import re
import shlex
import sys
import tarfile
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from . import config_items_yaml
from .excel_config_path import _is_placeholder

NAME = "compare-config"
DESCRIPTION = "パラメータシートと収集アーカイブ内の実設定ファイルを突き合わせ、差異をYAMLで報告する"

_SKIP_PARAMS = {"(ディレクティブ)"}
_SKIP_VALUES = {"(ディレクティブ)", "(空)"}


# --- 設定ファイルのパース -----------------------------------------------------


@dataclass
class Entry:
    name: str
    args: str
    active: bool
    scope: list[str]  # 外側→内側の正規化済みブロックヘッダ
    line_no: int
    raw: str
    kind: str = "directive"  # "directive" | "block"
    scope_ord: list[int] = field(default_factory=list)  # 同名ブロックの出現順(1始まり、コメントブロックは0)
    header_norm: str = ""  # kind=="block" のとき、正規化済みブロック見出し（設定階層との比較用）


class _ScopeStack:
    """ブロックの入れ子を追跡する。コメントアウトされたブロック(#<Tag>〜#</Tag>)も追跡し、
    その内側のコメント行が正しい階層で照合できるようにする。"""

    def __init__(self) -> None:
        self._stack: list[tuple[str, int, bool]] = []
        self._counts: dict[tuple[tuple[str, ...], str], int] = {}

    def push(self, norm: str, active: bool) -> None:
        ordinal = 0
        if active:
            key = (tuple(s for s, _, _ in self._stack), norm)
            self._counts[key] = self._counts.get(key, 0) + 1
            ordinal = self._counts[key]
        self._stack.append((norm, ordinal, active))

    def drop_commented(self) -> None:
        """末尾に残っているコメントアウト由来の階層を捨てる。
        有効な行がコメントアウトされたブロックの内側にあることはないため、
        「開始タグだけコメントアウトされ閉じタグが無い」ような実ファイルの崩れで
        以降の階層がずれるのを防ぐ。"""
        while self._stack and not self._stack[-1][2]:
            self._stack.pop()

    def pop(self, active: bool, tag: str | None = None) -> None:
        """tag(Apacheの閉じタグ名)が分かる場合は、同名の開始ブロックまで巻き戻す。
        開始タグだけがコメントアウトされていて閉じタグが無い、といった実ファイルの
        崩れに引きずられて以降の階層がずれるのを防ぐ。"""
        if not self._stack:
            return
        if tag is not None:
            for i in range(len(self._stack) - 1, -1, -1):
                if self._stack[i][0].split(" ", 1)[0] == tag.lower():
                    del self._stack[i:]
                    return
            return
        if active:
            while self._stack and not self._stack[-1][2]:
                self._stack.pop()
            if self._stack:
                self._stack.pop()
        elif not self._stack[-1][2]:
            self._stack.pop()

    def names(self) -> list[str]:
        return [s for s, _, _ in self._stack]

    def ordinals(self) -> list[int]:
        return [o for _, o, _ in self._stack]


_APACHE_OPEN_RE = re.compile(r"^<([A-Za-z]\w*)(\s[^>]*)?>\s*$")
_APACHE_CLOSE_RE = re.compile(r"^</([A-Za-z]\w*)>\s*$")


def _strip_quotes(token: str) -> str:
    if len(token) >= 2 and token[0] == token[-1] and token[0] in "\"'":
        return token[1:-1]
    return token


def _norm_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _split_tokens(text: str) -> list[str]:
    try:
        return shlex.split(text, posix=True)
    except ValueError:
        return text.split()


def _norm_apache_tag(text: str) -> str:
    """'<Directory "/var/www">' -> 'directory /var/www' 形式に正規化する。"""
    inner = _norm_ws(text).strip("<>").strip()
    tokens = [_strip_quotes(t) for t in _split_tokens(inner)]
    return " ".join(tokens).lower()


def _norm_block_header(text: str) -> str:
    tokens = [_strip_quotes(t) for t in _split_tokens(_norm_ws(text))]
    return " ".join(tokens).lower()


def _split_comment(line: str) -> tuple[bool, str]:
    """行頭の # を剥がし、(active, body) を返す。"""
    stripped = line.strip()
    if stripped.startswith("#"):
        return False, stripped.lstrip("#").strip()
    return True, stripped


def parse_apache(lines: list[str]) -> list[Entry]:
    entries: list[Entry] = []
    scope = _ScopeStack()
    buf = ""
    start_no = 0
    for idx, raw in enumerate(lines, start=1):
        if not buf:
            start_no = idx
        line = raw.rstrip("\n")
        if line.rstrip().endswith("\\"):
            buf += line.rstrip()[:-1] + " "
            continue
        body_line = buf + line
        buf = ""
        active, body = _split_comment(body_line)
        if not body:
            continue
        if active:
            scope.drop_commented()
        close = _APACHE_CLOSE_RE.match(body)
        if close:
            scope.pop(active, close.group(1))
            continue
        m = _APACHE_OPEN_RE.match(body)
        if m:
            name = m.group(1)
            args = " ".join(_split_tokens(m.group(2) or ""))
            norm = _norm_apache_tag(body)
            entries.append(Entry(name, args, active, scope.names(), start_no, body_line.strip(), "block", scope.ordinals(), norm))
            scope.push(norm, active)
            continue
        tokens = _split_tokens(body)
        if not tokens:
            continue
        entries.append(Entry(tokens[0], body[len(tokens[0]):].strip(), active, scope.names(), start_no,
                             body_line.strip(), "directive", scope.ordinals()))
    return entries


def parse_postfix(lines: list[str]) -> list[Entry]:
    entries: list[Entry] = []
    for idx, raw in enumerate(lines, start=1):
        line = raw.rstrip("\n")
        if not line.strip():
            continue
        if line[0] in " \t" and entries and entries[-1].active:
            entries[-1].args = (entries[-1].args + " " + line.strip()).strip()
            continue
        active, body = _split_comment(line)
        if "=" not in body:
            continue
        key, value = body.split("=", 1)
        key = key.strip()
        if not key or " " in key:
            continue
        entries.append(Entry(key, value.strip(), active, [], idx, line.strip()))
    return entries


def parse_dovecot(lines: list[str]) -> list[Entry]:
    entries: list[Entry] = []
    scope = _ScopeStack()
    for idx, raw in enumerate(lines, start=1):
        active, body = _split_comment(raw.rstrip("\n"))
        if not body:
            continue
        if active:
            scope.drop_commented()
        if body == "}":
            scope.pop(active)
            continue
        if body.endswith("{") and "=" not in body:
            header = body[:-1].strip()
            tokens = _split_tokens(header)
            name = tokens[0] if tokens else ""
            norm = _norm_block_header(header)
            entries.append(Entry(name, " ".join(tokens[1:]), active, scope.names(), idx, raw.strip(), "block", scope.ordinals(), norm))
            scope.push(norm, active)
            continue
        if body.startswith("!include"):
            tokens = body.split(None, 1)
            entries.append(Entry(tokens[0], tokens[1].strip() if len(tokens) > 1 else "", active, scope.names(), idx,
                                 raw.strip(), "directive", scope.ordinals()))
            continue
        if "=" in body:
            key, value = body.split("=", 1)
            entries.append(Entry(key.strip(), value.strip(), active, scope.names(), idx, raw.strip(), "directive", scope.ordinals()))
            continue
        tokens = body.split(None, 1)
        entries.append(Entry(tokens[0], tokens[1].strip() if len(tokens) > 1 else "", active, scope.names(), idx,
                             raw.strip(), "directive", scope.ordinals()))
    return entries


def parse_master_cf(lines: list[str]) -> list[Entry]:
    """postfix master.cf は列形式:
      service type private unpriv chroot wakeup maxproc command
    行頭が空白の "-o key=value" 継続行は、パラメータシート側でもサービス行の値に
    含めて記載されているため、独立した項目にせずサービス行のargsに畳み込む。"""
    entries: list[Entry] = []
    current: Entry | None = None
    for idx, raw in enumerate(lines, start=1):
        line = raw.rstrip("\n")
        if not line.strip():
            continue
        active, body = _split_comment(line)
        if not body:
            continue
        if line[0] in " \t":
            if current is not None and current.active == active:
                current.args = f"{current.args} {body}".strip()
                current.raw = f"{current.raw} {line.strip()}"
            continue
        toks = body.split()
        if len(toks) < 2:
            current = None
            continue
        current = Entry(toks[0], " ".join(toks[1:]), active, [], idx, line.strip())
        entries.append(current)
    return entries


def parse_generic(lines: list[str]) -> list[Entry]:
    entries: list[Entry] = []
    for idx, raw in enumerate(lines, start=1):
        active, body = _split_comment(raw.rstrip("\n"))
        if not body:
            continue
        tokens = body.split(None, 1)
        entries.append(Entry(tokens[0], tokens[1].strip() if len(tokens) > 1 else "", active, [], idx, raw.strip()))
    return entries


def _parser_for(path: str):
    if path.startswith("/etc/httpd/"):
        return parse_apache, "apache"
    if path.endswith("master.cf"):
        return parse_master_cf, "master.cf"
    if path.startswith(("/etc/postfix/", "/etc/vsftpd/")):
        # どちらも key = value / key=value 形式
        return parse_postfix, "keyvalue"
    if path.startswith("/etc/dovecot/"):
        return parse_dovecot, "dovecot"
    return parse_generic, "generic"


# 設定が格納されているディレクトリ（漏れチェックの走査対象）
COVERAGE_DIRS = ("/etc/httpd", "/etc/postfix", "/etc/dovecot", "/etc/vsftpd")


def _is_config_file(path: str) -> bool:
    """走査対象の設定ファイルか。ルックアップテーブル(access/canonical等)・ftpusers・magic・
    証明書・スクリプトは設定ファイルではないので除外する。"""
    name = Path(path).name
    return name.endswith((".conf", ".conf.ext")) or name in ("main.cf", "master.cf")


# --- パラメータシート側の解釈 -------------------------------------------------


@dataclass
class Target:
    name: str
    prefix_args: list[str]
    annotation: str | None
    hierarchy: list[str]  # 正規化済み。globalなら []
    hierarchy_ord: list[int | None] = field(default_factory=list)  # "#N" 指定があればN、無ければNone


_ANNOTATION_RE = re.compile(r"\s*[（(]([^()（）]*)[)）]\s*$")


def _parse_param(param: str) -> tuple[str, list[str], str | None]:
    text = _norm_ws(param)
    annotation = None
    m = _ANNOTATION_RE.search(text)
    if m:
        annotation = m.group(1).strip()
        text = text[: m.start()].strip()
    tokens = _split_tokens(text)
    if not tokens:
        return "", [], annotation
    return tokens[0], [_strip_quotes(t) for t in tokens[1:]], annotation


_ORDINAL_RE = re.compile(r"\s*#\s*(\d+)\s*$")


def _parse_hierarchy(text: object, fmt: str) -> tuple[list[str], list[int | None]]:
    """設定階層 -> (正規化ブロック列, 出現順指定列)。
    "親 > 子" の連結表記、および同名ブロックを区別する "<Tag> #2" の順序指定に対応。"""
    s = _norm_ws(str(text or ""))
    if not s or s.lower() == "global":
        return [], []
    names: list[str] = []
    ords: list[int | None] = []
    if "<" in s:
        for m in re.finditer(r"(<[^>]+>)(?:\s*#\s*(\d+))?", s):
            names.append(_norm_apache_tag(m.group(1)))
            ords.append(int(m.group(2)) if m.group(2) else None)
        return names, ords
    for part in (p.strip() for p in s.split(">") if p.strip()):
        m = _ORDINAL_RE.search(part)
        ordinal = int(m.group(1)) if m else None
        if m:
            part = part[: m.start()].strip()
        names.append(_norm_block_header(part))
        ords.append(ordinal)
    return names, ords


def _norm_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = _norm_ws(str(value))
    return _strip_quotes(text)


def _values_equal(expected: str, actual: str) -> bool:
    """引用符・エスケープ用バックスラッシュの有無と大文字小文字の差は同値とみなす
    （Apacheの On/on, None/none、LogFormat内の \\" 等）。"""
    strip = lambda s: re.sub(r"[\"'\\]", "", s).lower()
    return expected == actual or strip(expected) == strip(actual)


def _candidate_values(e: Entry, target: Target) -> list[str]:
    """突合に使う実ファイル側の値の候補を返す。
    シート側の値が「値のみ」「ディレクティブ全文」「ニックネーム(LogFormat combined等)抜き」の
    いずれの書き方でも一致できるようにする。"""
    toks = _split_tokens(e.args)
    values = [_norm_value(e.args), _norm_value(" ".join(toks[len(target.prefix_args):])), _norm_value(f"{e.name} {e.args}")]
    if target.annotation:
        without = [t for t in toks if t.lower() != target.annotation.lower()]
        if len(without) != len(toks):
            values.append(_norm_value(" ".join(without)))
    return values


# --- 突き合わせ本体 -----------------------------------------------------------


@dataclass
class Archive:
    label: str
    env: str  # "rod" | "ec2"
    collected_root: Path
    _tmp: tempfile.TemporaryDirectory | None = field(default=None, repr=False)


def _classify_env(name: str) -> str | None:
    lower = name.lower()
    has_rod = "rod" in lower
    has_ec2 = "ec2" in lower
    if has_rod and not has_ec2:
        return "rod"
    if has_ec2 and not has_rod:
        return "ec2"
    return None


def _find_collected_root(base: Path) -> Path | None:
    if (base / "collected").is_dir():
        return base / "collected"
    for child in sorted(base.iterdir()):
        if child.is_dir() and (child / "collected").is_dir():
            return child / "collected"
    return None


def open_archive(path: Path) -> Archive:
    env = _classify_env(path.name)
    if env is None:
        raise ValueError(f"アーカイブ名から rod/ec2 を判別できません（名前に rod か ec2 を含めてください）: {path.name}")
    tmp = None
    if path.is_dir():
        base = path
    else:
        tmp = tempfile.TemporaryDirectory(prefix="datacheck_")
        with tarfile.open(path) as tar:
            tar.extractall(tmp.name, filter="data")
        base = Path(tmp.name)
    root = _find_collected_root(base)
    if root is None:
        raise ValueError(f"アーカイブ内に collected/ ディレクトリが見つかりません: {path}")
    return Archive(label=path.name, env=env, collected_root=root, _tmp=tmp)


class _FileCache:
    def __init__(self, root: Path) -> None:
        self.root = root
        self._cache: dict[str, list[Entry] | None] = {}

    def entries(self, abs_path: str) -> list[Entry] | None:
        if abs_path not in self._cache:
            fs_path = self.root / abs_path.lstrip("/")
            if not fs_path.is_file():
                self._cache[abs_path] = None
            else:
                parser, _ = _parser_for(abs_path)
                with fs_path.open(encoding="utf-8", errors="replace") as f:
                    self._cache[abs_path] = parser(f.readlines())
        return self._cache[abs_path]


def _scope_matches(e: Entry, target: Target) -> bool:
    if not target.hierarchy:
        return not e.scope
    n = len(target.hierarchy)
    if len(e.scope) < n or e.scope[-n:] != target.hierarchy:
        return False
    ords = e.scope_ord[-n:] if e.scope_ord else [0] * n
    return all(want is None or want == got for want, got in zip(target.hierarchy_ord, ords))


def _prefix_ok(e: Entry, target: Target) -> bool:
    if not target.prefix_args:
        return True
    toks = [_strip_quotes(t) for t in _split_tokens(e.args)]
    return toks[: len(target.prefix_args)] == target.prefix_args


def _find_candidates(entries: list[Entry], target: Target) -> list[Entry]:
    found = [e for e in entries
             if e.name.lower() == target.name.lower() and _scope_matches(e, target) and _prefix_ok(e, target)]
    if not found:
        # master.cf は "qmgr unix n - n 300 1 oqmgr" のようにサービス名と実行コマンドが異なることがあり、
        # シート側はコマンド名(最終列)で項目を識別している場合がある(qmgr と oqmgr を別項目として扱う等)
        found = [e for e in entries
                 if e.args and _split_tokens(e.args)[-1].lower() == target.name.lower() and _scope_matches(e, target)]
    if target.annotation and len(found) > 1:
        ann = target.annotation.lower()
        narrowed = [e for e in found if ann in [t.lower() for t in _split_tokens(e.args)]]
        if narrowed:
            found = narrowed
    return found


def _expected_kind(expected: str) -> str:
    """ROD設定値/EC2設定値の値から、実ファイルに期待する状態を判定する。"""
    if not expected or expected in _SKIP_VALUES or _is_placeholder(expected):
        return "skip"
    if "コメントアウト" in expected:
        return "commented"
    if expected.startswith("記載なし"):
        return "absent_strict"
    if expected.startswith(("未設定", "設定なし")) or "設定項目なし" in expected:
        return "absent"
    if expected.startswith("有効"):
        return "active_presence"
    if expected.startswith(("無効", "※")):
        return "review"
    return "active_value"


def _expected_for_compare(expected: str) -> str:
    """'all -SSLv2 ※RODはSSLv3が有効' のような末尾の※注記を比較対象から外す。"""
    return re.split(r"\s*※", expected, maxsplit=1)[0].strip()


def _lines_view(entries: list[Entry]) -> list[dict]:
    return [{"line": e.line_no, "active": e.active, "text": e.raw} for e in entries]


# --- RHELバージョン差の考慮（ROD=RHEL7 / EC2=RHEL8） ---------------------------

# アーカイブ種別 -> reference/rhel-pkgs 配下のリリース名
_ENV_RELEASE = {"rod": "el7", "ec2": "el8"}

# 調査で確定した改名（新名 -> 旧名）。シートの「※旧設定名:」注記があればそちらを優先する。
# 根拠: reference/rhel-pkgs/{el7,el8}/configs/dovecot/10-ssl.conf の差分、および
# Dovecot公式アップグレードガイド (https://doc.dovecot.org/2.3/installation_guide/upgrading/from-2.2-to-2.3/)
_KNOWN_RENAMES = {
    "el7": {
        "ssl_min_protocol": "ssl_protocols",
        "ssl_dh": "ssl_dh_parameters_length",
    },
}

_RENAME_NOTE_RE = re.compile(r"※\s*旧設定名\s*[:：]\s*(\S+)(?:\s*<([^>]*)>)?")


class _Reference:
    """reference/rhel-pkgs/<release>/configs/<pkg>/<file> にあるパッケージ既定の設定ファイルを参照し、
    「その項目/ファイルがそのRHEL版にそもそも存在するか」を判定する。"""

    def __init__(self, root: Path | None) -> None:
        self.root = root if root is not None and root.is_dir() else None
        self._cache: dict[tuple[str, str], list[Entry] | None] = {}

    @property
    def available(self) -> bool:
        return self.root is not None

    def _stock_file(self, release: str, abs_path: str) -> Path | None:
        """configs/ 配下は実際の絶対パス構造（configs/etc/httpd/conf/httpd.conf）を保持している。"""
        path = self.root / release / "configs" / abs_path.lstrip("/")
        # postfixのel8は編集前の既定値が main.cf.proto 側にある
        proto = path.with_name(path.name + ".proto")
        if proto.is_file():
            return proto
        return path if path.is_file() else None

    def stock_entries(self, release: str, abs_path: str) -> list[Entry] | None:
        key = (release, abs_path)
        if key not in self._cache:
            path = self._stock_file(release, abs_path)
            if path is None:
                self._cache[key] = None
            else:
                parser, _ = _parser_for(abs_path)
                with path.open(encoding="utf-8", errors="replace") as f:
                    self._cache[key] = parser(f.readlines())
        return self._cache[key]

    def file_exists(self, release: str, abs_path: str) -> bool:
        return self.stock_entries(release, abs_path) is not None

    def has_target(self, release: str, abs_path: str, target: Target) -> bool | None:
        """階層は無視し、名前(+LoadModule等の識別引数)が既定ファイルに登場するか。"""
        entries = self.stock_entries(release, abs_path)
        if entries is None:
            return None
        loose = Target(target.name, target.prefix_args, target.annotation, [], [])
        return any(e.name.lower() == loose.name.lower() and _prefix_ok(e, loose) for e in entries)

    def has_block(self, release: str, abs_path: str, norm_header: str) -> bool | None:
        """設定階層に書かれたブロック(正規化済み)が既定ファイルに存在するか。"""
        entries = self.stock_entries(release, abs_path)
        if entries is None:
            return None
        return any(e.kind == "block" and e.header_norm == norm_header for e in entries)


def _other_release(release: str) -> str:
    return "el8" if release == "el7" else "el7"


def _rename_hint(expected: str, release: str, name: str) -> tuple[str | None, str | None]:
    """(旧設定名, 注記に書かれた値) を返す。シート注記 > 既知の改名表。"""
    m = _RENAME_NOTE_RE.search(expected)
    if m:
        return m.group(1), (m.group(2).strip() if m.group(2) else None)
    old = _KNOWN_RENAMES.get(release, {}).get(name)
    return old, None


def _apply_release_awareness(res: dict, kind: str, expected: str, target: Target, entries: list[Entry] | None,
                             archive: Archive, reference: _Reference) -> None:
    """missing / file_missing を、RHELバージョン差（項目が存在しない・改名された）として再分類する。"""
    release = _ENV_RELEASE.get(archive.env)
    if release is None:
        return
    path = res["設定ファイルパス"]

    if res["status"] == "file_missing":
        if reference.available and not reference.file_exists(release, path) and reference.file_exists(_other_release(release), path):
            res.update(status="not_in_release",
                       detail=f"{release}のパッケージ既定にこのファイルは存在しない（RHELバージョン差）")
        return

    if res["status"] != "missing" or entries is None:
        return

    old_name, old_value = _rename_hint(expected, release, target.name)
    if old_name:
        old_target = Target(old_name, target.prefix_args, target.annotation, target.hierarchy, target.hierarchy_ord)
        old_cands = _find_candidates(entries, old_target)
        if old_cands:
            res["matched_lines"] = _lines_view(old_cands)
            active = [e for e in old_cands if e.active]
            if old_value:
                actuals = [_norm_value(e.args) for e in active]
                if any(_values_equal(_norm_value(old_value), v) for e in active for v in _candidate_values(e, old_target)):
                    res.update(status="renamed", detail=f"旧設定名 {old_name} で存在・値一致（{old_value}）")
                elif active:
                    res.update(status="value_mismatch", detail=f"旧設定名 {old_name} で存在するが値が不一致（期待 {old_value}）",
                               actual_values=actuals)
                else:
                    res.update(status="renamed", detail=f"旧設定名 {old_name} がコメントアウト行として存在（期待値 {old_value} は未比較）")
            else:
                state = "有効行" if active else "コメントアウト行"
                res.update(status="renamed", detail=f"旧設定名 {old_name} で存在（{state}）")
            return

    if reference.available:
        other = _other_release(release)
        here = reference.has_target(release, path, target)
        there = reference.has_target(other, path, target)
        if here is False and there:
            res.update(status="not_in_release",
                       detail=f"{release}のパッケージ既定ファイルにこの項目は存在しない（RHELバージョン差）")
            return
        if target.hierarchy:
            block = target.hierarchy[-1]
            if reference.has_block(release, path, block) is False and reference.has_block(other, path, block):
                res.update(status="not_in_release",
                           detail=f"設定階層のブロックが{release}のパッケージ既定ファイルに存在しない（RHELバージョン差。"
                                  f"ブロック構文の違いや新設サービスの可能性）")


def compare_row(row: dict, archive: Archive, cache: _FileCache, reference: _Reference | None = None) -> list[dict]:
    env_col = "ROD設定値" if archive.env == "rod" else "EC2設定値"
    expected_raw = row.get(env_col)
    base = {
        "archive": archive.label,
        "env": archive.env,
        "sheet": row.get("sheet"),
        "No": row.get("No"),
        "設定項目名": row.get("設定項目名"),
        "設定階層": row.get("設定階層"),
        "パラメータ": row.get("パラメータ"),
        "expected_value": expected_raw,
    }

    param = _norm_ws(str(row.get("パラメータ") or ""))
    paths = str(row.get("設定ファイルパス") or "").split()
    if not paths:
        return [{**base, "設定ファイルパス": None, "status": "skipped", "detail": "設定ファイルパスなし"}]
    if not param or param in _SKIP_PARAMS or _is_placeholder(param):
        return [{**base, "設定ファイルパス": p, "status": "skipped", "detail": "パラメータが確認対象外"} for p in paths]

    expected = _norm_value(expected_raw)
    kind = _expected_kind(expected)
    if kind == "skip":
        return [{**base, "設定ファイルパス": p, "status": "skipped", "detail": f"{env_col}が空または確認対象外"} for p in paths]

    reference = reference or _Reference(None)
    results = []
    for path in paths:
        res = {**base, "設定ファイルパス": path}
        entries = cache.entries(path)
        _, fmt = _parser_for(path)
        name, prefix_args, annotation = _parse_param(param)
        hier, hier_ord = _parse_hierarchy(row.get("設定階層"), fmt)
        target = Target(name, prefix_args, annotation, hier, hier_ord)
        if entries is None:
            res.update(status="file_missing", detail="アーカイブ内にファイルなし")
            _apply_release_awareness(res, kind, expected, target, None, archive, reference)
            results.append(res)
            continue
        cands = _find_candidates(entries, target)
        active = [e for e in cands if e.active]
        res["matched_lines"] = _lines_view(cands)

        if kind == "active_value":
            if not cands:
                res.update(status="missing", detail="有効行もコメント行も見つからない")
            elif not active:
                res.update(status="not_active", detail="コメントアウト行のみ存在（有効行なし）")
            else:
                actuals = []
                matched = False
                expected_cmp = _expected_for_compare(expected)
                for e in active:
                    actuals.append(_norm_value(e.args))
                    if any(_values_equal(expected_cmp, v) for v in _candidate_values(e, target)):
                        matched = True
                if matched:
                    res.update(status="ok", detail="有効行あり・値一致")
                else:
                    res.update(status="value_mismatch", detail="有効行はあるが値が不一致", actual_values=actuals)
        elif kind == "active_presence":
            if not cands:
                res.update(status="missing", detail="有効化のはずが行が見つからない")
            elif not active:
                res.update(status="not_active", detail="有効化のはずがコメントアウト行のみ")
            else:
                res.update(status="ok", detail="有効行あり")
        elif kind == "commented":
            if not cands:
                res.update(status="missing", detail="コメントアウト行が見つからない")
            elif active:
                res.update(status="unexpected_active", detail="コメントアウトのはずが有効行が存在")
            else:
                res.update(status="ok", detail="コメントアウト行あり")
        elif kind == "absent_strict":
            if cands:
                res.update(status="unexpected_present", detail="記載なしのはずが行が存在")
            else:
                res.update(status="ok", detail="記載なし")
        elif kind == "absent":
            if active:
                res.update(status="unexpected_present", detail="未設定のはずが有効行が存在")
            elif cands:
                res.update(status="ok", detail="有効行なし（コメント行のみ存在）")
            else:
                res.update(status="ok", detail="有効行なし")
        else:  # review: 「無効(xxx無効)」など条件依存の無効はテキストだけで判定できない
            if not cands:
                res.update(status="missing", detail="行が見つからない")
            else:
                res.update(status="review", detail="条件依存の無効(IfModule等)のため行の存在のみ確認。目視確認が必要")
        _apply_release_awareness(res, kind, expected, target, entries, archive, reference)
        results.append(res)
    return results


def _row_targets(rows: list[dict]) -> dict[str, list[Target]]:
    """設定ファイルパス -> そのファイルでシートが言及しているTargetの一覧。"""
    targets: dict[str, list[Target]] = {}
    for row in rows:
        param = _norm_ws(str(row.get("パラメータ") or ""))
        for path in str(row.get("設定ファイルパス") or "").split():
            entry = targets.setdefault(path, [])
            if not param or param in _SKIP_PARAMS or _is_placeholder(param):
                continue
            _, fmt = _parser_for(path)
            name, prefix_args, annotation = _parse_param(param)
            hier, hier_ord = _parse_hierarchy(row.get("設定階層"), fmt)
            entry.append(Target(name, prefix_args, annotation, hier, hier_ord))
    return targets


def check_coverage(rows: list[dict], archive: Archive, cache: _FileCache) -> list[dict]:
    """アーカイブ内の設定ディレクトリを走査し、シートに記載が無い有効設定（漏れ）を洗い出す。"""
    targets = _row_targets(rows)
    results: list[dict] = []
    root = archive.collected_root

    for coverage_dir in COVERAGE_DIRS:
        base = root / coverage_dir.lstrip("/")
        if not base.is_dir():
            continue
        for fs_path in sorted(p for p in base.rglob("*") if p.is_file()):
            abs_path = "/" + str(fs_path.relative_to(root))
            if not _is_config_file(abs_path):
                continue
            entries = cache.entries(abs_path) or []
            active = [e for e in entries if e.active and e.kind == "directive"]
            base_res = {
                "archive": archive.label,
                "env": archive.env,
                "sheet": None,
                "No": None,
                "設定項目名": None,
                "設定ファイルパス": abs_path,
                "expected_value": None,
            }

            if abs_path not in targets:
                if active:
                    results.append({
                        **base_res,
                        "設定階層": None,
                        "パラメータ": None,
                        "status": "uncovered_file",
                        "detail": f"シートに記載の無いファイル（有効設定 {len(active)} 件）",
                        "matched_lines": _lines_view(active[:5]),
                    })
                continue

            file_targets = targets[abs_path]
            for e in active:
                if any(e.name.lower() == t.name.lower() and _scope_matches(e, t) and _prefix_ok(e, t)
                       for t in file_targets):
                    continue
                scope = " > ".join(e.scope) if e.scope else "global"
                results.append({
                    **base_res,
                    "設定階層": scope,
                    "パラメータ": f"{e.name} {e.args}".strip(),
                    "status": "uncovered",
                    "detail": "実ファイルに有効な設定があるがシートに記載が無い",
                    "matched_lines": _lines_view([e]),
                })
    return results


def compare(rows: list[dict], archives: list[Archive], reference: _Reference | None = None,
            coverage: bool = True) -> list[dict]:
    results: list[dict] = []
    for archive in archives:
        cache = _FileCache(archive.collected_root)
        for row in rows:
            results.extend(compare_row(row, archive, cache, reference))
        if coverage:
            results.extend(check_coverage(rows, archive, cache))
    return results


def _load_rows(path: Path) -> list[dict]:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return config_items_yaml.extract_workbook(path)
    if suffix in (".yaml", ".yml"):
        with path.open(encoding="utf-8") as f:
            data = yaml.safe_load(f)
    elif suffix == ".json":
        with path.open(encoding="utf-8") as f:
            data = json.load(f)
    else:
        raise ValueError(f"未対応の入力形式です(.xlsx/.xlsm/.yaml/.json): {path}")
    if not isinstance(data, list):
        raise ValueError("パラメータシートのデータはレコードのリストである必要があります")
    return data


def _summarize(results: list[dict]) -> dict:
    summary: dict[str, dict[str, int]] = {}
    for r in results:
        per = summary.setdefault(r["archive"], {})
        per[r["status"]] = per.get(r["status"], 0) + 1
    return summary


# --- 人間向け出力（表 / TSV） ------------------------------------------------

_STATUS_LABELS = {
    "ok": "一致",
    "value_mismatch": "値不一致",
    "not_active": "有効行なし",
    "missing": "行なし",
    "unexpected_active": "想定外に有効",
    "unexpected_present": "想定外に存在",
    "file_missing": "ファイルなし",
    "not_in_release": "版に存在せず",
    "renamed": "旧名で存在",
    "uncovered": "シート未記載",
    "uncovered_file": "ファイルごと未記載",
    "review": "要目視",
    "skipped": "対象外",
}

_TABLE_COLUMNS = [
    ("status", "状態", 12),
    ("sheet", "シート", 16),
    ("No", "No", 4),
    ("file", "ファイル", 20),
    ("設定階層", "設定階層", 26),
    ("パラメータ", "パラメータ", 28),
    ("expected", "期待値", 30),
    ("actual", "実際の値/行", 40),
    ("detail", "判定理由", 28),
]


def _display_width(text: str) -> int:
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def _truncate(text: str, width: int) -> str:
    if _display_width(text) <= width:
        return text
    out = ""
    for ch in text:
        if _display_width(out + ch) > width - 1:
            break
        out += ch
    return out + "…"


def _pad(text: str, width: int) -> str:
    return text + " " * max(0, width - _display_width(text))


def _cell(value: object) -> str:
    if value is None:
        return "-"
    return _norm_ws(str(value)) or "-"


def _actual_text(r: dict) -> str:
    if r.get("actual_values"):
        return " | ".join(r["actual_values"])
    lines = r.get("matched_lines") or []
    if lines:
        first = lines[0]
        marker = "" if first["active"] else "(#)"
        return f"L{first['line']}{marker}: {first['text']}"
    return "-"


def _row_fields(r: dict) -> dict[str, str]:
    return {
        "status": _STATUS_LABELS.get(r["status"], r["status"]),
        "sheet": _cell(r.get("sheet")),
        "No": _cell(r.get("No")),
        "file": str(r.get("設定ファイルパス") or "-").split("/")[-1],
        "path": _cell(r.get("設定ファイルパス")),
        "設定階層": _cell(r.get("設定階層")),
        "パラメータ": _cell(r.get("パラメータ")),
        "expected": _cell(r.get("expected_value")),
        "actual": _actual_text(r),
        "detail": _cell(r.get("detail")),
    }


def render_table(results: list[dict], summary: dict) -> str:
    out: list[str] = []
    out.append("== サマリー ==")
    for label, counts in summary.items():
        parts = [f"{_STATUS_LABELS.get(k, k)}={v}" for k, v in sorted(counts.items(), key=lambda kv: -kv[1])]
        out.append(f"[{label}] " + ", ".join(parts))
    out.append("")

    by_archive: dict[str, list[dict]] = {}
    for r in results:
        by_archive.setdefault(r["archive"], []).append(r)

    header = "  ".join(_pad(title, w) for _, title, w in _TABLE_COLUMNS)
    rule = "  ".join("-" * w for _, _, w in _TABLE_COLUMNS)
    for label, rows in by_archive.items():
        out.append(f"== {label} ({len(rows)}件) ==")
        out.append(header)
        out.append(rule)
        for r in rows:
            fields = _row_fields(r)
            out.append("  ".join(_pad(_truncate(fields[key], w), w) for key, _, w in _TABLE_COLUMNS))
        out.append("")
    out.append("(#) はコメントアウト行。列が切れている場合は --format tsv か yaml で全文を確認できます。")
    return "\n".join(out) + "\n"


def render_tsv(results: list[dict]) -> str:
    columns = ["archive", "env", "status", "状態", "sheet", "No", "設定ファイルパス", "設定階層", "パラメータ",
               "期待値", "実際の値/行", "判定理由"]
    lines = ["\t".join(columns)]
    for r in results:
        f = _row_fields(r)
        values = [r["archive"], r["env"], r["status"], _STATUS_LABELS.get(r["status"], ""), f["sheet"], f["No"],
                  f["path"], f["設定階層"], f["パラメータ"], f["expected"], f["actual"], f["detail"]]
        lines.append("\t".join(v.replace("\t", " ") for v in values))
    return "\n".join(lines) + "\n"


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("items", type=Path, help="config-items-yamlの出力(.yaml/.json) または パラメータシート(.xlsx/.xlsm)")
    parser.add_argument("archives", type=Path, nargs="+", help="収集アーカイブ(.tar.gz) または展開済みディレクトリ。名前に rod / ec2 を含める")
    parser.add_argument("-o", "--output", type=Path, default=None, help="レポートの出力先（省略時は標準出力）")
    parser.add_argument("--only-problems", action="store_true", help="status が ok / skipped 以外の結果だけを出力する")
    parser.add_argument(
        "--reference",
        type=Path,
        default=Path("reference/rhel-pkgs"),
        help="RHEL7/8のパッケージ既定設定ファイル置き場（rod=el7, ec2=el8 として、項目がその版に存在しない場合を "
             "not_in_release として区別する）。無ければこの判定はスキップ",
    )
    parser.add_argument(
        "--no-coverage",
        dest="coverage",
        action="store_false",
        help=f"漏れチェックを行わない。既定では {' '.join(COVERAGE_DIRS)} 配下の設定ファイルを走査し、"
             "有効な設定のうちシートに記載が無いものを uncovered として報告する",
    )
    parser.add_argument(
        "--format",
        choices=["yaml", "table", "tsv"],
        default="yaml",
        help="出力形式。yaml=機械処理向け(既定) / table=人間向けの整形表 / tsv=Excel等に貼り付けやすいタブ区切り",
    )


def run(args: argparse.Namespace) -> int:
    if not args.items.exists():
        print(f"エラー: ファイルが見つかりません: {args.items}", file=sys.stderr)
        return 1
    for a in args.archives:
        if not a.exists():
            print(f"エラー: アーカイブが見つかりません: {a}", file=sys.stderr)
            return 1

    try:
        rows = _load_rows(args.items)
        archives = [open_archive(a) for a in args.archives]
    except (ValueError, tarfile.TarError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1

    reference = _Reference(args.reference)
    if not reference.available:
        print(f"注意: 参考パッケージディレクトリが見つからないため、RHELバージョン差の判定は行いません: {args.reference}", file=sys.stderr)

    try:
        results = compare(rows, archives, reference, coverage=args.coverage)
    finally:
        for a in archives:
            if a._tmp is not None:
                a._tmp.cleanup()

    summary = _summarize(results)
    if args.only_problems:
        results = [r for r in results if r["status"] not in ("ok", "skipped")]

    if args.format == "table":
        text = render_table(results, summary)
    elif args.format == "tsv":
        text = render_tsv(results)
    else:
        text = yaml.dump({"summary": summary, "results": results}, allow_unicode=True, sort_keys=False, default_flow_style=False)

    if args.output is not None:
        args.output.write_text(text, encoding="utf-8")
        print(f"突合レポートを書き出しました: {args.output}")
        for label, counts in summary.items():
            print(f"  [{label}] " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    else:
        print(text, end="")
    return 0
