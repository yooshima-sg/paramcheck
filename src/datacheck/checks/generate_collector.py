"""チェック: excel-config-pathの抽出結果から、自己完結型のファイル収集シェルスクリプトを生成する。

生成されるシェルスクリプトは、コピー＆ペーストでサーバに配置することを前提としている。
抽出された「設定ファイルパス」を重複除去したうえで、素のテキスト（ヒアドキュメント）として
スクリプト本体にそのまま埋め込む。base64化等のエンコードは行わない。
  - 実行前に人間が一覧を目視で確認・修正できる
  - コピー＆ペースト時に文字化けやパディング崩れで壊れる心配がない（base64のような
    エンコード済みブロックと違い、1行1パスなので万一一部が欠けても被害は該当行に限られる）
収集後は出力ディレクトリ全体（collected/ の階層構造・manifest.tsv・collect.log）を
<output_dir>.tar.gz にアーカイブし、サーバから持ち帰る成果物を1ファイルにまとめる。
実行環境にjqやpython3が無くても、bash + coreutils（cp, mkdir, dirname等）+ tar/gzip だけで動作する。
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from . import excel_config_path

NAME = "generate-collect-script"
DESCRIPTION = "抽出結果から、コピペでサーバに配置してファイル収集できる自己完結型シェルスクリプトを生成する"

_HEREDOC_DELIMITER = "__DATACHECK_PATH_LIST__"

_SCRIPT_TEMPLATE = r"""#!/usr/bin/env bash
#
# Usage:
#   ./__SCRIPT_NAME__ [output_dir]
#   Produces <output_dir>/ (collected/, manifest.tsv, collect.log) and <output_dir>.tar.gz

set -uo pipefail

OUTPUT_DIR="${1:-./collected_$(hostname -s 2>/dev/null || hostname)_$(date +%Y%m%d_%H%M%S)}"

mkdir -p "$OUTPUT_DIR/collected"
MANIFEST_FILE="$OUTPUT_DIR/manifest.tsv"
LOG_FILE="$OUTPUT_DIR/collect.log"
: > "$LOG_FILE"
printf 'original_path\tstatus\tcollected_path\n' > "$MANIFEST_FILE"

total=0
collected=0
missing=0
failed=0

while IFS= read -r path; do
    [[ -z "$path" ]] && continue
    total=$((total + 1))

    if [[ -e "$path" ]]; then
        rel="${path#/}"
        dest="$OUTPUT_DIR/collected/$rel"
        mkdir -p "$(dirname "$dest")"
        if cp -p -- "$path" "$dest" 2>>"$LOG_FILE"; then
            collected=$((collected + 1))
            printf '%s\tcollected\t%s\n' "$path" "collected/$rel" >> "$MANIFEST_FILE"
        else
            failed=$((failed + 1))
            printf '%s\tcopy_failed\t\n' "$path" >> "$MANIFEST_FILE"
        fi
    else
        missing=$((missing + 1))
        printf '%s\tnot_found\t\n' "$path" >> "$MANIFEST_FILE"
    fi
done <<'__DATACHECK_PATH_LIST__'
__PATH_LIST__
__DATACHECK_PATH_LIST__

{
    echo "Total paths (deduped): $total"
    echo "Collected            : $collected"
    echo "Not found            : $missing"
    echo "Copy failed          : $failed"
    echo "Output dir           : $OUTPUT_DIR"
} | tee -a "$LOG_FILE"

PARENT_DIR="$(dirname "$OUTPUT_DIR")"
BASE_NAME="$(basename "$OUTPUT_DIR")"
ARCHIVE="$PARENT_DIR/$BASE_NAME.tar.gz"
if tar -czf "$ARCHIVE" -C "$PARENT_DIR" "$BASE_NAME" 2>>"$LOG_FILE"; then
    echo "Archive              : $ARCHIVE" | tee -a "$LOG_FILE"
else
    echo "Archive              : FAILED (see $LOG_FILE)" | tee -a "$LOG_FILE" >&2
    exit 1
fi
"""


def _collect_unique_paths(extracted: dict[str, list[dict]]) -> list[str]:
    """全シートの抽出結果から値を集め、空白除去・重複除去して並べ替えたものを返す。"""
    paths: set[str] = set()
    for matches in extracted.values():
        for m in matches:
            value = m.get("value")
            if value is None:
                continue
            text = str(value).strip()
            if not text:
                continue
            text = text.replace("\t", " ").replace("\r", " ").replace("\n", " ")
            paths.add(text)
    return sorted(paths)


def generate_script(extracted: dict[str, list[dict]], source_label: str, script_name: str) -> str:
    paths = _collect_unique_paths(extracted)

    if any(p == _HEREDOC_DELIMITER for p in paths):
        raise ValueError(
            f"抽出結果にヒアドキュメント終端記号と同じ値が含まれているため生成できません: {_HEREDOC_DELIMITER}"
        )

    script = _SCRIPT_TEMPLATE
    script = script.replace("__SOURCE__", source_label)
    script = script.replace("__GENERATED_AT__", datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"))
    script = script.replace("__ENTRY_COUNT__", str(len(paths)))
    script = script.replace("__SCRIPT_NAME__", script_name)
    script = script.replace("__PATH_LIST__", "\n".join(paths))
    return script


def _load_extracted(input_path: Path) -> tuple[dict[str, list[dict]], str]:
    suffix = input_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return excel_config_path.extract_workbook(input_path), f"Excel: {input_path.name}"
    if suffix == ".json":
        with input_path.open(encoding="utf-8") as f:
            data = json.load(f)
        return data, f"JSON: {input_path.name}"
    raise ValueError(f"未対応の拡張子です(.xlsx/.xlsm/.jsonのいずれかを指定してください): {input_path}")


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "input_path",
        type=Path,
        help="入力元。Excelファイル(.xlsx/.xlsm)、または datacheck excel-config-path の出力JSON(.json)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="生成するシェルスクリプトのパス（省略時: collect_<入力ファイル名>.sh）",
    )


def run(args: argparse.Namespace) -> int:
    if not args.input_path.exists():
        print(f"エラー: ファイルが見つかりません: {args.input_path}", file=sys.stderr)
        return 1

    try:
        extracted, source_label = _load_extracted(args.input_path)
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1

    paths = _collect_unique_paths(extracted)
    if not paths:
        print("警告: 抽出対象のパスが1件もありません。空の収集スクリプトを生成します。", file=sys.stderr)

    output_path = args.output or Path(f"collect_{args.input_path.stem}.sh")
    try:
        script_text = generate_script(extracted, source_label, output_path.name)
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    output_path.write_text(script_text, encoding="utf-8", newline="\n")
    output_path.chmod(0o755)

    print(f"収集スクリプトを生成しました: {output_path}（重複除去後のパス数: {len(paths)}）")
    return 0
