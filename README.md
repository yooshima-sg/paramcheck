# datacheck

パラメータシート(Excel)と、サーバ上の実際の設定ファイルを突き合わせるためのチェックプログラム集。

サーバ移行(RHEL7 → RHEL8)にあたり、「パラメータシートに書かれている設定が、
移行元(ROD)・移行先(EC2)の実機で本当にその通りになっているか」を機械的に検証することを目的とする。

## セットアップ

```bash
uv sync
uv run datacheck --help
```

以降の例では `uv run datacheck ...` を `datacheck ...` と略記する。

## 全体の流れ

```
パラメータシート(.xlsx)
   │
   ├─(1) excel-config-path ─────────→ 設定ファイルパス一覧(JSON)
   │
   ├─(2) generate-collect-script ───→ collect_*.sh ──→ 各サーバで実行
   │                                                      │
   │                                                      ↓
   │                                          収集アーカイブ(.tar.gz)
   │                                          <名前>/collected/<元の絶対パス>
   ├─(3) config-items-yaml ─────────→ 設定項目一覧(YAML)
   │                                                      │
   └─(4) compare-config ←─────────────────────────────────┘
                 │
                 ↓
            突合レポート(YAML / 表 / TSV)
```

## コマンド

### 1. excel-config-path — 設定ファイルパスの抽出

各シートから「設定ファイルパス」項目の値だけを抜き出す。

```bash
datacheck excel-config-path sheets/パラメータシート_Smaph2.xlsx -o paths.json
```

- Excelテーブル(ListObject)の列見出しが「設定ファイルパス」の場合と、通常のラベルセルの直下に
  値が並んでいる場合の両方に対応
- 1つのセルに空白・改行区切りで複数パスが入っている場合は分割して抽出
- 「ー」などダッシュ類のみの値(未設定マーカー)は除外
- `-o` 省略時は標準出力

### 2. generate-collect-script — 収集スクリプトの生成

対象サーバに配置して実行するだけで、設定ファイルを収集する自己完結型シェルスクリプトを生成する。

```bash
datacheck generate-collect-script sheets/パラメータシート_Smaph2.xlsx -o collect_smaph2.sh
```

入力はExcelでも、`excel-config-path` が出力したJSONでもよい。

生成されるスクリプトの特徴:

- 収集対象パスを**重複除去した平文のリスト**としてスクリプト内に直接埋め込む
  （base64等のエンコードはしない。コピー&ペーストでの配置を前提としており、実行前に一覧を目視確認できる）
- 依存は bash と coreutils + tar/gzip のみ。jq や python3 は不要
- RHEL準拠（Windowsパスの考慮は含まない）

サーバ側での実行:

```bash
./collect_smaph2.sh [出力先ディレクトリ]
```

生成物:

| パス | 内容 |
|---|---|
| `<出力先>/collected/<元の絶対パス>` | 収集できたファイル（階層構造を保持） |
| `<出力先>/manifest.tsv` | 元パス・収集結果(collected/not_found/copy_failed)の一覧 |
| `<出力先>/collect.log` | ログ |
| `<出力先>.tar.gz` | 上記一式のアーカイブ（持ち帰り用） |

アーカイブのファイル名には **`rod` または `ec2` を含める**こと（後段の突合でどちらの設定値と
比較するかの判別に使う）。例: `20260921-rod-smaph2.tar`

### 3. config-items-yaml — 設定項目の抽出

「No」「カテゴリ」「設定項目名」「設定ファイルパス」「設定階層」「パラメータ」が揃っている
シート/テーブルの内容をYAMLで出力する。

```bash
datacheck config-items-yaml sheets/パラメータシート_Smaph2.xlsx -o items.yaml
```

- 出力は1行=1レコードのフラットなリスト。各レコードに元のシート名・テーブル名が付く
- 上記6項目が絞り込み条件。「ビルトイン設定」「EC2パッケージ同梱値」「ROD設定値」「EC2設定値」は
  任意項目で、列があれば収集し、無ければ `null`
- 列名「分類」は「カテゴリ」の別名として扱う（vsftpd設定・Postfix_サービス定義シート等）

出力例:

```yaml
- sheet: Apache設定
  table: T11_
  'No': 1
  カテゴリ: 基本
  設定項目名: サーバールートディレクトリ
  設定ファイルパス: /etc/httpd/conf/httpd.conf
  設定階層: global
  パラメータ: ServerRoot
  ビルトイン設定: /etc/httpd
  EC2パッケージ同梱値: /etc/httpd
  ROD設定値: /etc/httpd
  EC2設定値: /etc/httpd
```

### 4. compare-config — 突合

パラメータシートと、収集アーカイブ内の実設定ファイルを突き合わせる。

```bash
# 人間向けの表で、問題のある行だけを見る
datacheck compare-config sheets/パラメータシート_Smaph2.xlsx \
    settingarchives/20260921-rod-smaph2.tar \
    settingarchives/20260921-ec2-smaph2.tar \
    --format table --only-problems

# Excelに貼り付けるためのTSVをファイルに出力
datacheck compare-config items.yaml settingarchives/*.tar --format tsv -o report.tsv
```

第1引数はExcelでも `config-items-yaml` の出力YAMLでもよい。
アーカイブは `.tar` / `.tar.gz` / 展開済みディレクトリを複数指定できる。

| オプション | 説明 |
|---|---|
| `--format yaml\|table\|tsv` | 既定は `yaml`。`table`=整形表、`tsv`=Excel貼り付け用 |
| `--only-problems` | `ok` / `skipped` 以外だけを出力 |
| `--no-coverage` | 漏れチェックを行わない |
| `--reference DIR` | RHEL既定設定の参照先。既定は `reference/rhel-pkgs` |
| `-o FILE` | 出力先（省略時は標準出力） |

## compare-config の判定ルール

### 期待する状態は「ROD設定値 / EC2設定値」の値から決まる

アーカイブ名に `rod` が含まれればROD設定値、`ec2` が含まれればEC2設定値と比較する。

| シートの値 | 期待する実ファイルの状態 |
|---|---|
| `…コメントアウト` | コメントアウト行が存在し、有効行が無い |
| `記載なし` | その項目の行が(コメント含め)一切存在しない |
| `未設定…` / `設定なし` / `※設定項目なし` | 有効行が無い（コメント行は許容） |
| `有効化` / `有効…` | 有効行が存在する（値は比較しない） |
| `無効(…)` / `※` で始まる注記のみ | テキストでは判定不能 → `review`（目視確認） |
| 上記以外の具体値 | 有効行が存在し、値が一致する（末尾の `※…` 注記は比較対象外） |
| 空欄・ダッシュ類 | `skipped` |

値の比較では、引用符・エスケープ用バックスラッシュの有無と大文字小文字の差は無視する
（Apacheの `On`/`on`、`None`/`none`、LogFormat内の `\"` など）。

### 設定階層

- `global` … ファイルのトップレベル
- 入れ子は `親 > 子` の連結表記。例: `<VirtualHost _default_:443> > <Directory "/var/www">`
- 同名ブロックが複数ある場合は `#N` で何番目かを指定。例: `<VirtualHost *:10443> #2`
- コメントアウトされたブロック（`#<Directory …>` 〜 `#</Directory>`）の内側も正しい階層として扱う

### ステータス一覧

| status | 表示 | 意味 |
|---|---|---|
| `ok` | 一致 | 期待どおり |
| `value_mismatch` | 値不一致 | 有効行はあるが値が違う |
| `not_active` | 有効行なし | 有効であるべきだがコメントアウトのみ |
| `unexpected_active` | 想定外に有効 | コメントアウトのはずが有効になっている |
| `unexpected_present` | 想定外に存在 | 無い/未設定のはずが行がある |
| `missing` | 行なし | 期待する行が見つからない |
| `file_missing` | ファイルなし | アーカイブにそのファイルが無い |
| `not_in_release` | 版に存在せず | そのRHEL版のパッケージ既定に項目/ファイル/ブロックが存在しない |
| `renamed` | 旧名で存在 | 設定名が版間で変わっており、旧名で見つかった |
| `uncovered` | シート未記載 | 実機に有効な設定があるがシートに記載が無い（漏れ） |
| `uncovered_file` | ファイルごと未記載 | シートが全く言及していない設定ファイル |
| `review` | 要目視 | 自動判定できないため目視確認が必要 |
| `skipped` | 対象外 | 比較対象外の行 |

### RHELバージョン差の考慮

`--reference`（既定 `reference/rhel-pkgs`）配下のRHEL7/8パッケージ既定設定ファイルを参照し、
`rod`=el7 / `ec2`=el8 として、単なる「行なし」と「そのRHEL版にそもそも存在しない」を区別する。

- 項目・ファイル・設定階層のブロックが、その版の既定に無く、もう一方の版にはある → `not_in_release`
- シートの `※旧設定名:ssl_protocols <!SSLv2>` 注記、または既知の改名表で旧名の行が見つかる → `renamed`
  （既知の改名: Dovecot 2.2→2.3 の `ssl_protocols`→`ssl_min_protocol`、`ssl_dh_parameters_length`→`ssl_dh`）

参照ディレクトリが無い場合はこの判定をスキップする（警告を表示）。詳細は
[reference/rhel-pkgs/README.md](reference/rhel-pkgs/README.md) を参照。

### 漏れチェック

既定で `/etc/httpd` `/etc/postfix` `/etc/dovecot` `/etc/vsftpd` 配下の設定ファイルを走査し、
**実機に有効な設定があるのにシートに記載が無いもの**を洗い出す。

- 対象ファイルは `*.conf` `*.conf.ext` `main.cf` `master.cf` のみ
  （access/canonical等のルックアップテーブル、ftpusers、magic、証明書、スクリプトは除外）
- コメントアウト行は対象外（実際に効いている設定のみを漏れとして扱う）
- シートが1行も言及していないファイルは、個々の設定を列挙せず `uncovered_file` 1行にまとめる

## 対応する設定ファイル形式

| パス | 形式 |
|---|---|
| `/etc/httpd/**` | Apache httpd（`<Tag>` セクション、`\` による行継続、大文字小文字非区別） |
| `/etc/postfix/main.cf` | `key = value`（空白インデントによる継続行に対応） |
| `/etc/postfix/master.cf` | 列形式（`service type private unpriv chroot wakeup maxproc command`）。`-o key=value` 継続行はサービス行に畳み込む |
| `/etc/dovecot/**` | `key = value`、`name [引数] {` ブロック、`!include` |
| `/etc/vsftpd/**` | `key=value` |
| その他 | 先頭トークンを名前とする汎用パーサ |

構文は各公式ドキュメントに基づく:
[Apache](https://httpd.apache.org/docs/2.4/configuring.html) /
[Dovecot](https://doc.dovecot.org/2.3/configuration_manual/config_file/) /
[Postfix](https://www.postfix.org/postconf.5.html)

## ディレクトリ構成

```
src/datacheck/
  __init__.py          CLI（サブコマンドの組み立て）
  __main__.py          python -m datacheck 用
  checks/
    __init__.py        チェックの登録（CHECKS リスト）
    excel_config_path.py
    generate_collector.py
    config_items_yaml.py
    compare_config.py
sheets/                パラメータシート(実データ、gitignore)
settingarchives/       収集アーカイブ(gitignore)
reference/rhel-pkgs/   RHEL7/8のパッケージ既定設定ファイル（RPM本体はgitignore）
```

### チェックプログラムの追加方法

`src/datacheck/checks/` に以下を備えたモジュールを追加し、`checks/__init__.py` の
`CHECKS` リストに加えるだけでよい（CLIへの配線は自動）。

```python
NAME = "コマンド名"
DESCRIPTION = "説明"

def add_arguments(parser: argparse.ArgumentParser) -> None: ...
def run(args: argparse.Namespace) -> int: ...
```
