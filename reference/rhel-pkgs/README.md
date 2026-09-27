# RHEL7 / RHEL8 参考パッケージ

RHEL7とRHEL8で、httpd/mod_ssl/mod_http2/postfix/dovecot/rsyslogの設定ファイルが
どう変わるか（存在の有無・ディレクティブの追加/廃止/名称変更/デフォルト値変更）を
確認するために取得した参考データ。

RHELは購読が無いと直接取得できないため、同一SRPMから再ビルドされたバイナリ互換
ディストリを代わりに使用（%filesリスト・設定ファイルの中身はRHELと同一）。

- RHEL7相当: CentOS 7.9.2009 (vault.centos.org) — RHEL 7.9時点の内容。RHEL7系の
  これらのパッケージは主要バージョンが上がらない限り中身は変わらないため、
  7.2などの古いマイナーバージョンでも実質的に同じと考えてよい。
- RHEL8.10相当: Rocky Linux 8 の `el8_10` タグ付きビルド (dl.rockylinux.org) —
  RHEL 8.10と厳密に対応するビルドを選定済み。

取得日: 2026-09-21

## 取得パッケージ

| パッケージ | el7 (CentOS 7.9.2009) | el8.10 (Rocky 8) |
|---|---|---|
| httpd | 2.4.6-95.el7.centos | 2.4.37-65.module+el8.10.0+40312+2c72bb9b.10 |
| mod_ssl | 2.4.6-95.el7.centos | 2.4.37-65.module+el8.10.0+40312+2c72bb9b.10 |
| mod_http2 | (パッケージ自体が存在しない) | 1.15.7-10.module+el8.10.0+40257+286895ef.7 |
| postfix | 2.10.1-9.el7 | 3.5.8-8.el8_10 |
| dovecot | 2.2.36-8.el7 | 2.3.16-8.el8_10 |
| vsftpd | 3.0.2-28.el7 | 3.0.3-36.el8_10.3 |
| rsyslog | 8.24.0-55.el7 | 8.2102.0-15.el8_10.1 |

## 構成

```
el7/rpms/                     ダウンロードしたRPM本体
el7/configs/etc/httpd/...     RPMから展開した設定ファイル(実際の絶対パス構造を保持)
el7/configs/etc/postfix/...
el7/configs/etc/dovecot/...
el7/configs/etc/vsftpd/...
el8/rpms/
el8/configs/etc/...
```

展開対象は /etc/httpd /etc/postfix /etc/dovecot /etc/vsftpd 配下の全ファイル。
`compare-config` は `configs/<絶対パス>` で直接引くため、この構造を崩さないこと。

el8 の `/etc/httpd/conf.d/rocky-snipolicy.conf` は Rocky Linux 固有で、RHEL 8 には存在しない。

## 取得・展開方法

システムに`rpm2cpio`/`cpio`/`dnf`が無い環境のため、純Python製のRPMリーダーを使用:

```bash
uv run --with rpmfile --with zstandard python3 -c "
import rpmfile
with rpmfile.open('el7/rpms/httpd-....rpm') as rpm:
    for m in rpm.getmembers():
        print(m.name)
"
```

## compare-config からの利用

`datacheck compare-config` は既定でこのディレクトリ（`--reference` で変更可）を参照し、
rod=el7 / ec2=el8 として「行が見つからない」結果を再分類する:

- その版の既定ファイルにも項目/ファイル/ブロックが無く、もう一方の版にはある → `not_in_release`（版に存在せず）
- シートの `※旧設定名:xxx <値>` 注記、または `_KNOWN_RENAMES` により旧名で見つかる → `renamed`（旧名で存在）

## 用途・注意

- あくまで「RHELパッケージが出荷時にデフォルトで作る設定ファイルの中身」の参考データ。
  実際にs-giken案件で使うExcelパラメータシート(`sheets/`, gitignore対象・非公開)の
  実データとは別物。
- ディレクティブが「このバージョンの既定ファイルに記載が無い」からといって、必ずしも
  「そのバージョンで使えない」とは限らない（単にテンプレートが簡素なだけの場合がある）。
  postfixのel7 main.cfはTLS関連の記述が一切無いが、TLS関連パラメータ自体は
  Postfix 2.3〜2.6時代から存在する。個別に裏取りすること。
- このディレクトリはgit管理対象外(`.gitignore`)。再取得したい場合は上表のURLパターンを
  vault.centos.org / dl.rockylinux.org に対して使えば同じものが手に入る。
