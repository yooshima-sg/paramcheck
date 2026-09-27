"""データチェックプログラム集の登録。

新しいチェックを追加する場合は、`NAME` / `DESCRIPTION` / `add_arguments(parser)` /
`run(args) -> int` を備えたモジュールをこのパッケージに追加し、下の `CHECKS` に加える。
"""

from . import compare_config, config_items_yaml, excel_config_path, generate_collector

CHECKS = [
    excel_config_path,
    generate_collector,
    config_items_yaml,
    compare_config,
]
