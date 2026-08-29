#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
既存の data/ 以下にある games/json と season フォルダを
docs/data/ 以下へ移動する1回限りの移行スクリプト。
raw/ と games/datamart/ はそのまま data/ に残します（非公開のバックアップ用）。

使い方:
  プロジェクトルート（scripts/ と data/ がある場所）で実行してください。
  python3 migrate_to_docs.py
"""
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC_ROOTS = [ROOT / "data" / "プロ野球", ROOT / "data" / "MLB"]

moved = 0
for src_root in SRC_ROOTS:
    if not src_root.exists():
        continue
    for games_json_dir in src_root.rglob("games/json"):
        rel = games_json_dir.relative_to(ROOT / "data")
        dest = ROOT / "docs" / "data" / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        print(f"移動: {games_json_dir} -> {dest}")
        shutil.move(str(games_json_dir), str(dest))
        moved += 1
    for season_dir in src_root.rglob("season"):
        rel = season_dir.relative_to(ROOT / "data")
        dest = ROOT / "docs" / "data" / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        print(f"移動: {season_dir} -> {dest}")
        shutil.move(str(season_dir), str(dest))
        moved += 1

print(f"\n完了: {moved} 件のフォルダを docs/data/ 配下へ移動しました。")
print("dashboard.html（改修版）を docs/index.html としてコピーするのを忘れずに。")
