# -*- coding: utf-8 -*-
"""
export_llm_input.py
====================
index.html 側でクライアントサイド（JavaScript）に実装されている、投手の
シーズン集計・球種別詳細・コース分布・カウント別パターンの集計ロジックを
Pythonに移植し、Claudeへ渡すLLM入力用xlsx（縦持ち・複数シート）を生成する。

ポート元:
  - calcSeasonStats()      → シーズン投手成績（防御率・K-BB%など）
  - _aggregateSeasonMix()  → 球種別シーズン集計（cbsのマージを含む）
  - 9分割コース分布ロジック → コース分布（対右/対左、ゾーン内9マス）

入力:
  run.py の run_dashboard() が既に生成している「日別ダッシュボードJSON」
  （games/json/{date}.json）を、対象期間ぶんすべて読み込んで使う。
  players/json/{選手名}.json（MLB専用のキャッシュ）は使わない
  → NPBはこの日別JSONの積み上げだけでシーズン集計が可能なため。

出力:
  1つのxlsxに以下のシートを縦持ち（long形式）で書き出す。
    - シーズン集計       : 1行 = 1投手
    - 球種別詳細         : 1行 = 1投手 × 1球種
    - コース分布         : 1行 = 1投手 × 対戦打者(右/左) × ゾーン(1-9)
    - カウント別パターン : 1行 = 1投手 × 球種 × カウント状況

使い方:
    python export_llm_input.py \
        --games-json-dir "docs/data/プロ野球/2026年/1軍/レギュラーシーズン/games/json" \
        --out "data/datamart/llm_input/プロ野球_2026_投手データ.xlsx" \
        --min-ip 10

    # 投球回フィルタなしで全投手を対象にする場合は --min-ip を省略（デフォルト0）
"""

from __future__ import annotations

import argparse
import glob
import json
import os
from collections import defaultdict

import pandas as pd


# ==================================================
# Section 1. 日別JSONの読み込み・appearances構築
#   index.html の _doOpenPlayerPage() 相当
# ==================================================

def load_daily_games(games_json_dir: str, verbose: bool = True) -> dict:
    """games/json/{date}.json を全部読み込み、{date: [game, ...]} にまとめる。
    壊れた/想定外の形式のファイルやゲームエントリはスキップし、警告を出して処理を継続する。
    """
    all_data: dict[str, list] = {}
    paths = sorted(glob.glob(os.path.join(games_json_dir, "*.json")))
    if not paths:
        raise FileNotFoundError(f"games json が見つかりません: {games_json_dir}")

    def _add_games(date_key: str, games) -> None:
        if not isinstance(games, list):
            if verbose:
                print(f"  [SKIP] {date_key}: gamesがlistではない型({type(games).__name__})のためスキップ")
            return
        valid_games = [g for g in games if isinstance(g, dict)]
        skipped = len(games) - len(valid_games)
        if skipped > 0 and verbose:
            print(f"  [SKIP] {date_key}: dict以外のgameエントリを{skipped}件スキップ")
        all_data.setdefault(date_key, []).extend(valid_games)

    for path in paths:
        date = os.path.splitext(os.path.basename(path))[0]
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            if verbose:
                print(f"  [SKIP] {os.path.basename(path)}: 読み込み失敗({e})")
            continue

        if not isinstance(payload, dict):
            if verbose:
                print(f"  [SKIP] {os.path.basename(path)}: 想定外のトップレベル型({type(payload).__name__})")
            continue

        if date in payload:
            _add_games(date, payload[date])
        else:
            for k, v in payload.items():
                if k == "highlights" or (isinstance(k, str) and k.startswith("_")):
                    continue
                _add_games(k, v)
    return all_data


def build_all_pitcher_names(all_data: dict) -> set[str]:
    """全登場投手名を収集（壊れたエントリはスキップ）"""
    names = set()
    for date, games in all_data.items():
        if date == "highlights" or date.startswith("_"):
            continue
        for g in games:
            if not isinstance(g, dict):
                continue
            pitchers = g.get("pitchers")
            if not isinstance(pitchers, dict):
                continue
            for side in ("home", "away"):
                for p in (pitchers.get(side) or []):
                    if isinstance(p, dict) and p.get("name"):
                        names.add(p["name"])
    return names


def build_appearances(all_data: dict, player_name: str) -> list[dict]:
    """特定投手の全登板 = appearances（index.htmlのappearances配列と同じ形）。壊れたエントリはスキップ"""
    appearances = []
    dates = sorted(d for d in all_data.keys() if d != "highlights" and not d.startswith("_"))
    name_lower = player_name.lower()
    for date in dates:
        for g in all_data.get(date, []):
            if not isinstance(g, dict):
                continue
            pitchers = g.get("pitchers")
            if not isinstance(pitchers, dict):
                continue
            for side in ("home", "away"):
                for p in (pitchers.get(side) or []):
                    if not isinstance(p, dict):
                        continue
                    if (p.get("name") or "").lower() == name_lower:
                        appearances.append({"date": date, "game": g, "side": side, "player": p})
    return appearances


# ==================================================
# Section 2. シーズン集計（calcSeasonStats のポート）
# ==================================================

def _ip_to_outs(ip_str) -> int:
    """'6.2' 形式（6回2/3）→ アウト数（6*3+2=20）"""
    s = str(ip_str or "0")
    whole, _, frac = s.partition(".")
    whole_n = int(whole) if whole.strip().lstrip("-").isdigit() else 0
    frac_n = 1 if frac == "1" else 2 if frac == "2" else 0
    return whole_n * 3 + frac_n


def _outs_to_ip_str(outs: int) -> str:
    whole = outs // 3
    rem = outs % 3
    return f"{whole}.{rem}"


def calc_season_stats(appearances: list[dict]) -> dict:
    """calcSeasonStats(mode='all') のポート。防御率・K-BB%などシーズン集計を1行返す"""
    appearances = [ap for ap in appearances if isinstance(ap.get("player"), dict)]
    if not appearances:
        return {"選手名": None, "登板数": 0, "投球回": "0.0", "奪三振": 0, "与四球": 0,
                "被安打": 0, "自責点": 0, "防御率": None, "K%": None, "BB%": None, "K-BB%": None,
                "空振り率": None, "ゾーン外スイング率": None, "ストライク率": None, "ゾーン率": None, "ゴロ率": None}

    total_ip_raw = sum(_ip_to_outs(ap["player"].get("ip")) for ap in appearances)
    total_k  = sum(ap["player"].get("k", 0) or 0 for ap in appearances)
    total_bb = sum(ap["player"].get("bb", 0) or 0 for ap in appearances)
    total_h  = sum(ap["player"].get("h", 0) or 0 for ap in appearances)
    total_er = sum(ap["player"].get("er", 0) or 0 for ap in appearances)

    ip_num = total_ip_raw / 3
    era = round(total_er / ip_num * 9, 2) if ip_num > 0 else None

    total_tbf = sum(
        (ap["player"].get("tbf") or (ap["player"].get("k", 0) or 0)
         + (ap["player"].get("bb", 0) or 0) + (ap["player"].get("h", 0) or 0))
        for ap in appearances
    )
    kpct   = round(total_k / total_tbf * 100, 1) if total_tbf > 0 else None
    bbpct  = round(total_bb / total_tbf * 100, 1) if total_tbf > 0 else None
    kbbpct = round((total_k - total_bb) / total_tbf * 100, 1) if total_tbf > 0 else None

    tot_pitches = swstr_sum = oswing_sum = strike_sum = zone_sum = 0.0
    gb_sum, gb_count = 0.0, 0
    for ap in appearances:
        p = ap["player"]
        n = p.get("pitches", 0) or 0
        if n > 0:
            tot_pitches += n
            if p.get("swstr")  is not None: swstr_sum  += p["swstr"]  * n
            if p.get("oSwing") is not None: oswing_sum += p["oSwing"] * n
            if p.get("strike") is not None: strike_sum += p["strike"] * n
            if p.get("zone")   is not None: zone_sum   += p["zone"]   * n
        if p.get("gbpct") is not None:
            gb_sum += p["gbpct"]
            gb_count += 1

    return {
        "選手名":   appearances[0]["player"]["name"] if appearances else None,
        "登板数":   len(appearances),
        "投球回":   _outs_to_ip_str(total_ip_raw),
        "奪三振":   total_k,
        "与四球":   total_bb,
        "被安打":   total_h,
        "自責点":   total_er,
        "防御率":   era,
        "K%":      kpct,
        "BB%":     bbpct,
        "K-BB%":   kbbpct,
        "空振り率":  round(swstr_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゾーン外スイング率": round(oswing_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ストライク率": round(strike_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゾーン率":  round(zone_sum / tot_pitches, 1) if tot_pitches > 0 else None,
        "ゴロ率":   round(gb_sum / gb_count, 1) if gb_count > 0 else None,
    }


# ==================================================
# Section 3. 球種別シーズン集計（_aggregateSeasonMix のポート）
# ==================================================

_CBS_FIELDS = ["c", "sw", "fo", "lo", "ba", "ou", "hi"]


def aggregate_season_mix(appearances: list[dict], mix_key: str = "mix") -> list[dict]:
    """
    mix_key: 'mix'（対戦打者を問わない合算） / 'mixVsR' / 'mixVsL'
    球種ごとに投球数加重平均で指標をまとめ、cbs（カウント別集計）もマージする。
    """
    km: dict[str, dict] = {}
    for ap in appearances:
        player = ap.get("player")
        if not isinstance(player, dict):
            continue
        if mix_key == "mixVsR":
            src = player.get("mixVsR") or player.get("mix") or []
        elif mix_key == "mixVsL":
            src = player.get("mixVsL") or player.get("mix") or []
        else:
            src = player.get("mix") or []
        if not isinstance(src, list):
            continue

        for m in src:
            if not isinstance(m, dict):
                continue
            key = m.get("key")
            if key not in km:
                km[key] = {
                    "name": m.get("name"), "key": key, "count": 0,
                    "swstr_sum": 0.0, "zone_sum": 0.0,
                    "oswing_sum": 0.0, "oswing_cnt": 0,
                    "vel_sum": 0.0, "vel_cnt": 0, "max_vel": 0.0,
                    "hits": 0, "hr": 0,
                    "strike_sum": 0.0, "strike_cnt": 0,
                    "gb_sum": 0.0, "gb_cnt": 0,
                    "cbs_merged": defaultdict(lambda: {f: 0 for f in _CBS_FIELDS}),
                }
            k = km[key]
            count = m.get("count", 0) or 0
            k["count"] += count
            k["swstr_sum"] += (m.get("swstr") or 0) * count
            k["zone_sum"]  += (m.get("zone")  or 0) * count
            if m.get("oSwing") is not None:
                k["oswing_sum"] += (m["oSwing"] or 0) * count
                k["oswing_cnt"] += count
            if m.get("vel"):
                k["vel_sum"] += m["vel"] * count
                k["vel_cnt"] += count
            if m.get("maxVel") and m["maxVel"] > k["max_vel"]:
                k["max_vel"] = m["maxVel"]
            k["hits"] += m.get("hits", 0) or 0
            k["hr"]   += m.get("hr", 0) or 0
            if m.get("strike") is not None:
                k["strike_sum"] += (m["strike"] or 0) * count
                k["strike_cnt"] += count
            if m.get("gbpct") is not None:
                k["gb_sum"] += m["gbpct"] or 0
                k["gb_cnt"] += 1
            # cbs（カウント別）マージ
            cbs = m.get("cbs") or {}
            if isinstance(cbs, dict):
                for count_key, cv in cbs.items():
                    if not isinstance(cv, dict):
                        continue
                    cm = k["cbs_merged"][count_key]
                    for f in _CBS_FIELDS:
                        cm[f] += cv.get(f, 0) or 0

    merged = sorted(km.values(), key=lambda x: -x["count"])
    total = sum(m["count"] for m in merged)

    out = []
    for m in merged:
        out.append({
            "球種名": m["name"],
            "球種コード": m["key"],
            "投球数": m["count"],
            "投球割合%": round(m["count"] / total * 100, 1) if total > 0 else 0.0,
            "平均球速": round(m["vel_sum"] / m["vel_cnt"], 1) if m["vel_cnt"] > 0 else None,
            "最高球速": m["max_vel"] or None,
            "空振り率": round(m["swstr_sum"] / m["count"], 1) if m["count"] > 0 else 0.0,
            "ゾーン率": round(m["zone_sum"] / m["count"], 1) if m["count"] > 0 else 0.0,
            "ゾーン外スイング率": round(m["oswing_sum"] / m["oswing_cnt"], 1) if m["oswing_cnt"] > 0 else None,
            "ストライク率": round(m["strike_sum"] / m["strike_cnt"], 1) if m["strike_cnt"] > 0 else None,
            "GB%": round(m["gb_sum"] / m["gb_cnt"], 1) if m["gb_cnt"] > 0 else None,
            "H": m["hits"],
            "HR": m["hr"],
            "_cbs": dict(m["cbs_merged"]),  # カウント別パターンシート生成用（出力シートには含めない）
        })
    return out


# ==================================================
# Section 4. コース分布（9分割ロジックのポート）
# ==================================================

_ZONE = 1.0
_EDGES9 = [-_ZONE, -_ZONE / 3, _ZONE / 3, _ZONE]


def _get_cell(v: float) -> int:
    for i in range(len(_EDGES9) - 1):
        if _EDGES9[i] <= v < _EDGES9[i + 1]:
            return i
    return 0 if v < _EDGES9[0] else len(_EDGES9) - 2


def aggregate_course_distribution(appearances: list[dict]) -> dict:
    """
    対右打者(R) / 対左打者(L) それぞれについて、ゾーン内投球を3x3(9マス)に分類し
    件数・割合を返す。index.htmlの9分割ロジック（投手視点、ゾーン内のみ分母）と同じ。
    戻り値: {"R": {"total_in_zone": n, "cells": [[cnt,...]x3]}, "L": {...}}
    """
    result = {"R": [[0, 0, 0] for _ in range(3)], "L": [[0, 0, 0] for _ in range(3)]}
    totals = {"R": 0, "L": 0}

    for ap in appearances:
        player = ap.get("player")
        if not isinstance(player, dict):
            continue
        for side_key, mix_key in (("R", "mixVsR"), ("L", "mixVsL")):
            for m in (player.get(mix_key) or []):
                if not isinstance(m, dict):
                    continue
                for loc in (m.get("locs") or []):
                    # loc: [x, y, resultCode, inZoneFlag, ...]
                    if not isinstance(loc, list) or len(loc) < 4 or loc[3] != 1:
                        continue
                    x, y = loc[0], loc[1]
                    col = _get_cell(x)
                    row = _get_cell(-y)
                    result[side_key][row][col] += 1
                    totals[side_key] += 1

    out = {}
    for side_key in ("R", "L"):
        cells = result[side_key]
        total = totals[side_key]
        rows = []
        for r in range(3):
            for c in range(3):
                cnt = cells[r][c]
                rows.append({
                    "ゾーン番号": r * 3 + c + 1,  # 1〜9（左上→右下）
                    "球数": cnt,
                    "割合%": round(cnt / total * 100, 1) if total > 0 else 0.0,
                })
        out[side_key] = {"total_in_zone": total, "rows": rows}
    return out


# ==================================================
# Section 5. カウント別パターン（cbsマージ結果を整形）
# ==================================================

def build_count_pattern_rows(player_name: str, season_mix: list[dict]) -> list[dict]:
    """season_mix の各球種が持つ _cbs を、カウント別パターンシート用の行に変換"""
    rows = []
    for m in season_mix:
        for count_key, cv in (m.get("_cbs") or {}).items():
            c = cv.get("c", 0) or 0
            if c == 0:
                continue
            rows.append({
                "選手名": player_name,
                "球種名": m["球種名"],
                "カウント": count_key,  # 例: "0-0", "0-2", "1-2" など
                "球数": c,
                "空振り": cv.get("sw", 0),
                "ファウル": cv.get("fo", 0),
                "見逃し": cv.get("lo", 0),
                "ボール": cv.get("ba", 0),
                "凡打": cv.get("ou", 0),
                "被安打": cv.get("hi", 0),
                "空振り率%": round(cv.get("sw", 0) / c * 100, 1),
                "ファウル率%": round(cv.get("fo", 0) / c * 100, 1),
            })
    return rows


# ==================================================
# Section 6. メイン: 全投手を対象に4シート分のDataFrameを組み立ててxlsx出力
# ==================================================

def export_llm_input_xlsx(games_json_dir: str, out_path: str, min_ip: float = 0.0,
                           target_names: list[str] | None = None) -> str:
    all_data = load_daily_games(games_json_dir)
    names = set(target_names) if target_names else build_all_pitcher_names(all_data)

    season_rows, mix_rows, course_rows, count_rows = [], [], [], []

    for name in sorted(names):
        try:
            appearances = build_appearances(all_data, name)
            if not appearances:
                continue

            season = calc_season_stats(appearances)

            # 投球回フィルタ（例: 10回以上登板した投手のみ対象）
            ip_num = _ip_to_outs(season["投球回"]) / 3
            if ip_num < min_ip:
                continue

            season_rows.append(season)

            season_mix_all = aggregate_season_mix(appearances, "mix")
            for m in season_mix_all:
                row = {"選手名": name, **{k: v for k, v in m.items() if k != "_cbs"}}
                mix_rows.append(row)

            count_rows.extend(build_count_pattern_rows(name, season_mix_all))

            course = aggregate_course_distribution(appearances)
            for side_key, side_label in (("R", "対右打者"), ("L", "対左打者")):
                for row in course[side_key]["rows"]:
                    course_rows.append({
                        "選手名": name,
                        "対戦打者": side_label,
                        "ゾーン内総数": course[side_key]["total_in_zone"],
                        **row,
                    })
        except Exception as e:
            print(f"  [SKIP] {name}: 集計中にエラーのためスキップ({type(e).__name__}: {e})")
            continue

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame(season_rows).to_excel(writer, sheet_name="シーズン集計", index=False)
        pd.DataFrame(mix_rows).to_excel(writer, sheet_name="球種別詳細", index=False)
        pd.DataFrame(course_rows).to_excel(writer, sheet_name="コース分布", index=False)
        pd.DataFrame(count_rows).to_excel(writer, sheet_name="カウント別パターン", index=False)

    return out_path


# ==================================================
# Section 7. CLI
# ==================================================

def parse_args():
    p = argparse.ArgumentParser(description="LLM入力用xlsx（シーズン投手データ）を生成する")
    p.add_argument("--games-json-dir", required=True, help="games/json/{date}.json が入っているディレクトリ")
    p.add_argument("--out", required=True, help="出力xlsxのパス")
    p.add_argument("--min-ip", type=float, default=0.0, help="この投球回以上の投手のみ対象にする（デフォルト0=全員）")
    p.add_argument("--players", nargs="*", default=None, help="対象選手名を絞り込む場合はスペース区切りで指定（省略時は全投手）")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    out = export_llm_input_xlsx(
        games_json_dir=args.games_json_dir,
        out_path=args.out,
        min_ip=args.min_ip,
        target_names=args.players,
    )
    print(f"✅ 出力完了: {out}")
