"""
HOGY社方式との差異調査 — 8試行の稼働率比較
============================================
弊社（メドライン）計算: 76.0%
HOGY社計算: 67.9%
差: 8.1ポイント

各試行で異なる条件の稼働率を算出し、67.9%に最も近い組み合わせを探索する。
"""

import openpyxl
import datetime as dt
import os
import sys

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "時間帯別稼働推移元データ.xlsx")


def to_minutes(t):
    if isinstance(t, dt.time):
        return t.hour * 60 + t.minute
    elif isinstance(t, dt.timedelta):
        return int(t.total_seconds()) // 60
    elif isinstance(t, str):
        parts = t.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    else:
        return t.hour * 60 + t.minute


def main():
    print(f"入力ファイル: {INPUT_FILE}")
    wb = openpyxl.load_workbook(INPUT_FILE)

    # --- 定義シートから設定読み込み ---
    ws_def = wb["定義"]
    room_weight = {}
    for row in ws_def.iter_rows(min_row=2, max_row=20, min_col=1, max_col=2, values_only=True):
        if row[0] is not None and row[1] is not None:
            try:
                room_weight[str(row[0])] = float(row[1])
            except (ValueError, TypeError):
                break
    print(f"対象手術室: {room_weight}")
    print(f"ウェイト合計: {sum(room_weight.values())}")

    exclude_weekdays = set()
    r = 14
    while True:
        r += 1
        v = ws_def.cell(row=r, column=1).value
        if v is None or v == "":
            break
        exclude_weekdays.add(str(v))

    # --- 元データ読み込み ---
    ws_data = wb["時間帯別稼働推移元データ"]
    records = []
    for row in ws_data.iter_rows(min_row=2, max_row=ws_data.max_row, values_only=True):
        mgmt_no, op_date, weekday, room, start_time, end_time, category = row
        if room is not None:
            records.append({
                "date": str(op_date) if op_date else "",
                "weekday": str(weekday) if weekday else "",
                "room": str(room),
                "start": start_time,
                "end": end_time,
                "category": str(category) if category else "",
            })
    wb.close()

    records_filtered = [r for r in records if r["weekday"] not in exclude_weekdays]
    all_dates = sorted(set(r["date"] for r in records_filtered))
    num_days = len(all_dates)
    print(f"対象レコード数: {len(records_filtered)}, 対象日数: {num_days}")

    # 区分別件数
    cat_counts = {}
    for r in records_filtered:
        cat_counts[r["category"]] = cat_counts.get(r["category"], 0) + 1
    print(f"区分別件数: {cat_counts}")

    # --- 定時内スナップショット時刻: 9:00〜16:30 = 16区間 ---
    slots_9_1630 = []
    for h in range(9, 17):
        slots_9_1630.append(dt.time(h, 0))
        if h < 17:
            slots_9_1630.append(dt.time(h, 30))
    # 9:00, 9:30, 10:00, ..., 16:00, 16:30 = 16個
    slots_9_1630 = [s for s in slots_9_1630 if to_minutes(s) <= 16 * 60 + 30]
    print(f"定時内スナップショット: {len(slots_9_1630)}区間 ({slots_9_1630[0]}〜{slots_9_1630[-1]})")

    weight_total = sum(room_weight.values())

    # ========== 計算関数 ==========

    def calc_interval_overlap(data, slots, weights, weight_sum):
        """弊社方式: 区間重なり（-14分〜+15分）"""
        numerator = 0.0
        denominator = weight_sum * num_days * len(slots)
        for d in all_dates:
            day_recs = [r for r in data if r["date"] == d]
            for snap in slots:
                snap_min = to_minutes(snap)
                interval_a = snap_min - 14
                interval_b = snap_min + 15
                count = 0.0
                for r in day_recs:
                    room = r["room"]
                    if room not in weights:
                        continue
                    s = to_minutes(r["start"])
                    e = to_minutes(r["end"])
                    if interval_a <= e and s <= interval_b:
                        count += weights[room]
                numerator += count
        return (numerator / denominator * 100) if denominator > 0 else 0.0

    def calc_snapshot(data, slots, weights, weight_sum):
        """HOGY社方式: スナップショット（点判定: start ≤ snap < end）"""
        numerator = 0.0
        denominator = weight_sum * num_days * len(slots)
        for d in all_dates:
            day_recs = [r for r in data if r["date"] == d]
            for snap in slots:
                snap_min = to_minutes(snap)
                count = 0.0
                for r in day_recs:
                    room = r["room"]
                    if room not in weights:
                        continue
                    s = to_minutes(r["start"])
                    e = to_minutes(r["end"])
                    if s <= snap_min < e:
                        count += weights[room]
                numerator += count
        return (numerator / denominator * 100) if denominator > 0 else 0.0

    # ========== 各種データセット ==========
    all_surgery = [r for r in records_filtered if r["room"] in room_weight]
    scheduled_only = [r for r in all_surgery if r["category"] == "定時"]

    # ウェイト1.0固定
    room_weight_flat = {room: 1.0 for room in room_weight}
    weight_total_flat = sum(room_weight_flat.values())

    # 部屋除外版（ｱﾝｷﾞｵ除外 — 手術室として一般的でない）
    room_weight_no_angio = {k: v for k, v in room_weight.items() if k != "ｱﾝｷﾞｵ"}
    weight_no_angio = sum(room_weight_no_angio.values())
    scheduled_no_angio = [r for r in scheduled_only if r["room"] in room_weight_no_angio]

    # ========== 8試行実行 ==========
    print("\n" + "=" * 60)
    print("=== 稼働率試行計算結果 ===")
    print("=" * 60)

    results = {}

    # 試行1: 区間重なり + 全ウェイト + 全手術
    r1 = calc_interval_overlap(all_surgery, slots_9_1630, room_weight, weight_total)
    results["試行1"] = r1
    print(f"試行1: 区間重なり + 全ウェイト + 全手術        → {r1:.1f}%")

    # 試行2: スナップショット + 全ウェイト + 全手術
    r2 = calc_snapshot(all_surgery, slots_9_1630, room_weight, weight_total)
    results["試行2"] = r2
    print(f"試行2: スナップショット + 全ウェイト + 全手術    → {r2:.1f}%")

    # 試行3: スナップショット + 全ウェイト + 定時のみ
    r3 = calc_snapshot(scheduled_only, slots_9_1630, room_weight, weight_total)
    results["試行3"] = r3
    print(f"試行3: スナップショット + 全ウェイト + 定時のみ  → {r3:.1f}%")

    # 試行4: スナップショット + ウェイト1.0 + 定時のみ
    r4 = calc_snapshot(scheduled_only, slots_9_1630, room_weight_flat, weight_total_flat)
    results["試行4"] = r4
    print(f"試行4: スナップショット + ウェイト1.0 + 定時のみ → {r4:.1f}%")

    # 試行5: スナップショット + ｱﾝｷﾞｵ除外 + 定時のみ
    r5 = calc_snapshot(scheduled_no_angio, slots_9_1630, room_weight_no_angio, weight_no_angio)
    results["試行5"] = r5
    print(f"試行5: スナップショット + ｱﾝｷﾞｵ除外 + 定時のみ → {r5:.1f}%")

    # 試行6: 区間重なり + 全ウェイト + 定時のみ
    r6 = calc_interval_overlap(scheduled_only, slots_9_1630, room_weight, weight_total)
    results["試行6"] = r6
    print(f"試行6: 区間重なり + 全ウェイト + 定時のみ      → {r6:.1f}%")

    # 試行7: スナップショット + 全手術 + 16区間（9:00-16:30, 同じ）
    # ※ 試行2と同じスロットだが明示的に確認
    r7 = calc_snapshot(all_surgery, slots_9_1630, room_weight, weight_total)
    results["試行7"] = r7
    print(f"試行7: スナップショット + 全手術 + 16区間       → {r7:.1f}%")

    # 試行8: スナップショット + 全手術 + 分母=実部屋数×日数×時間帯数
    num_rooms_actual = len(room_weight)  # 実部屋数（ウェイトではなく部屋数）
    r8_num = 0.0
    r8_denom = num_rooms_actual * num_days * len(slots_9_1630)
    for d in all_dates:
        day_recs = [r for r in all_surgery if r["date"] == d]
        for snap in slots_9_1630:
            snap_min = to_minutes(snap)
            count = 0.0
            for r in day_recs:
                room = r["room"]
                if room not in room_weight:
                    continue
                s = to_minutes(r["start"])
                e = to_minutes(r["end"])
                if s <= snap_min < e:
                    count += room_weight[room]
            r8_num += count
    r8 = (r8_num / r8_denom * 100) if r8_denom > 0 else 0.0
    results["試行8"] = r8
    print(f"試行8: スナップショット + 分母=実部屋数        → {r8:.1f}%")

    # ========== 目標との比較 ==========
    target = 67.9
    print(f"\n目標: {target}%")

    closest = min(results.items(), key=lambda x: abs(x[1] - target))
    print(f"最も近い試行: {closest[0]}（{closest[1]:.1f}%、差 {abs(closest[1]-target):.1f}pt）")

    # ========== 詳細表 ==========
    print(f"\n{'='*80}")
    print(f"{'試行':8s} {'判定方式':16s} {'ウェイト':12s} {'対象手術':10s} {'部屋':10s} {'稼働率':8s} {'差':8s}")
    print(f"{'-'*80}")

    trial_info = [
        ("試行1", "区間重なり(-14/+15)", "定義シート準拠", "全手術", "全室", r1),
        ("試行2", "スナップショット", "定義シート準拠", "全手術", "全室", r2),
        ("試行3", "スナップショット", "定義シート準拠", "定時のみ", "全室", r3),
        ("試行4", "スナップショット", "全室1.0固定", "定時のみ", "全室", r4),
        ("試行5", "スナップショット", "定義シート準拠", "定時のみ", "ｱﾝｷﾞｵ除外", r5),
        ("試行6", "区間重なり(-14/+15)", "定義シート準拠", "定時のみ", "全室", r6),
        ("試行7", "スナップショット", "定義シート準拠", "全手術", "全室(16区間)", r7),
        ("試行8", "スナップショット", "定義シート準拠", "全手術", "分母=実部屋数", r8),
    ]

    for name, method, weight, surgery, room, rate in trial_info:
        diff = rate - target
        mark = " ★" if abs(diff) == abs(closest[1] - target) else ""
        print(f"{name:8s} {method:16s} {weight:12s} {surgery:10s} {room:12s} {rate:6.1f}% {diff:+6.1f}pt{mark}")

    # ========== 追加探索: 組み合わせで67.9%に近づく可能性 ==========
    print(f"\n{'='*80}")
    print("=== 追加探索: 条件組み合わせ ===")
    print(f"{'='*80}")

    combos = []

    # combo A: スナップショット + 定時のみ + ウェイト1.0 + ｱﾝｷﾞｵ除外
    room_flat_no_angio = {k: 1.0 for k in room_weight if k != "ｱﾝｷﾞｵ"}
    wt_flat_no_angio = sum(room_flat_no_angio.values())
    sched_no_angio_data = [r for r in scheduled_only if r["room"] in room_flat_no_angio]
    rA = calc_snapshot(sched_no_angio_data, slots_9_1630, room_flat_no_angio, wt_flat_no_angio)
    combos.append(("A: SS+定時+W1.0+ｱﾝｷﾞｵ除外", rA))

    # combo B: スナップショット + 全手術 + ウェイト1.0
    rB = calc_snapshot(all_surgery, slots_9_1630, room_weight_flat, weight_total_flat)
    combos.append(("B: SS+全手術+W1.0", rB))

    # combo C: 区間重なり + 定時のみ + ｱﾝｷﾞｵ除外
    rC = calc_interval_overlap(scheduled_no_angio, slots_9_1630, room_weight_no_angio, weight_no_angio)
    combos.append(("C: 区間+定時+ｱﾝｷﾞｵ除外", rC))

    # combo D: スナップショット + 全手術 + ｱﾝｷﾞｵ除外
    all_no_angio = [r for r in all_surgery if r["room"] in room_weight_no_angio]
    rD = calc_snapshot(all_no_angio, slots_9_1630, room_weight_no_angio, weight_no_angio)
    combos.append(("D: SS+全手術+ｱﾝｷﾞｵ除外", rD))

    # combo E: スナップショット + 定時のみ + 分母=実部屋数
    rE_num = 0.0
    rE_denom = len(room_weight) * num_days * len(slots_9_1630)
    for d in all_dates:
        day_recs = [r for r in scheduled_only if r["date"] == d]
        for snap in slots_9_1630:
            snap_min = to_minutes(snap)
            count = 0.0
            for r in day_recs:
                room = r["room"]
                if room not in room_weight:
                    continue
                s = to_minutes(r["start"])
                e = to_minutes(r["end"])
                if s <= snap_min < e:
                    count += room_weight[room]
            rE_num += count
    rE = (rE_num / rE_denom * 100) if rE_denom > 0 else 0.0
    combos.append(("E: SS+定時+分母=実部屋数", rE))

    # combo F: スナップショット + 定時のみ + 01A/01Bを1室として統合
    # 01A+01Bを「01」として扱い、どちらかが使用中なら1.0
    def calc_snapshot_merged_01(data, slots, weights_merged, weight_sum_merged):
        """01A/01Bを統合して1室扱いのスナップショット"""
        numerator = 0.0
        denominator = weight_sum_merged * num_days * len(slots)
        for d in all_dates:
            day_recs = [r for r in data if r["date"] == d]
            for snap in slots:
                snap_min = to_minutes(snap)
                count = 0.0
                room_01_active = False
                for r in day_recs:
                    room = r["room"]
                    if room not in room_weight:
                        continue
                    s = to_minutes(r["start"])
                    e = to_minutes(r["end"])
                    if s <= snap_min < e:
                        if room in ("01A", "01B"):
                            room_01_active = True
                        else:
                            count += weights_merged.get(room, 0)
                if room_01_active:
                    count += 1.0
                numerator += count
        return (numerator / denominator * 100) if denominator > 0 else 0.0

    weights_merged = {k: v for k, v in room_weight.items() if k not in ("01A", "01B")}
    weights_merged["01"] = 1.0
    wt_merged = sum(weights_merged.values())
    rF = calc_snapshot_merged_01(scheduled_only, slots_9_1630, weights_merged, wt_merged)
    combos.append(("F: SS+定時+01AB統合1室", rF))

    # combo G: スナップショット + 定時のみ + ウェイト1.0 + 01AB統合 + ｱﾝｷﾞｵ除外
    def calc_snapshot_merged_01_flat(data, slots, rooms_set, num_rooms):
        """01A/01B統合 + ウェイト1.0 + 特定部屋のみ"""
        numerator = 0.0
        denominator = num_rooms * num_days * len(slots)
        for d in all_dates:
            day_recs = [r for r in data if r["date"] == d]
            for snap in slots:
                snap_min = to_minutes(snap)
                count = 0.0
                room_01_active = False
                for r in day_recs:
                    room = r["room"]
                    if room not in room_weight:
                        continue
                    s = to_minutes(r["start"])
                    e = to_minutes(r["end"])
                    if s <= snap_min < e:
                        if room in ("01A", "01B"):
                            room_01_active = True
                        elif room in rooms_set:
                            count += 1.0
                if room_01_active and "01" in rooms_set:
                    count += 1.0
                numerator += count
        return (numerator / denominator * 100) if denominator > 0 else 0.0

    rooms_g = {"01", "02", "03", "05", "06", "07", "08", "09", "10"}  # 9室（ｱﾝｷﾞｵ除外、01AB統合）
    rG = calc_snapshot_merged_01_flat(scheduled_only, slots_9_1630, rooms_g, len(rooms_g))
    combos.append(("G: SS+定時+W1.0+01統合+ｱﾝｷﾞｵ除", rG))

    for name, rate in combos:
        diff = rate - target
        mark = " ★" if abs(diff) < 1.0 else ""
        print(f"  {name:40s} → {rate:6.1f}% (差 {diff:+5.1f}pt){mark}")

    # 全結果から最も近いものを選出
    all_results = list(results.items()) + [(f"追加{n}", r) for n, r in combos]
    best = min(all_results, key=lambda x: abs(x[1] - target))
    print(f"\n全試行中最も近い: {best[0]}（{best[1]:.1f}%、差 {abs(best[1]-target):.1f}pt）")

    # ========== 分析 ==========
    print(f"\n{'='*80}")
    print("=== 分析 ===")
    print(f"{'='*80}")
    print(f"弊社現行方式（試行1）: {r1:.1f}%")
    print(f"HOGY社目標値: {target}%")
    print(f"差: {r1 - target:.1f}pt")
    print()
    print("【判定方式の影響】")
    print(f"  区間重なり→スナップショットの変更効果: {r2 - r1:+.1f}pt (試行1→試行2)")
    print()
    print("【対象手術の影響】")
    print(f"  全手術→定時のみの変更効果(SS方式): {r3 - r2:+.1f}pt (試行2→試行3)")
    print(f"  全手術→定時のみの変更効果(区間方式): {r6 - r1:+.1f}pt (試行1→試行6)")
    print()
    print("【ウェイトの影響】")
    print(f"  定義準拠→1.0固定の変更効果: {r4 - r3:+.1f}pt (試行3→試行4)")
    print()
    print("【部屋除外の影響】")
    print(f"  全室→ｱﾝｷﾞｵ除外の変更効果: {r5 - r3:+.1f}pt (試行3→試行5)")
    print()
    print("【HOGY社方式の推定】")
    print(f"  最有力候補: {best[0]} = {best[1]:.1f}%")
    print(f"  HOGY社との差: {abs(best[1]-target):.1f}pt")
    if abs(best[1] - target) > 2.0:
        print("  ※ 2pt以上の差があるため、HOGY社は追加の条件差がある可能性:")
        print("    - 対象月・対象期間の違い")
        print("    - 手術室の定義差（特定部屋の除外）")
        print("    - 稼働率の分母・分子定義の違い")
        print("    - 端数処理・丸め方の違い")


if __name__ == "__main__":
    main()
