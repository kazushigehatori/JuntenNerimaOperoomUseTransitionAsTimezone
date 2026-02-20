"""
HOGY社方式との差異調査 — 全14試行の稼働率比較
==============================================
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

    cat_counts = {}
    for r in records_filtered:
        cat_counts[r["category"]] = cat_counts.get(r["category"], 0) + 1
    print(f"区分別件数: {cat_counts}")

    weight_total = sum(room_weight.values())

    # --- スロット定義 ---
    # 9:00〜16:30 = 16区間
    slots_16 = []
    for h in range(9, 17):
        slots_16.append(dt.time(h, 0))
        if h < 17:
            slots_16.append(dt.time(h, 30))
    slots_16 = [s for s in slots_16 if to_minutes(s) <= 16 * 60 + 30]
    print(f"スロット(16区間): {len(slots_16)}個 (9:00〜16:30)")

    # --- データセット ---
    all_surgery = [r for r in records_filtered if r["room"] in room_weight]
    scheduled_only = [r for r in all_surgery if r["category"] == "定時"]
    room_weight_flat = {room: 1.0 for room in room_weight}
    weight_total_flat = sum(room_weight_flat.values())
    room_weight_no_angio = {k: v for k, v in room_weight.items() if k != "ｱﾝｷﾞｵ"}
    weight_no_angio = sum(room_weight_no_angio.values())
    scheduled_no_angio = [r for r in scheduled_only if r["room"] in room_weight_no_angio]

    # ========== 汎用計算関数 ==========
    def calc_overlap(data, slots, weights, weight_sum, offset_a, offset_b):
        """区間重なり方式: [snap+offset_a, snap+offset_b] と手術時間の重なり判定"""
        numerator = 0.0
        denominator = weight_sum * num_days * len(slots)
        for d in all_dates:
            day_recs = [r for r in data if r["date"] == d]
            for snap in slots:
                snap_min = to_minutes(snap)
                ia = snap_min + offset_a
                ib = snap_min + offset_b
                count = 0.0
                for r in day_recs:
                    room = r["room"]
                    if room not in weights:
                        continue
                    s = to_minutes(r["start"])
                    e = to_minutes(r["end"])
                    if ia <= e and s <= ib:
                        count += weights[room]
                numerator += count
        return (numerator / denominator * 100) if denominator > 0 else 0.0

    def calc_snapshot(data, slots, weights, weight_sum):
        """スナップショット方式: start ≤ snap < end"""
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

    # ========== 全14試行 ==========
    target = 67.9

    trials = []

    # 試行1: 弊社区間(-14/+15) + 全ウェイト + 全手術 + 9:00-16:30
    r1 = calc_overlap(all_surgery, slots_16, room_weight, weight_total, -14, +15)
    trials.append(("試行1", "弊社区間(-14/+15)", "定義準拠", "全手術", "9:00-16:30", r1))

    # 試行2: スナップショット + 全ウェイト + 全手術 + 9:00-16:30
    r2 = calc_snapshot(all_surgery, slots_16, room_weight, weight_total)
    trials.append(("試行2", "スナップショット", "定義準拠", "全手術", "9:00-16:30", r2))

    # 試行3: スナップショット + 全ウェイト + 定時のみ + 9:00-16:30
    r3 = calc_snapshot(scheduled_only, slots_16, room_weight, weight_total)
    trials.append(("試行3", "スナップショット", "定義準拠", "定時のみ", "9:00-16:30", r3))

    # 試行4: スナップショット + ウェイト1.0 + 定時のみ + 9:00-16:30
    r4 = calc_snapshot(scheduled_only, slots_16, room_weight_flat, weight_total_flat)
    trials.append(("試行4", "スナップショット", "W1.0固定", "定時のみ", "9:00-16:30", r4))

    # 試行5: スナップショット + ｱﾝｷﾞｵ除外 + 定時のみ + 9:00-16:30
    r5 = calc_snapshot(scheduled_no_angio, slots_16, room_weight_no_angio, weight_no_angio)
    trials.append(("試行5", "スナップショット", "定義準拠", "定時のみ", "9:00-16:30(ｱﾝｷﾞｵ除)", r5))

    # 試行6: 弊社区間(-14/+15) + 全ウェイト + 定時のみ + 9:00-16:30
    r6 = calc_overlap(scheduled_only, slots_16, room_weight, weight_total, -14, +15)
    trials.append(("試行6", "弊社区間(-14/+15)", "定義準拠", "定時のみ", "9:00-16:30", r6))

    # 試行7: スナップショット + 全手術 + 16区間 (=試行2と同じ、明示確認)
    r7 = r2
    trials.append(("試行7", "スナップショット", "定義準拠", "全手術", "9:00-16:30(16区間)", r7))

    # 試行8: スナップショット + 全手術 + 分母=実部屋数
    num_rooms_actual = len(room_weight)
    r8_denom = num_rooms_actual * num_days * len(slots_16)
    r8_num = 0.0
    for d in all_dates:
        day_recs = [r for r in all_surgery if r["date"] == d]
        for snap in slots_16:
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
    trials.append(("試行8", "スナップショット", "定義準拠", "全手術", "分母=実部屋数", r8))

    # 試行9: HOGY区間(0/+29) + 定時のみ + 9:00-16:30
    r9 = calc_overlap(scheduled_only, slots_16, room_weight, weight_total, 0, +29)
    trials.append(("試行9", "HOGY区間(0/+29)", "定義準拠", "定時のみ", "9:00-16:30", r9))

    # 試行10: HOGY区間(0/+29) + 全手術 + 9:00-16:30
    r10 = calc_overlap(all_surgery, slots_16, room_weight, weight_total, 0, +29)
    trials.append(("試行10", "HOGY区間(0/+29)", "定義準拠", "全手術", "9:00-16:30", r10))

    # 試行11: HOGY区間(0/+29) + 定時のみ + 9:00-16:30 (=試行9と同じだが明示)
    r11 = r9
    trials.append(("試行11", "HOGY区間(0/+29)", "定義準拠", "定時のみ", "9:00-16:30(16区間)", r11))

    # 試行12: 弊社区間(-14/+15) + 定時のみ + 9:00-16:30 (=試行6と同じだが明示)
    r12 = r6
    trials.append(("試行12", "弊社区間(-14/+15)", "定義準拠", "定時のみ", "9:00-16:30(16区間)", r12))

    # 試行13: HOGY区間(0/+29) + 全手術 + 9:00-16:30 (=試行10と同じだが明示)
    r13 = r10
    trials.append(("試行13", "HOGY区間(0/+29)", "定義準拠", "全手術", "9:00-16:30(16区間)", r13))

    # 試行14: 弊社区間(-14/+15) + 全手術 + 9:00-16:30 (=試行1と同じだが明示)
    r14 = r1
    trials.append(("試行14", "弊社区間(-14/+15)", "定義準拠", "全手術", "9:00-16:30(16区間)", r14))

    # ========== 結果表示 ==========
    print("\n" + "=" * 90)
    print("=== 全試行結果一覧 ===")
    print("=" * 90)

    # 重複排除した実質的な試行のみ表示
    unique_trials = [
        ("試行1", "弊社区間(-14/+15) + 全ウェイト + 全手術 + 9:00-16:30", r1),
        ("試行2", "スナップショット + 全ウェイト + 全手術 + 9:00-16:30", r2),
        ("試行3", "スナップショット + 全ウェイト + 定時のみ + 9:00-16:30", r3),
        ("試行4", "スナップショット + ウェイト1.0 + 定時のみ + 9:00-16:30", r4),
        ("試行5", "スナップショット + ｱﾝｷﾞｵ除外 + 定時のみ + 9:00-16:30", r5),
        ("試行6", "弊社区間(-14/+15) + 全ウェイト + 定時のみ + 9:00-16:30", r6),
        ("試行7", "スナップショット + 全手術 + 9:00-16:30（= 試行2）", r7),
        ("試行8", "スナップショット + 分母=実部屋数 + 全手術 + 9:00-16:30", r8),
        ("試行9", "HOGY区間(0/+29) + 全ウェイト + 定時のみ + 9:00-16:30", r9),
        ("試行10", "HOGY区間(0/+29) + 全ウェイト + 全手術 + 9:00-16:30", r10),
        ("試行11", "HOGY区間(0/+29) + 定時のみ + 9:00-16:30（= 試行9）", r11),
        ("試行12", "弊社区間(-14/+15) + 定時のみ + 9:00-16:30（= 試行6）", r12),
        ("試行13", "HOGY区間(0/+29) + 全手術 + 9:00-16:30（= 試行10）", r13),
        ("試行14", "弊社区間(-14/+15) + 全手術 + 9:00-16:30（= 試行1）", r14),
    ]

    for name, desc, rate in unique_trials:
        diff = rate - target
        mark = " ★" if abs(diff) <= 2.0 else ""
        print(f"  {name:6s}: {desc:58s} → {rate:5.1f}% (差{diff:+5.1f}pt){mark}")

    # ========== 詳細比較表 ==========
    print(f"\n{'='*100}")
    print("=== 詳細比較表（重複除外） ===")
    print(f"{'='*100}")
    print(f"{'試行':8s} {'区間方式':20s} {'ウェイト':12s} {'対象手術':10s} {'時間帯':16s} {'稼働率':8s} {'差':8s}")
    print(f"{'-'*100}")

    dedup_trials = [
        ("試行1", "弊社(-14/+15)", "定義準拠", "全手術", "9:00-16:30", r1),
        ("試行2", "SS(点判定)", "定義準拠", "全手術", "9:00-16:30", r2),
        ("試行3", "SS(点判定)", "定義準拠", "定時のみ", "9:00-16:30", r3),
        ("試行4", "SS(点判定)", "全室1.0", "定時のみ", "9:00-16:30", r4),
        ("試行5", "SS(点判定)", "定義準拠", "定時ｱﾝｷﾞｵ除", "9:00-16:30", r5),
        ("試行6", "弊社(-14/+15)", "定義準拠", "定時のみ", "9:00-16:30", r6),
        ("試行8", "SS(点判定)", "定義準拠", "全手術", "分母=実部屋数", r8),
        ("試行9", "HOGY(0/+29)", "定義準拠", "定時のみ", "9:00-16:30", r9),
        ("試行10", "HOGY(0/+29)", "定義準拠", "全手術", "9:00-16:30", r10),
    ]

    closest_rate = min(dedup_trials, key=lambda x: abs(x[5] - target))
    for name, method, weight, surgery, slots_desc, rate in dedup_trials:
        diff = rate - target
        mark = " ★" if name == closest_rate[0] else ""
        print(f"{name:8s} {method:20s} {weight:12s} {surgery:12s} {slots_desc:16s} {rate:6.1f}% {diff:+6.1f}pt{mark}")

    print(f"\n目標: {target}%")
    print(f"最も近い試行: {closest_rate[0]}（{closest_rate[5]:.1f}%、差 {abs(closest_rate[5]-target):.1f}pt）")

    # ========== 影響分析 ==========
    print(f"\n{'='*80}")
    print("=== 条件別の影響分析 ===")
    print(f"{'='*80}")

    print("\n【区間方式の影響】")
    print(f"  弊社(-14/+15)→SS(点判定): {r2-r1:+.1f}pt (全手術: 試行1→2)")
    print(f"  弊社(-14/+15)→SS(点判定): {r3-r6:+.1f}pt (定時: 試行6→3)")
    print(f"  弊社(-14/+15)→HOGY(0/+29): {r10-r1:+.1f}pt (全手術: 試行1→10)")
    print(f"  弊社(-14/+15)→HOGY(0/+29): {r9-r6:+.1f}pt (定時: 試行6→9)")

    print("\n【対象手術の影響】")
    print(f"  全手術→定時のみ（弊社区間）: {r6-r1:+.1f}pt (試行1→6)")
    print(f"  全手術→定時のみ（SS方式）:   {r3-r2:+.1f}pt (試行2→3)")
    print(f"  全手術→定時のみ（HOGY区間）: {r9-r10:+.1f}pt (試行10→9)")

    print("\n【ウェイトの影響】")
    print(f"  定義準拠→1.0固定: {r4-r3:+.1f}pt (試行3→4)")

    print("\n【区間定義の比較（9:00枠の具体例）】")
    print(f"  弊社方式: 8:46 〜 9:15 (-14分〜+15分)")
    print(f"  HOGY方式: 9:00 〜 9:29 ( 0分〜+29分)")
    print(f"  SS方式 :  9:00 の瞬間（点判定）")

    print(f"\n{'='*80}")
    print("=== 結論 ===")
    print(f"{'='*80}")
    print(f"最も67.9%に近い: {closest_rate[0]} = {closest_rate[5]:.1f}%（差{abs(closest_rate[5]-target):.1f}pt）")
    print(f"条件: {closest_rate[1]} + {closest_rate[2]} + {closest_rate[3]} + {closest_rate[4]}")
    if abs(closest_rate[5] - target) <= 2.0:
        print("→ 差2pt以内。HOGY社はこの条件に近い計算方式と推定される。")
    else:
        print("→ 差2pt超。HOGY社は追加の条件差（対象期間・部屋定義等）がある可能性。")


if __name__ == "__main__":
    main()
