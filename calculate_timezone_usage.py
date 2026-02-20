"""
時間帯別稼働推移 計算スクリプト
================================
手術実施データから、30分おきのスナップショット時刻を中心とした
計測区間（-14分〜+15分）で何室稼働中かを計算し、計算結果シートに出力します。

グラフは元ファイルの「グラフ表示」シートに事前設定済みのため、
計算結果が書き込まれると自動でグラフに反映されます。

使い方:
    python calculate_timezone_usage.py

入力: 時間帯別稼働推移元データ.xlsx（同一フォルダに配置）
出力: 時間帯別稼働推移-結果.xlsx（同一フォルダに生成）
"""

import openpyxl
import datetime as dt
import os
import sys

# ========== 設定 ==========
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "時間帯別稼働推移元データ.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "時間帯別稼働推移-結果.xlsx")


def to_minutes(t):
    """時刻を分に変換（time, timedelta, str対応）"""
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
    print(f"入力ファイル読み込み: {INPUT_FILE}")
    wb = openpyxl.load_workbook(INPUT_FILE)

    # --- 定義シートから設定読み込み ---
    ws_def = wb["定義"]

    # 対象手術室とウェイト
    room_weight = {}
    for row in ws_def.iter_rows(min_row=2, max_row=20, min_col=1, max_col=2, values_only=True):
        if row[0] is not None and row[1] is not None:
            try:
                room_weight[str(row[0])] = float(row[1])
            except (ValueError, TypeError):
                break
    print(f"対象手術室: {room_weight}")

    # 除外曜日
    exclude_weekdays = set()
    r = 14
    while True:
        r += 1
        v = ws_def.cell(row=r, column=1).value
        if v is None or v == "":
            break
        exclude_weekdays.add(str(v))
    print(f"除外曜日: {exclude_weekdays}")

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

    print(f"総レコード数: {len(records)}")

    # 除外曜日フィルタリング
    records_filtered = [r for r in records if r["weekday"] not in exclude_weekdays]
    print(f"除外後レコード数: {len(records_filtered)}")

    # --- スナップショット時刻（8:00から30分おき、20:00まで = 25個）---
    snapshot_times = []
    for h in range(8, 20):
        snapshot_times.append(dt.time(h, 0))
        snapshot_times.append(dt.time(h, 30))
    snapshot_times.append(dt.time(20, 0))

    # --- 集計関数 ---
    def count_rooms_at_snapshots(data):
        """各スナップショット時刻の計測区間（-14分〜+15分）での稼働室数（ウェイト付き）を日平均で算出"""
        days = {}
        for r in data:
            d = r["date"]
            if d not in days:
                days[d] = []
            days[d].append(r)

        num_days = len(days)
        if num_days == 0:
            return [0.0] * len(snapshot_times)

        totals = [0.0] * len(snapshot_times)

        for day_str, day_records in days.items():
            for si, snap in enumerate(snapshot_times):
                snap_min = to_minutes(snap)
                interval_a = snap_min - 14  # 計測区間開始
                interval_b = snap_min + 15  # 計測区間終了
                count = 0.0
                for r in day_records:
                    room = r["room"]
                    if room not in room_weight:
                        continue
                    start_min = to_minutes(r["start"])
                    end_min = to_minutes(r["end"])
                    # 閉区間 [A,B] と [C,D] の重なり判定
                    if interval_a <= end_min and start_min <= interval_b:
                        count += room_weight[room]
                totals[si] += count

        averages = [round(t / num_days, 2) for t in totals]
        return averages

    # --- 全手術（定時・臨時・緊急）---
    all_surgery = [r for r in records_filtered if r["room"] in room_weight]
    print(f"\n全手術（対象室のみ）: {len(all_surgery)} 件")
    all_results = count_rooms_at_snapshots(all_surgery)

    # --- 予定手術のみ（定時のみ）---
    scheduled_only = [r for r in records_filtered if r["room"] in room_weight and r["category"] == "定時"]
    print(f"予定手術のみ: {len(scheduled_only)} 件")
    sched_results = count_rooms_at_snapshots(scheduled_only)

    # --- 計算結果シートに書き込み ---
    ws_result = wb["計算結果"]

    # 全体集計（Row2-3）
    for i, val in enumerate(all_results):
        ws_result.cell(row=2, column=2 + i, value=val)

    for i, val in enumerate(sched_results):
        ws_result.cell(row=3, column=2 + i, value=val)

    # --- 曜日別集計（Row5〜33）---
    weekday_rows = {
        "月曜日": {"all_row": 7,  "sched_row": 8},
        "火曜日": {"all_row": 12, "sched_row": 13},
        "水曜日": {"all_row": 17, "sched_row": 18},
        "木曜日": {"all_row": 22, "sched_row": 23},
        "金曜日": {"all_row": 27, "sched_row": 28},
        "土曜日": {"all_row": 32, "sched_row": 33},
    }

    print("\n--- 曜日別集計 ---")
    for weekday_name, rows in weekday_rows.items():
        weekday_records = [r for r in records if r["weekday"] == weekday_name and r["room"] in room_weight]
        weekday_scheduled = [r for r in weekday_records if r["category"] == "定時"]

        wd_all_results = count_rooms_at_snapshots(weekday_records)
        wd_sched_results = count_rooms_at_snapshots(weekday_scheduled)

        for i, val in enumerate(wd_all_results):
            ws_result.cell(row=rows["all_row"], column=2 + i, value=val)
        for i, val in enumerate(wd_sched_results):
            ws_result.cell(row=rows["sched_row"], column=2 + i, value=val)

        # 対象日数を算出
        wd_days = set(r["date"] for r in weekday_records)
        print(f"  {weekday_name}: 全手術={wd_all_results}, 予定={wd_sched_results}, 対象日数={len(wd_days)}")

    # --- 検証用シート（定時・臨時・緊急別の部屋数）---
    def count_rooms_by_day(data):
        """日別×スナップショット時刻の稼働室数（ウェイト付き）を返す: {date: [val, ...]}"""
        days = {}
        for r in data:
            d = r["date"]
            if d not in days:
                days[d] = []
            days[d].append(r)

        result = {}
        for day_str, day_records in sorted(days.items()):
            counts = []
            for si, snap in enumerate(snapshot_times):
                snap_min = to_minutes(snap)
                interval_a = snap_min - 14
                interval_b = snap_min + 15
                count = 0.0
                for r in day_records:
                    room = r["room"]
                    if room not in room_weight:
                        continue
                    start_min = to_minutes(r["start"])
                    end_min = to_minutes(r["end"])
                    if interval_a <= end_min and start_min <= interval_b:
                        count += room_weight[room]
                counts.append(count)
            result[day_str] = counts
        return result

    # 区分別データ
    sched_data = [r for r in records_filtered if r["room"] in room_weight and r["category"] == "定時"]
    urgent_data = [r for r in records_filtered if r["room"] in room_weight and r["category"] == "臨時"]
    emerg_data = [r for r in records_filtered if r["room"] in room_weight and r["category"] == "緊急"]

    sched_by_day = count_rooms_by_day(sched_data)
    urgent_by_day = count_rooms_by_day(urgent_data)
    emerg_by_day = count_rooms_by_day(emerg_data)
    all_by_day = count_rooms_by_day(all_surgery)

    # 全日付の和集合（ソート済み）
    all_dates = sorted(set(list(sched_by_day.keys()) + list(urgent_by_day.keys()) +
                           list(emerg_by_day.keys()) + list(all_by_day.keys())))

    # 既存シートがあれば削除して再作成
    verify_sheet_name = "検証_定時臨時緊急別"
    if verify_sheet_name in wb.sheetnames:
        del wb[verify_sheet_name]
    ws_verify = wb.create_sheet(verify_sheet_name)

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    header_font = Font(name="Meiryo UI", size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2980B9")
    label_font = Font(name="Meiryo UI", size=10, bold=True)
    data_font = Font(name="Meiryo UI", size=9)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    def write_section(ws, start_row, section_label, by_day_data, dates):
        """1セクションを書き込み、最終行+1を返す"""
        # セクションラベル
        cell = ws.cell(row=start_row, column=1, value=section_label)
        cell.font = label_font

        # ヘッダ行
        hr = start_row + 1
        cell = ws.cell(row=hr, column=1, value="時間帯")
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

        # B列: 全平日合計ヘッダ
        total_fill = PatternFill("solid", fgColor="1A5276")
        cell = ws.cell(row=hr, column=2, value="全平日合計")
        cell.font = header_font
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

        # C列以降: 日付ヘッダ
        for di, d in enumerate(dates):
            cell = ws.cell(row=hr, column=3 + di, value=d)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

        # データ行
        for si, snap in enumerate(snapshot_times):
            r = hr + 1 + si
            time_label = f"{snap.hour}:{snap.minute:02d}"
            cell = ws.cell(row=r, column=1, value=time_label)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

            # B列: 全平日合計（全日付の合算）
            day_total = 0.0
            for d in dates:
                vals = by_day_data.get(d, [0.0] * len(snapshot_times))
                day_total += vals[si] if si < len(vals) else 0.0
            cell = ws.cell(row=r, column=2, value=day_total)
            cell.font = Font(name="Meiryo UI", size=9, bold=True)
            cell.border = thin_border
            cell.number_format = "0.0"

            # C列以降: 日別値
            for di, d in enumerate(dates):
                vals = by_day_data.get(d, [0.0] * len(snapshot_times))
                val = vals[si] if si < len(vals) else 0.0
                cell = ws.cell(row=r, column=3 + di, value=val)
                cell.font = data_font
                cell.border = thin_border
                cell.number_format = "0.0"

        return hr + 1 + len(snapshot_times) + 1  # 次セクション開始行（1行空け）

    # 列幅設定
    ws_verify.column_dimensions["A"].width = 10
    ws_verify.column_dimensions["B"].width = 14
    for di in range(len(all_dates)):
        col_letter = openpyxl.utils.get_column_letter(3 + di)
        ws_verify.column_dimensions[col_letter].width = 12

    # 4セクション書き込み
    current_row = 1
    current_row = write_section(ws_verify, current_row, "【定時のみ】", sched_by_day, all_dates)
    current_row = write_section(ws_verify, current_row, "【臨時のみ】", urgent_by_day, all_dates)
    current_row = write_section(ws_verify, current_row, "【緊急のみ】", emerg_by_day, all_dates)
    current_row = write_section(ws_verify, current_row, "【合計（検証用）】", all_by_day, all_dates)

    print(f"\n検証シート '{verify_sheet_name}' を作成しました")

    # --- 検証：定時+臨時+緊急 = 全手術か ---
    mismatch_count = 0
    total_cells = 0
    for di, d in enumerate(all_dates):
        s_vals = sched_by_day.get(d, [0.0] * len(snapshot_times))
        u_vals = urgent_by_day.get(d, [0.0] * len(snapshot_times))
        e_vals = emerg_by_day.get(d, [0.0] * len(snapshot_times))
        a_vals = all_by_day.get(d, [0.0] * len(snapshot_times))
        for si in range(len(snapshot_times)):
            total_cells += 1
            s = s_vals[si] if si < len(s_vals) else 0.0
            u = u_vals[si] if si < len(u_vals) else 0.0
            e = e_vals[si] if si < len(e_vals) else 0.0
            a = a_vals[si] if si < len(a_vals) else 0.0
            if abs((s + u + e) - a) > 0.01:
                mismatch_count += 1

    print(f"\n=== 検証結果 ===")
    print(f"1. 定時+臨時+緊急 = 全手術 の一致チェック: "
          f"{'全セル一致 OK' if mismatch_count == 0 else f'{mismatch_count}/{total_cells} セル不一致'}")

    # 臨時+緊急の割合
    print(f"\n2. 件数内訳:")
    print(f"   定時: {len(sched_data)} 件")
    print(f"   臨時: {len(urgent_data)} 件")
    print(f"   緊急: {len(emerg_data)} 件")
    print(f"   合計: {len(sched_data) + len(urgent_data) + len(emerg_data)} 件 "
          f"(全手術={len(all_surgery)} 件)")

    non_scheduled = len(urgent_data) + len(emerg_data)
    total = len(all_surgery)
    ratio = non_scheduled / total * 100 if total > 0 else 0
    print(f"\n3. 臨時+緊急が全体に占める割合: {non_scheduled}/{total} = {ratio:.1f}%")

    # --- 別名保存 ---
    wb.save(OUTPUT_FILE)
    print(f"\n計算完了: {OUTPUT_FILE}")
    print(f"  全手術:   {all_results}")
    print(f"  予定のみ: {sched_results}")


if __name__ == "__main__":
    main()
