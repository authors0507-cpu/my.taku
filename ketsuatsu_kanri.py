import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import os
import json
import random

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

DATA_FILE = "血圧記録.json"
EXCEL_FILE = "血圧記録.xlsx"

TIME_SLOTS = ["朝", "日中", "夜"]
TIME_SLOT_ORDER = {"朝": 0, "日中": 1, "夜": 2}

# JSH2019家庭血圧の目安を簡略化した6段階（数値は参考情報であり診断ではありません）
LEVELS = [
    ("至適血圧", "#27ae60"),
    ("正常血圧", "#2ecc71"),
    ("正常高値血圧", "#f4d03f"),
    ("I度高血圧", "#f39c12"),
    ("II度高血圧", "#e67e22"),
    ("III度高血圧", "#e74c3c"),
]

LEVEL_MESSAGES = {
    0: "とても良い数値です。この調子で今の生活習慣を続けましょう。",
    1: "正常範囲です。今の生活習慣を維持しましょう。",
    2: "やや高めです。減塩や運動を意識すると安心です。",
    3: "高血圧の可能性があります。生活習慣の見直しをおすすめします。",
    4: "高血圧の可能性が高い数値です。継続するようであれば医療機関への相談をご検討ください。",
    5: "非常に高い数値です。早めに医療機関の受診をおすすめします。",
}

GENERAL_TIPS = [
    "麺類の汁は残す（塩分の多くはスープに含まれています）",
    "醤油やソースは「かける」より「つける」で使用量を減らす",
    "ハム・練り物・インスタント食品などの加工食品は控えめに",
    "だしを効かせると、塩を減らしても美味しく食べられます",
    "生姜・ねぎ・大葉・柑橘類など香味野菜や酸味で味にアクセントを",
    "バナナ・ほうれん草・アボカドなどカリウムを含む野菜・果物を意識的に（腎機能に不安がある方は主治医に相談）",
    "漬物・佃煮・梅干しは量を控えめに",
    "味噌汁は具だくさんにして汁の量を減らす",
    "外食・コンビニ食は栄養成分表示の食塩相当量を確認する習慣を",
    "お酒はほどほどに（目安は1日あたり日本酒1合・ビール中瓶1本程度まで）",
    "減塩醤油・減塩味噌など減塩調味料を活用する",
    "スナック菓子や加工肉を減らし、素材そのものの味を楽しむ",
    "毎日同じ時間・同じ条件で血圧を測る習慣をつける（起床後1時間以内・朝食前や排尿後がおすすめ）",
    "ウォーキングなど軽い有酸素運動を1日30分程度心がける",
    "適正体重を維持する（肥満は血圧を上げやすくなります）",
    "禁煙・受動喫煙を避ける",
    "十分な睡眠とストレスケアを心がける",
    "湯船にゆっくり浸かるなど、リラックスできる時間を作る",
]


def classify_bp(sbp, dbp):
    if sbp < 115:
        sl = 0
    elif sbp < 125:
        sl = 1
    elif sbp < 135:
        sl = 2
    elif sbp < 145:
        sl = 3
    elif sbp < 160:
        sl = 4
    else:
        sl = 5

    if dbp < 75:
        dl = 0
    elif dbp < 85:
        dl = 2
    elif dbp < 90:
        dl = 3
    elif dbp < 100:
        dl = 4
    else:
        dl = 5

    level = max(sl, dl)
    label, color = LEVELS[level]
    return level, label, color


def build_advice(level):
    lines = [LEVEL_MESSAGES[level], "", "【今日の食事・生活アドバイス】"]
    for tip in random.sample(GENERAL_TIPS, k=3):
        lines.append(f"・{tip}")
    return "\n".join(lines)


def load_records():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, encoding="utf-8") as f:
            return json.load(f)
    return []


def save_records(records):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def sort_key(r):
    return (r["date"], TIME_SLOT_ORDER.get(r["time_slot"], 1))


def compute_weekly_stats(records):
    today = datetime.date.today()
    recent = [
        r for r in records
        if (today - datetime.date.fromisoformat(r["date"])).days < 7
    ]
    if not recent:
        return None
    avg_sbp = sum(r["sbp"] for r in recent) / len(recent)
    avg_dbp = sum(r["dbp"] for r in recent) / len(recent)
    return avg_sbp, avg_dbp, len(recent)


def export_to_excel(records):
    if not EXCEL_AVAILABLE:
        raise RuntimeError("openpyxl がインストールされていません（pip install openpyxl）")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "血圧記録"
    headers = ["日付", "時間帯", "収縮期(mmHg)", "拡張期(mmHg)", "脈拍(回/分)", "判定", "メモ"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for r in sorted(records, key=sort_key):
        _, label, color = classify_bp(r["sbp"], r["dbp"])
        row = [
            r["date"], r["time_slot"], r["sbp"], r["dbp"],
            r.get("pulse") or "", label, r.get("memo", ""),
        ]
        ws.append(row)
        fill = PatternFill(start_color=color.replace("#", ""), end_color=color.replace("#", ""), fill_type="solid")
        for cell in ws[ws.max_row]:
            cell.fill = fill

    for col, width in zip("ABCDEFG", (12, 8, 14, 14, 12, 16, 30)):
        ws.column_dimensions[col].width = width

    wb.save(EXCEL_FILE)


class BPApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("血圧健康管理アプリ")
        self.geometry("640x760")
        self.records = load_records()

        self._build_input_frame()
        self._build_history_frame()
        self._build_advice_frame()
        self._build_chart_frame()

        self.refresh_all()

    def _build_input_frame(self):
        frame = ttk.LabelFrame(self, text="今日の記録")
        frame.pack(fill="x", padx=10, pady=8)

        ttk.Label(frame, text="日付").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(frame, textvariable=self.date_var, width=12).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="時間帯").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.slot_var = tk.StringVar(value="朝")
        ttk.Combobox(frame, textvariable=self.slot_var, values=TIME_SLOTS, width=6, state="readonly").grid(
            row=0, column=3, padx=5, pady=5, sticky="w"
        )

        ttk.Label(frame, text="収縮期(上)").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.sbp_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.sbp_var, width=8).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="拡張期(下)").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.dbp_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.dbp_var, width=8).grid(row=1, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="脈拍").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.pulse_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.pulse_var, width=8).grid(row=1, column=5, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="メモ").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.memo_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.memo_var, width=40).grid(
            row=2, column=1, columnspan=5, padx=5, pady=5, sticky="w"
        )

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=6, pady=8)
        ttk.Button(btn_frame, text="記録する", command=self.on_save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="選択した記録を削除", command=self.on_delete).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Excelに出力", command=self.on_export).pack(side="left", padx=5)

    def _build_history_frame(self):
        frame = ttk.LabelFrame(self, text="記録履歴")
        frame.pack(fill="both", expand=False, padx=10, pady=8)

        columns = ("date", "slot", "sbp", "dbp", "pulse", "label", "memo")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        headings = {
            "date": "日付", "slot": "時間帯", "sbp": "収縮期", "dbp": "拡張期",
            "pulse": "脈拍", "label": "判定", "memo": "メモ",
        }
        widths = {"date": 90, "slot": 50, "sbp": 55, "dbp": 55, "pulse": 50, "label": 100, "memo": 140}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="center")
        self.tree.column("memo", anchor="w")
        self.tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.stats_label = ttk.Label(self, text="")
        self.stats_label.pack(fill="x", padx=15)

    def _build_advice_frame(self):
        frame = ttk.LabelFrame(self, text="判定とアドバイス")
        frame.pack(fill="both", expand=False, padx=10, pady=8)

        self.judgement_label = ttk.Label(frame, text="", font=("", 14, "bold"))
        self.judgement_label.pack(anchor="w", padx=10, pady=(5, 0))

        self.advice_text = tk.Text(frame, height=8, wrap="word", state="disabled")
        self.advice_text.pack(fill="both", expand=True, padx=10, pady=5)

        ttk.Button(frame, text="アドバイスを更新", command=self.refresh_advice).pack(pady=(0, 8))

    def _build_chart_frame(self):
        frame = ttk.LabelFrame(self, text="血圧の推移（直近14件）")
        frame.pack(fill="both", expand=True, padx=10, pady=8)
        self.canvas = tk.Canvas(frame, width=580, height=220, bg="white", highlightthickness=0)
        self.canvas.pack(padx=5, pady=5)

    def on_save(self):
        try:
            date_str = self.date_var.get().strip()
            datetime.date.fromisoformat(date_str)
        except ValueError:
            messagebox.showerror("入力エラー", "日付は YYYY-MM-DD の形式で入力してください")
            return

        try:
            sbp = int(self.sbp_var.get())
            dbp = int(self.dbp_var.get())
        except ValueError:
            messagebox.showerror("入力エラー", "収縮期・拡張期血圧は数値で入力してください")
            return

        if not (50 <= sbp <= 300 and 30 <= dbp <= 200):
            messagebox.showerror("入力エラー", "血圧の値が現実的な範囲を超えています")
            return

        pulse_str = self.pulse_var.get().strip()
        pulse = None
        if pulse_str:
            try:
                pulse = int(pulse_str)
            except ValueError:
                messagebox.showerror("入力エラー", "脈拍は数値で入力してください")
                return

        record = {
            "date": date_str,
            "time_slot": self.slot_var.get(),
            "sbp": sbp,
            "dbp": dbp,
            "pulse": pulse,
            "memo": self.memo_var.get().strip(),
            "created_at": datetime.datetime.now().isoformat(),
        }
        self.records.append(record)
        save_records(self.records)

        self.sbp_var.set("")
        self.dbp_var.set("")
        self.pulse_var.set("")
        self.memo_var.set("")

        self.refresh_all()

    def on_delete(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("削除", "削除する記録を履歴から選択してください")
            return
        index = int(selected[0])
        sorted_records = sorted(self.records, key=sort_key)
        target = sorted_records[index]
        self.records.remove(target)
        save_records(self.records)
        self.refresh_all()

    def on_export(self):
        if not self.records:
            messagebox.showinfo("Excelに出力", "記録がありません")
            return
        try:
            export_to_excel(self.records)
        except RuntimeError as e:
            messagebox.showerror("Excelに出力", str(e))
            return
        messagebox.showinfo("Excelに出力", f"{EXCEL_FILE} に出力しました")

    def refresh_all(self):
        self.refresh_history()
        self.refresh_stats()
        self.refresh_advice()
        self.refresh_chart()

    def refresh_history(self):
        self.tree.delete(*self.tree.get_children())
        sorted_records = sorted(self.records, key=sort_key)
        for i, r in enumerate(sorted_records):
            _, label, _ = classify_bp(r["sbp"], r["dbp"])
            self.tree.insert(
                "", "end", iid=str(i),
                values=(r["date"], r["time_slot"], r["sbp"], r["dbp"], r.get("pulse") or "", label, r.get("memo", "")),
            )
        if sorted_records:
            self.tree.selection_set(str(len(sorted_records) - 1))
            self.tree.see(str(len(sorted_records) - 1))

    def refresh_stats(self):
        stats = compute_weekly_stats(self.records)
        if stats is None:
            self.stats_label.config(text="直近7日間の記録はまだありません")
        else:
            avg_sbp, avg_dbp, n = stats
            self.stats_label.config(
                text=f"直近7日間の平均: {avg_sbp:.0f} / {avg_dbp:.0f} mmHg（{n}件）"
            )

    def refresh_advice(self):
        if not self.records:
            self.judgement_label.config(text="記録がありません")
            level = None
        else:
            latest = sorted(self.records, key=sort_key)[-1]
            level, label, color = classify_bp(latest["sbp"], latest["dbp"])
            self.judgement_label.config(
                text=f"最新の記録: {latest['sbp']}/{latest['dbp']} mmHg → {label}", foreground=color
            )

        self.advice_text.config(state="normal")
        self.advice_text.delete("1.0", "end")
        if level is None:
            self.advice_text.insert(
                "1.0",
                "まずは今日の血圧を記録してみましょう。\n\n【今日の食事・生活アドバイス】\n・"
                + "\n・".join(random.sample(GENERAL_TIPS, k=3)),
            )
        else:
            self.advice_text.insert("1.0", build_advice(level))
        self.advice_text.config(state="disabled")

    def refresh_chart(self):
        canvas = self.canvas
        canvas.delete("all")
        recent = sorted(self.records, key=sort_key)[-14:]
        if len(recent) < 2:
            canvas.create_text(290, 110, text="記録が2件以上になるとグラフを表示します", fill="#888888")
            return

        w, h = 580, 220
        pad_l, pad_r, pad_t, pad_b = 45, 15, 20, 30
        plot_w = w - pad_l - pad_r
        plot_h = h - pad_t - pad_b
        vmin, vmax = 50, 200

        def y_of(v):
            v = max(vmin, min(vmax, v))
            return pad_t + plot_h * (1 - (v - vmin) / (vmax - vmin))

        def x_of(i):
            if len(recent) == 1:
                return pad_l
            return pad_l + plot_w * i / (len(recent) - 1)

        canvas.create_line(pad_l, pad_t, pad_l, h - pad_b, fill="#999999")
        canvas.create_line(pad_l, h - pad_b, w - pad_r, h - pad_b, fill="#999999")
        for v in (100, 135, 140, 160):
            y = y_of(v)
            canvas.create_line(pad_l, y, w - pad_r, y, fill="#dddddd", dash=(2, 2))
            canvas.create_text(pad_l - 8, y, text=str(v), anchor="e", font=("", 8))

        sbp_pts = [(x_of(i), y_of(r["sbp"])) for i, r in enumerate(recent)]
        dbp_pts = [(x_of(i), y_of(r["dbp"])) for i, r in enumerate(recent)]
        for pts, color in ((sbp_pts, "#e74c3c"), (dbp_pts, "#3498db")):
            for i in range(len(pts) - 1):
                canvas.create_line(*pts[i], *pts[i + 1], fill=color, width=2)
            for x, y in pts:
                canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=color, outline="")

        step = max(1, len(recent) // 7)
        for i, r in enumerate(recent):
            if i % step == 0:
                canvas.create_text(x_of(i), h - pad_b + 12, text=r["date"][5:], font=("", 7))

        canvas.create_text(pad_l + 15, 8, text="● 収縮期", fill="#e74c3c", anchor="w", font=("", 8))
        canvas.create_text(pad_l + 75, 8, text="● 拡張期", fill="#3498db", anchor="w", font=("", 8))


if __name__ == "__main__":
    app = BPApp()
    app.mainloop()
