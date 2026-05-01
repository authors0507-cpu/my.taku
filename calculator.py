import tkinter as tk
from tkinter import ttk, messagebox
import datetime
import os
import re
import json
import urllib.request
import urllib.parse
import threading

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

EXCEL_FILE = "せどり_利益計算.xlsx"
CONFIG_FILE = "keepa_config.json"

PLATFORM_DEFAULTS = {
    "Amazon FBA": {
        "referral_rate": 15.0,
        "fba_fee": 400,
        "shipping": 0,
    },
    "Amazon 自己発送": {
        "referral_rate": 15.0,
        "fba_fee": 0,
        "shipping": 600,
    },
    "メルカリ": {
        "referral_rate": 10.0,
        "fba_fee": 0,
        "shipping": 500,
    },
}


def load_api_key():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f).get("api_key", "")
    return ""


def calc(purchase, selling, referral_rate, fba_fee, shipping):
    referral_fee = selling * (referral_rate / 100)
    profit = selling - purchase - referral_fee - fba_fee - shipping
    profit_rate = (profit / selling * 100) if selling > 0 else 0
    return referral_fee, profit, profit_rate


def extract_asin(text):
    """URLまたは直接入力からASINを抽出する"""
    text = text.strip()
    # URL から ASIN 抽出
    m = re.search(r"/(?:dp|gp/product)/([A-Z0-9]{10})", text)
    if m:
        return m.group(1)
    # 直接 ASIN (10文字英数字)
    if re.fullmatch(r"[A-Z0-9]{10}", text):
        return text
    return None


def fetch_keepa_price(api_key, asin):
    """Keepa APIで現在のAmazon Japan価格を取得する。価格(円)またはNoneを返す"""
    url = (
        "https://api.keepa.com/product"
        f"?key={urllib.parse.quote(api_key)}"
        f"&domain=5&asin={asin}&history=1&stats=1"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "SedoriTool/1.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read().decode())

    products = data.get("products")
    if not products:
        return None, None

    p = products[0]
    title = p.get("title", "")

    # stats.current[0] = Amazon現在価格, [1] = マーケットプレイス新品最安値
    stats = p.get("stats", {})
    current = stats.get("current", [])

    price = None
    for idx in (0, 1):
        if idx < len(current) and current[idx] and current[idx] > 0:
            price = current[idx]
            break

    # Keepa の価格は JPY の場合も /100 が必要
    if price is not None:
        price = price / 100

    return price, title


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("せどり 利益計算ツール（Keepa連携）")
        self.root.geometry("520x780")
        self.root.resizable(False, False)
        self.root.configure(bg="#F4F6F8")

        self.api_key = load_api_key()

        self.platform = tk.StringVar(value="Amazon FBA")
        self.asin_input = tk.StringVar()
        self.purchase = tk.StringVar()
        self.selling = tk.StringVar()
        self.referral_rate = tk.StringVar(value="15.0")
        self.fba_fee = tk.StringVar(value="400")
        self.shipping = tk.StringVar(value="0")
        self.product_title = tk.StringVar(value="")

        self._build()

        for v in (self.purchase, self.selling, self.referral_rate,
                  self.fba_fee, self.shipping):
            v.trace_add("write", self._refresh)

    # ── UI 構築 ──────────────────────────────────────────────────────────────

    def _build(self):
        tk.Label(self.root, text="せどり 利益計算ツール",
                 font=("Helvetica", 17, "bold"),
                 bg="#F4F6F8", fg="#2C3E50").pack(pady=(14, 2))

        self._section_keepa()
        self._section_platform()
        self._section_input()
        self._section_result()
        self._section_buttons()

        if not EXCEL_AVAILABLE:
            tk.Label(self.root,
                     text="※ openpyxl 未インストール — pip install openpyxl で Excel保存が使えます",
                     font=("Helvetica", 8), bg="#F4F6F8", fg="#E74C3C").pack(pady=2)

    def _section_keepa(self):
        frm = tk.LabelFrame(self.root, text="Keepa 価格自動取得",
                            font=("Helvetica", 10), bg="#F4F6F8", padx=12, pady=8)
        frm.pack(fill="x", padx=20, pady=(6, 4))
        frm.columnconfigure(1, weight=1)

        tk.Label(frm, text="ASIN / URL", font=("Helvetica", 11),
                 bg="#F4F6F8", anchor="w").grid(row=0, column=0, sticky="w", pady=4)

        entry_frm = tk.Frame(frm, bg="#F4F6F8")
        entry_frm.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        entry_frm.columnconfigure(0, weight=1)

        tk.Entry(entry_frm, textvariable=self.asin_input,
                 font=("Helvetica", 11), width=28).grid(row=0, column=0, sticky="ew")

        self.fetch_btn = tk.Button(entry_frm, text="価格取得",
                                   command=self._fetch_price,
                                   font=("Helvetica", 10, "bold"),
                                   bg="#2980B9", fg="white", relief="flat",
                                   padx=8, pady=2, cursor="hand2")
        self.fetch_btn.grid(row=0, column=1, padx=(6, 0))

        # 商品名表示
        self.title_lbl = tk.Label(frm, textvariable=self.product_title,
                                  font=("Helvetica", 9), bg="#F4F6F8",
                                  fg="#555", wraplength=430, justify="left", anchor="w")
        self.title_lbl.grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

        # 取得状態メッセージ
        self.keepa_status = tk.StringVar(value="")
        tk.Label(frm, textvariable=self.keepa_status, font=("Helvetica", 9),
                 bg="#F4F6F8", fg="#888").grid(row=2, column=0, columnspan=2,
                                               sticky="w", pady=(0, 2))

    def _section_platform(self):
        frm = tk.LabelFrame(self.root, text="プラットフォーム",
                            font=("Helvetica", 10), bg="#F4F6F8", padx=12, pady=6)
        frm.pack(fill="x", padx=20, pady=4)
        for p in PLATFORM_DEFAULTS:
            tk.Radiobutton(frm, text=p, variable=self.platform, value=p,
                           command=self._on_platform, font=("Helvetica", 11),
                           bg="#F4F6F8").pack(side="left", padx=8)

    def _section_input(self):
        frm = tk.LabelFrame(self.root, text="入力項目",
                            font=("Helvetica", 10), bg="#F4F6F8", padx=16, pady=8)
        frm.pack(fill="x", padx=20, pady=4)
        frm.columnconfigure(1, weight=1)

        rows = [
            ("仕入れ値 (円)",   self.purchase,      False),
            ("販売価格 (円)",   self.selling,       False),
            ("手数料率 (%)",    self.referral_rate, True),
            ("FBA配送料 (円)",  self.fba_fee,       True),
            ("送料 (円)",       self.shipping,      True),
        ]
        for i, (label, var, muted) in enumerate(rows):
            fg = "#888" if muted else "#2C3E50"
            tk.Label(frm, text=label, font=("Helvetica", 11), bg="#F4F6F8",
                     fg=fg, anchor="w", width=18).grid(row=i, column=0,
                                                        sticky="w", pady=4)
            tk.Entry(frm, textvariable=var, font=("Helvetica", 12),
                     justify="right", width=14,
                     bg="#FFFFFF" if not muted else "#F0F0F0").grid(
                         row=i, column=1, sticky="e", pady=4, padx=(0, 4))

    def _section_result(self):
        frm = tk.LabelFrame(self.root, text="計算結果",
                            font=("Helvetica", 10), bg="#F4F6F8", padx=16, pady=8)
        frm.pack(fill="x", padx=20, pady=4)
        frm.columnconfigure(1, weight=1)

        self.r_referral = tk.StringVar(value="¥ 0")
        self.r_fba      = tk.StringVar(value="¥ 0")
        self.r_shipping = tk.StringVar(value="¥ 0")
        self.r_profit   = tk.StringVar(value="¥ 0")
        self.r_rate     = tk.StringVar(value="0.0 %")
        self.r_status   = tk.StringVar(value="")

        for i, (lbl, var) in enumerate([
            ("カテゴリー手数料", self.r_referral),
            ("FBA配送料",        self.r_fba),
            ("送料",             self.r_shipping),
        ]):
            tk.Label(frm, text=lbl, font=("Helvetica", 10), bg="#F4F6F8",
                     fg="#666", anchor="w").grid(row=i, column=0, sticky="w", pady=2)
            tk.Label(frm, textvariable=var, font=("Helvetica", 10),
                     bg="#F4F6F8", fg="#555").grid(row=i, column=1, sticky="e", pady=2)

        ttk.Separator(frm, orient="horizontal").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=6)

        tk.Label(frm, text="利益", font=("Helvetica", 13, "bold"),
                 bg="#F4F6F8").grid(row=4, column=0, sticky="w", pady=3)
        self.profit_lbl = tk.Label(frm, textvariable=self.r_profit,
                                   font=("Helvetica", 18, "bold"),
                                   bg="#F4F6F8", fg="#27AE60")
        self.profit_lbl.grid(row=4, column=1, sticky="e", pady=3)

        tk.Label(frm, text="利益率", font=("Helvetica", 13, "bold"),
                 bg="#F4F6F8").grid(row=5, column=0, sticky="w", pady=3)
        tk.Label(frm, textvariable=self.r_rate,
                 font=("Helvetica", 15, "bold"),
                 bg="#F4F6F8", fg="#2980B9").grid(row=5, column=1, sticky="e")

        self.status_lbl = tk.Label(frm, textvariable=self.r_status,
                                   font=("Helvetica", 16, "bold"), bg="#F4F6F8")
        self.status_lbl.grid(row=6, column=0, columnspan=2, pady=(6, 2))

    def _section_buttons(self):
        frm = tk.Frame(self.root, bg="#F4F6F8")
        frm.pack(pady=12)

        tk.Button(frm, text="  Excelに保存  ", command=self._save,
                  font=("Helvetica", 12, "bold"),
                  bg="#27AE60", fg="white", relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=8)

        tk.Button(frm, text="  クリア  ", command=self._clear,
                  font=("Helvetica", 12),
                  bg="#95A5A6", fg="white", relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=8)

    # ── Keepa 連携 ────────────────────────────────────────────────────────────

    def _fetch_price(self):
        if not self.api_key:
            messagebox.showerror("APIキーエラー",
                f"keepa_config.json が見つかりません。\n"
                "api_key を設定してください。")
            return

        asin = extract_asin(self.asin_input.get())
        if not asin:
            messagebox.showwarning("入力エラー",
                "ASINまたはAmazonのURLを入力してください。\n"
                "例: B08XYZ1234  または  https://www.amazon.co.jp/dp/B08XYZ1234")
            return

        self.fetch_btn.config(state="disabled", text="取得中...")
        self.keepa_status.set("Keepaに問い合わせ中...")
        self.product_title.set("")

        def worker():
            try:
                price, title = fetch_keepa_price(self.api_key, asin)
                self.root.after(0, self._on_price_fetched, price, title)
            except Exception as e:
                self.root.after(0, self._on_fetch_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_price_fetched(self, price, title):
        self.fetch_btn.config(state="normal", text="価格取得")
        if price is None:
            self.keepa_status.set("価格が取得できませんでした（在庫なし等）")
            return
        self.selling.set(str(int(price)))
        self.keepa_status.set(f"取得成功: ¥{int(price):,}")
        if title:
            short = title[:50] + ("…" if len(title) > 50 else "")
            self.product_title.set(short)

    def _on_fetch_error(self, err):
        self.fetch_btn.config(state="normal", text="価格取得")
        self.keepa_status.set(f"エラー: {err}")

    # ── ロジック ─────────────────────────────────────────────────────────────

    def _on_platform(self):
        d = PLATFORM_DEFAULTS[self.platform.get()]
        self.referral_rate.set(str(d["referral_rate"]))
        self.fba_fee.set(str(d["fba_fee"]))
        self.shipping.set(str(d["shipping"]))
        self._refresh()

    def _float(self, var, default=0.0):
        try:
            return float(var.get().replace(",", "").strip())
        except ValueError:
            return default

    def _refresh(self, *_):
        purchase = self._float(self.purchase)
        selling  = self._float(self.selling)
        ref_rate = self._float(self.referral_rate, 15.0)
        fba      = self._float(self.fba_fee)
        ship     = self._float(self.shipping)

        referral_fee, profit, profit_rate = calc(
            purchase, selling, ref_rate, fba, ship)

        self.r_referral.set(f"¥ {referral_fee:,.0f}")
        self.r_fba.set(f"¥ {fba:,.0f}")
        self.r_shipping.set(f"¥ {ship:,.0f}")
        self.r_profit.set(f"¥ {profit:,.0f}")
        self.r_rate.set(f"{profit_rate:.1f} %")

        if purchase > 0 and selling > 0:
            color = "#27AE60" if profit > 0 else "#E74C3C"
            self.r_status.set("▲ 黒字" if profit > 0 else "▼ 赤字")
            self.profit_lbl.config(fg=color)
            self.status_lbl.config(fg=color)
        else:
            self.r_status.set("")

    def _clear(self):
        self.purchase.set("")
        self.selling.set("")
        self.asin_input.set("")
        self.product_title.set("")
        self.keepa_status.set("")
        self._on_platform()

    def _save(self):
        if not EXCEL_AVAILABLE:
            messagebox.showerror("エラー",
                "openpyxl がインストールされていません。\n"
                "コマンドプロンプトで:\n  pip install openpyxl\nを実行してください。")
            return

        purchase = self._float(self.purchase)
        selling  = self._float(self.selling)
        if purchase == 0 or selling == 0:
            messagebox.showwarning("入力エラー", "仕入れ値と販売価格を入力してください。")
            return

        ref_rate = self._float(self.referral_rate, 15.0)
        fba      = self._float(self.fba_fee)
        ship     = self._float(self.shipping)
        referral_fee, profit, profit_rate = calc(purchase, selling, ref_rate, fba, ship)

        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "利益計算"
            headers = ["日付", "プラットフォーム", "ASIN", "商品名",
                       "仕入れ値", "販売価格", "手数料率(%)",
                       "カテゴリー手数料", "FBA配送料", "送料", "利益", "利益率(%)"]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2C3E50")
                cell.alignment = Alignment(horizontal="center")
            for i, w in enumerate([12, 16, 12, 30, 12, 12, 11, 14, 12, 10, 12, 11], 1):
                ws.column_dimensions[
                    openpyxl.utils.get_column_letter(i)].width = w

        asin_val = extract_asin(self.asin_input.get()) or ""
        title_val = self.product_title.get()

        row_data = [
            datetime.datetime.now().strftime("%Y/%m/%d"),
            self.platform.get(),
            asin_val, title_val,
            purchase, selling, ref_rate,
            round(referral_fee), round(fba), round(ship),
            round(profit), round(profit_rate, 1),
        ]
        ws.append(row_data)

        last_row = ws.max_row
        profit_cell = ws.cell(row=last_row, column=11)
        if profit > 0:
            profit_cell.fill = PatternFill("solid", fgColor="D5F5E3")
            profit_cell.font = Font(color="1E8449", bold=True)
        else:
            profit_cell.fill = PatternFill("solid", fgColor="FADBD8")
            profit_cell.font = Font(color="C0392B", bold=True)

        wb.save(EXCEL_FILE)
        messagebox.showinfo("保存完了",
            f"Excelに追記しました！\nファイル: {os.path.abspath(EXCEL_FILE)}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
