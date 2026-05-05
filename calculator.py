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

EXCEL_FILE = "せどり_集計.xlsx"
CONFIG_FILE = "keepa_config.json"

PURCHASE_SHOPS = ["楽天", "Yahooショッピング", "その他"]
SALES_CHANNELS = ["Amazon FBA", "Amazon 自己発送", "メルカリ"]

CHANNEL_DEFAULTS = {
    "Amazon FBA":      {"shipping": 0},
    "Amazon 自己発送": {"shipping": 600},
    "メルカリ":         {"shipping": 500},
}

BG = "#F4F6F8"


def load_api_key():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f).get("api_key", "")
    return ""


def extract_asin(text):
    text = text.strip()
    m = re.search(r"/(?:dp|gp/product)/([A-Z0-9]{10})", text)
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Z0-9]{10}", text):
        return text
    return None


def fetch_keepa_price(api_key, asin):
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
    stats = p.get("stats", {})
    current = stats.get("current", [])

    price = None
    for idx in (0, 1):
        if idx < len(current) and current[idx] and current[idx] > 0:
            price = current[idx]
            break

    if price is not None:
        price = price / 100

    return price, title


class OrderSelectDialog(tk.Toplevel):
    """Amazon注文メール選択ダイアログ"""

    def __init__(self, parent, orders):
        super().__init__(parent)
        self.title("Amazon注文メール選択")
        self.geometry("540x300")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.grab_set()
        self.selected = None
        self.orders = orders

        tk.Label(self, text="取得するメールを選択してください",
                 font=("Helvetica", 11, "bold"), bg=BG).pack(pady=(12, 6), padx=15)

        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="both", expand=True, padx=15, pady=4)

        sb = tk.Scrollbar(frame)
        sb.pack(side="right", fill="y")

        self.listbox = tk.Listbox(frame, font=("Helvetica", 10),
                                   yscrollcommand=sb.set, selectmode="single", height=8)
        self.listbox.pack(fill="both", expand=True)
        sb.config(command=self.listbox.yview)

        for order in orders:
            label = f"{order.get('date', '')[:16]}  {order.get('product_name', '(商品名不明)')[:30]}"
            if order.get('amazon_fee'):
                label += f"  手数料:¥{order['amazon_fee']:,}"
            self.listbox.insert(tk.END, label)

        if orders:
            self.listbox.selection_set(0)

        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="選択", command=self._select,
                  font=("Helvetica", 11, "bold"),
                  bg="#27AE60", fg="white", relief="flat",
                  padx=16, pady=6, cursor="hand2").pack(side="left", padx=8)

        tk.Button(btn_frame, text="キャンセル", command=self.destroy,
                  font=("Helvetica", 11),
                  bg="#95A5A6", fg="white", relief="flat",
                  padx=16, pady=6, cursor="hand2").pack(side="left", padx=8)

    def _select(self):
        sel = self.listbox.curselection()
        if sel:
            self.selected = self.orders[sel[0]]
        self.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("せどり 利益計算ツール（Gmail連携）")
        self.root.geometry("580x940")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.api_key = load_api_key()
        self._refreshing = False

        # 仕入れ情報
        today = datetime.date.today().strftime("%Y/%m/%d")
        self.purchase_date = tk.StringVar(value=today)
        self.shop_name     = tk.StringVar(value="楽天")
        self.product_name  = tk.StringVar()
        self.purchase_price = tk.StringVar()
        self.points        = tk.StringVar(value="0")

        # 販売情報
        self.channel       = tk.StringVar(value="Amazon FBA")
        self.sale_date     = tk.StringVar(value=today)
        self.asin_input    = tk.StringVar()
        self.selling_price = tk.StringVar()
        self.outsource     = tk.StringVar(value="0")
        self.shipping      = tk.StringVar(value="0")
        self.fee           = tk.StringVar(value="0")

        # 表示用
        self.r_actual_purchase = tk.StringVar(value="¥ 0")
        self.r_profit  = tk.StringVar(value="¥ 0")
        self.r_roi     = tk.StringVar(value="0.0 %")
        self.r_rate    = tk.StringVar(value="0.0 %")
        self.r_status  = tk.StringVar(value="")
        self.product_title  = tk.StringVar()
        self.keepa_status   = tk.StringVar()

        self._build()

        for v in (self.purchase_price, self.points, self.selling_price,
                  self.outsource, self.shipping, self.fee):
            v.trace_add("write", self._refresh)

        self._on_channel_change()

    # ── UI構築 ───────────────────────────────────────────────────────────────

    def _build(self):
        tk.Label(self.root, text="せどり 利益計算ツール",
                 font=("Helvetica", 17, "bold"),
                 bg=BG, fg="#2C3E50").pack(pady=(14, 4))

        self._section_purchase()
        self._section_sales()
        self._section_result()
        self._section_buttons()

        if not EXCEL_AVAILABLE:
            tk.Label(self.root,
                     text="※ openpyxl 未インストール — pip install openpyxl で Excel保存が使えます",
                     font=("Helvetica", 8), bg=BG, fg="#E74C3C").pack(pady=2)

    def _section_purchase(self):
        frm = tk.LabelFrame(self.root, text="仕入れ情報",
                            font=("Helvetica", 10), bg=BG, padx=14, pady=8)
        frm.pack(fill="x", padx=20, pady=(6, 4))
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(3, weight=1)

        # 購入日 / ショップ
        tk.Label(frm, text="購入日", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=0, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.purchase_date,
                 font=("Helvetica", 11), width=14).grid(row=0, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="ショップ", font=("Helvetica", 11), bg=BG,
                 anchor="w").grid(row=0, column=2, sticky="w", pady=4, padx=(14, 0))
        ttk.Combobox(frm, textvariable=self.shop_name,
                     values=PURCHASE_SHOPS, width=15,
                     state="readonly").grid(row=0, column=3, sticky="ew", pady=4)

        # 商品名
        tk.Label(frm, text="商品名", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=1, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.product_name,
                 font=("Helvetica", 11)).grid(row=1, column=1, columnspan=3,
                                              sticky="ew", pady=4)

        # 仕入れ価格 / ポイント
        tk.Label(frm, text="仕入れ価格 (円)", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=2, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.purchase_price,
                 font=("Helvetica", 12), justify="right",
                 width=14).grid(row=2, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="獲得ポイント", font=("Helvetica", 11), bg=BG,
                 anchor="w").grid(row=2, column=2, sticky="w", pady=4, padx=(14, 0))
        tk.Entry(frm, textvariable=self.points,
                 font=("Helvetica", 12), justify="right",
                 width=14).grid(row=2, column=3, sticky="ew", pady=4)

        # 実質仕入れ価格（自動計算）
        tk.Label(frm, text="実質仕入れ価格", font=("Helvetica", 11), bg=BG,
                 fg="#666", anchor="w", width=14).grid(row=3, column=0, sticky="w", pady=4)
        tk.Label(frm, textvariable=self.r_actual_purchase,
                 font=("Helvetica", 13, "bold"), bg=BG,
                 fg="#2980B9").grid(row=3, column=1, columnspan=3, sticky="e", pady=4)

    def _section_sales(self):
        frm = tk.LabelFrame(self.root, text="販売情報",
                            font=("Helvetica", 10), bg=BG, padx=14, pady=8)
        frm.pack(fill="x", padx=20, pady=4)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(3, weight=1)

        # 販路
        tk.Label(frm, text="販路", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=0, column=0, sticky="w", pady=4)
        ch_frm = tk.Frame(frm, bg=BG)
        ch_frm.grid(row=0, column=1, columnspan=3, sticky="w")
        for ch in SALES_CHANNELS:
            tk.Radiobutton(ch_frm, text=ch, variable=self.channel, value=ch,
                           command=self._on_channel_change,
                           font=("Helvetica", 11), bg=BG).pack(side="left", padx=6)

        # 販売日
        tk.Label(frm, text="販売日", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=1, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.sale_date,
                 font=("Helvetica", 11), width=14).grid(row=1, column=1, sticky="ew", pady=4)

        # ASIN / URL（Keepa）
        tk.Label(frm, text="ASIN / URL", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=2, column=0, sticky="w", pady=4)
        asin_frm = tk.Frame(frm, bg=BG)
        asin_frm.grid(row=2, column=1, columnspan=3, sticky="ew")
        asin_frm.columnconfigure(0, weight=1)

        tk.Entry(asin_frm, textvariable=self.asin_input,
                 font=("Helvetica", 11)).grid(row=0, column=0, sticky="ew")
        self.keepa_btn = tk.Button(asin_frm, text="Keepa価格取得",
                                   command=self._fetch_keepa,
                                   font=("Helvetica", 10, "bold"),
                                   bg="#2980B9", fg="white", relief="flat",
                                   padx=8, pady=2, cursor="hand2")
        self.keepa_btn.grid(row=0, column=1, padx=(6, 0))

        tk.Label(frm, textvariable=self.keepa_status,
                 font=("Helvetica", 9), bg=BG,
                 fg="#888").grid(row=3, column=0, columnspan=4, sticky="w")
        tk.Label(frm, textvariable=self.product_title,
                 font=("Helvetica", 9), bg=BG, fg="#555",
                 wraplength=520, justify="left",
                 anchor="w").grid(row=4, column=0, columnspan=4, sticky="w")

        # 販売価格 / 外注
        tk.Label(frm, text="販売価格 (円)", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=5, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.selling_price,
                 font=("Helvetica", 12), justify="right",
                 width=14).grid(row=5, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="外注 (円)", font=("Helvetica", 11), bg=BG,
                 anchor="w").grid(row=5, column=2, sticky="w", pady=4, padx=(14, 0))
        tk.Entry(frm, textvariable=self.outsource,
                 font=("Helvetica", 12), justify="right",
                 width=14).grid(row=5, column=3, sticky="ew", pady=4)

        # 送料 / 手数料
        tk.Label(frm, text="送料 (円)", font=("Helvetica", 11), bg=BG,
                 anchor="w", width=14).grid(row=6, column=0, sticky="w", pady=4)
        tk.Entry(frm, textvariable=self.shipping,
                 font=("Helvetica", 12), justify="right",
                 width=14).grid(row=6, column=1, sticky="ew", pady=4)

        tk.Label(frm, text="手数料 (円)", font=("Helvetica", 11), bg=BG,
                 anchor="w").grid(row=6, column=2, sticky="w", pady=4, padx=(14, 0))

        fee_frm = tk.Frame(frm, bg=BG)
        fee_frm.grid(row=6, column=3, sticky="ew", pady=4)
        fee_frm.columnconfigure(0, weight=1)

        self.fee_entry = tk.Entry(fee_frm, textvariable=self.fee,
                                   font=("Helvetica", 12), justify="right")
        self.fee_entry.grid(row=0, column=0, sticky="ew")

        self.gmail_btn = tk.Button(fee_frm, text="Gmail取得",
                                   command=self._fetch_gmail,
                                   font=("Helvetica", 10, "bold"),
                                   bg="#C0392B", fg="white", relief="flat",
                                   padx=6, pady=2, cursor="hand2")
        self.gmail_btn.grid(row=0, column=1, padx=(5, 0))

        self.fee_note = tk.Label(frm, text="",
                                  font=("Helvetica", 8), bg=BG, fg="#888")
        self.fee_note.grid(row=7, column=0, columnspan=4, sticky="w", pady=(0, 2))

    def _section_result(self):
        frm = tk.LabelFrame(self.root, text="計算結果",
                            font=("Helvetica", 10), bg=BG, padx=16, pady=8)
        frm.pack(fill="x", padx=20, pady=4)
        frm.columnconfigure(1, weight=1)

        items = [
            ("利益額",                    self.r_profit, 18, "#27AE60"),
            ("ROI  (利益額 ÷ 実質仕入れ)", self.r_roi,    13, "#8E44AD"),
            ("利益率 (利益額 ÷ 販売価格)", self.r_rate,   13, "#2980B9"),
        ]
        for i, (label, var, size, color) in enumerate(items):
            tk.Label(frm, text=label, font=("Helvetica", 11),
                     bg=BG, anchor="w").grid(row=i, column=0, sticky="w", pady=4)
            lbl = tk.Label(frm, textvariable=var,
                           font=("Helvetica", size, "bold"),
                           bg=BG, fg=color)
            lbl.grid(row=i, column=1, sticky="e", pady=4)
            if i == 0:
                self.profit_lbl = lbl

        self.status_lbl = tk.Label(frm, textvariable=self.r_status,
                                    font=("Helvetica", 16, "bold"), bg=BG)
        self.status_lbl.grid(row=3, column=0, columnspan=2, pady=(4, 2))

    def _section_buttons(self):
        frm = tk.Frame(self.root, bg=BG)
        frm.pack(pady=14)

        tk.Button(frm, text="  Excelに保存  ", command=self._save,
                  font=("Helvetica", 12, "bold"),
                  bg="#27AE60", fg="white", relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=8)

        tk.Button(frm, text="  クリア  ", command=self._clear,
                  font=("Helvetica", 12),
                  bg="#95A5A6", fg="white", relief="flat",
                  padx=16, pady=8, cursor="hand2").pack(side="left", padx=8)

    # ── 販路切替 ─────────────────────────────────────────────────────────────

    def _on_channel_change(self):
        ch = self.channel.get()
        self.shipping.set(str(CHANNEL_DEFAULTS[ch]["shipping"]))

        is_mercari = (ch == "メルカリ")
        self.fee_entry.config(
            state="readonly" if is_mercari else "normal",
            bg="#F0F0F0" if is_mercari else "#FFFFFF",
        )
        self.gmail_btn.config(state="disabled" if is_mercari else "normal")

        if is_mercari:
            self.fee_note.config(
                text="※ メルカリは販売価格の10%を自動計算", fg="#888")
        else:
            self.fee_note.config(
                text="※ Amazon手数料は「Gmail取得」で自動入力できます", fg="#666")

        self._refresh()

    # ── Keepa連携 ────────────────────────────────────────────────────────────

    def _fetch_keepa(self):
        if not self.api_key:
            messagebox.showerror("APIキーエラー",
                "keepa_config.json が見つかりません。\napi_key を設定してください。")
            return

        asin = extract_asin(self.asin_input.get())
        if not asin:
            messagebox.showwarning("入力エラー",
                "ASINまたはAmazonのURLを入力してください。")
            return

        self.keepa_btn.config(state="disabled", text="取得中...")
        self.keepa_status.set("Keepaに問い合わせ中...")
        self.product_title.set("")

        def worker():
            try:
                price, title = fetch_keepa_price(self.api_key, asin)
                self.root.after(0, self._on_keepa_result, price, title)
            except Exception as e:
                self.root.after(0, self._on_keepa_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_keepa_result(self, price, title):
        self.keepa_btn.config(state="normal", text="Keepa価格取得")
        if price is None:
            self.keepa_status.set("価格を取得できませんでした（在庫なし等）")
            return
        self.selling_price.set(str(int(price)))
        self.keepa_status.set(f"取得成功: ¥{int(price):,}")
        if title:
            self.product_title.set(title[:50] + ("…" if len(title) > 50 else ""))
            if not self.product_name.get():
                self.product_name.set(title[:100])

    def _on_keepa_error(self, err):
        self.keepa_btn.config(state="normal", text="Keepa価格取得")
        self.keepa_status.set(f"エラー: {err}")

    # ── Gmail連携 ────────────────────────────────────────────────────────────

    def _fetch_gmail(self):
        self.gmail_btn.config(state="disabled", text="読込中...")

        def worker():
            try:
                from gmail_reader import fetch_amazon_orders
                orders = fetch_amazon_orders(max_results=5)
                self.root.after(0, self._on_gmail_result, orders, None)
            except Exception as e:
                self.root.after(0, self._on_gmail_result, None, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_gmail_result(self, orders, error):
        self.gmail_btn.config(state="normal", text="Gmail取得")

        if error:
            messagebox.showerror("Gmail エラー", error)
            return

        if not orders:
            messagebox.showinfo("情報", "Amazon注文確定メールが見つかりませんでした。")
            return

        if len(orders) == 1:
            self._apply_order(orders[0])
        else:
            dlg = OrderSelectDialog(self.root, orders)
            self.root.wait_window(dlg)
            if dlg.selected:
                self._apply_order(dlg.selected)

    def _apply_order(self, order):
        applied = []

        if order.get('amazon_fee'):
            self.fee.set(str(order['amazon_fee']))
            applied.append(f"手数料: ¥{order['amazon_fee']:,}")

        if order.get('selling_price') and not self.selling_price.get():
            self.selling_price.set(str(order['selling_price']))
            applied.append(f"販売価格: ¥{order['selling_price']:,}")

        if order.get('product_name') and not self.product_name.get():
            self.product_name.set(order['product_name'])
            applied.append(f"商品名: {order['product_name'][:30]}")

        self._refresh()

        if applied:
            messagebox.showinfo("Gmail取得完了",
                "\n".join(applied) + "\nを入力しました。\n確認・修正後「Excelに保存」してください。")

    # ── 計算 ─────────────────────────────────────────────────────────────────

    def _flt(self, var, default=0.0):
        try:
            return float(var.get().replace(",", "").strip())
        except (ValueError, AttributeError):
            return default

    def _refresh(self, *_):
        if self._refreshing:
            return
        self._refreshing = True
        try:
            purchase = self._flt(self.purchase_price)
            points   = self._flt(self.points)
            actual   = purchase - points
            selling  = self._flt(self.selling_price)
            outsource = self._flt(self.outsource)
            ship     = self._flt(self.shipping)

            # メルカリは10%自動計算
            if self.channel.get() == "メルカリ" and selling > 0:
                self.fee.set(str(round(selling * 0.10)))

            fee = self._flt(self.fee)

            self.r_actual_purchase.set(f"¥ {actual:,.0f}")

            if purchase > 0 and selling > 0:
                profit = selling - actual - outsource - ship - fee
                roi    = (profit / actual * 100) if actual > 0 else 0
                rate   = (profit / selling * 100) if selling > 0 else 0

                self.r_profit.set(f"¥ {profit:,.0f}")
                self.r_roi.set(f"{roi:.1f} %")
                self.r_rate.set(f"{rate:.1f} %")

                color = "#27AE60" if profit > 0 else "#E74C3C"
                self.r_status.set("▲ 黒字" if profit > 0 else "▼ 赤字")
                self.profit_lbl.config(fg=color)
                self.status_lbl.config(fg=color)
            else:
                self.r_profit.set("¥ 0")
                self.r_roi.set("0.0 %")
                self.r_rate.set("0.0 %")
                self.r_status.set("")
        finally:
            self._refreshing = False

    # ── クリア ───────────────────────────────────────────────────────────────

    def _clear(self):
        today = datetime.date.today().strftime("%Y/%m/%d")
        self.purchase_date.set(today)
        self.shop_name.set("楽天")
        self.product_name.set("")
        self.purchase_price.set("")
        self.points.set("0")
        self.channel.set("Amazon FBA")
        self.sale_date.set(today)
        self.asin_input.set("")
        self.product_title.set("")
        self.keepa_status.set("")
        self.selling_price.set("")
        self.outsource.set("0")
        self.fee.set("0")
        self._on_channel_change()

    # ── Excelに保存 ──────────────────────────────────────────────────────────

    def _save(self):
        if not EXCEL_AVAILABLE:
            messagebox.showerror("エラー",
                "openpyxl が未インストールです。\npip install openpyxl")
            return

        purchase = self._flt(self.purchase_price)
        selling  = self._flt(self.selling_price)

        if purchase == 0 or selling == 0:
            messagebox.showwarning("入力エラー", "仕入れ価格と販売価格を入力してください。")
            return

        points    = self._flt(self.points)
        actual    = purchase - points
        outsource = self._flt(self.outsource)
        ship      = self._flt(self.shipping)
        fee       = self._flt(self.fee)
        profit    = selling - actual - outsource - ship - fee
        roi       = round(profit / actual * 100, 1) if actual > 0 else 0.0
        rate      = round(profit / selling * 100, 1) if selling > 0 else 0.0

        headers = [
            "購入日", "ショップ名", "商品名",
            "仕入れ価格\n(クーポン込み)", "獲得ポイント", "実質仕入れ価格",
            "販路", "販売日", "販売価格",
            "外注", "送料", "手数料",
            "利益額", "ROI\n(利益額/仕入れ価格)", "利益率\n(利益額/販売価格)",
        ]

        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "せどり集計"
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2C3E50")
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 36
            for i, w in enumerate(
                [12, 16, 28, 14, 12, 14, 14, 12, 12, 10, 10, 12, 12, 14, 14], 1
            ):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        ws.append([
            self.purchase_date.get(),
            self.shop_name.get(),
            self.product_name.get(),
            round(purchase),
            round(points),
            round(actual),
            self.channel.get(),
            self.sale_date.get(),
            round(selling),
            round(outsource),
            round(ship),
            round(fee),
            round(profit),
            roi,
            rate,
        ])

        last = ws.max_row
        profit_cell = ws.cell(row=last, column=13)
        pct_fmt = '0.0"%"'
        ws.cell(row=last, column=14).number_format = pct_fmt  # ROI
        ws.cell(row=last, column=15).number_format = pct_fmt  # 利益率

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
