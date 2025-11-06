"""
send_schedule_fullapp.py
Ứng dụng desktop gửi mail lịch thi theo file Excel.

Chức năng:
- Chọn file Excel (.xlsx/.xls)
- Xem trước dữ liệu (bảng)
- Thống kê: tổng số giảng viên, tổng số môn, email hợp lệ
- Xem trước mẫu email (HTML)
- Gộp lịch theo Email + Giảng viên; bảng lịch gồm: Ngành, Môn thi, Lớp, Ngày thi, Giờ thi
- Gửi mail từng giảng viên, hiển thị trạng thái (Đang gửi / Thành công / Lỗi) và progress bar
"""

import os
import re
import threading
import queue
import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---- Load env ----
load_dotenv()
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# ---- Email validation ----
EMAIL_REGEX = re.compile(r"^[^@]+@[^@]+\.[^@]+$")

def is_valid_email(e):
    return bool(EMAIL_REGEX.match(str(e).strip()))

# ---- Send email function ----
def send_email(to_email, subject, html_content):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(html_content, "html"))
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

# ---- Tạo bảng dữ liệu cho email ----
def build_html_table(group_df):
    # Giữ các cột theo thứ tự được yêu cầu
    cols = ["Nganh", "Hoc_phan", "Lop", "Ngay_thi", "Gio_thi"]
    # If some columns missing, try to fallback using available ones
    out_df = group_df.copy()
    # Ensure all desired columns exist
    for c in cols:
        if c not in out_df.columns:
            out_df[c] = ""
    out_df = out_df[cols]
    # Chuyển định dạng ngày
    out_df["Ngay_thi"] = pd.to_datetime(out_df["Ngay_thi"]).dt.strftime("%d/%m/%Y")
    # to_html for nice table
    html_table = out_df.to_html(index=False, border=1, justify="center")
    return html_table

# ---- GUI App ----
class SendScheduleApp:
    def __init__(self, root):
        self.root = root
        root.title("📧 Gửi lịch thi - Ứng dụng hoàn chỉnh")
        root.geometry("1100x800")

        # Top frame: file selection + stats
        top = ttk.Frame(root)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Chọn file Excel:").pack(side="left")
        self.filevar = tk.StringVar()
        self.entry_file = ttk.Entry(top, textvariable=self.filevar, width=70)
        self.entry_file.pack(side="left", padx=6)
        ttk.Button(top, text="Chọn file...", command=self.choose_file).pack(side="left", padx=6)
        ttk.Button(top, text="Tải lại dữ liệu", command=self.load_file).pack(side="left", padx=6)

        # Left: preview dataframe
        left_frame = ttk.Frame(root)
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=6)

        # ====== KHUNG XEM TRƯỚC FILE EXCEL ======
        frame_preview = ttk.LabelFrame(self.root, text="📋 Xem trước dữ liệu Excel")
        frame_preview.pack(padx=10, pady=10, fill="both", expand=False)

        # Tạo frame chứa bảng và thanh cuộn
        table_container = ttk.Frame(frame_preview)
        table_container.pack(fill="both", expand=True)

        # Thanh cuộn dọc
        self.scrollbar_y = ttk.Scrollbar(table_container, orient="vertical")
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Thanh cuộn ngang
        self.scrollbar_x = ttk.Scrollbar(table_container, orient="horizontal")
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview hiển thị dữ liệu Excel
        self.tree = ttk.Treeview(
            table_container,
            columns=("Email", "Giang_vien", "Nganh", "Lop", "Hoc_phan", "Ngay_thi", "Gio_thi"),
            show="headings",
            yscrollcommand=self.scrollbar_y.set,
            xscrollcommand=self.scrollbar_x.set,
            height=8  # 👈 Giới hạn hiển thị 8 dòng để tránh tràn
        )

        self.tree.pack(fill="both", expand=True)

        # Gán thanh cuộn
        self.scrollbar_y.config(command=self.tree.yview)
        self.scrollbar_x.config(command=self.tree.xview)
        
        # ======PHẦN THỐNG KÊ NHANH ======
        # Tiêu đề cột
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        
        # Right: controls, stats, preview email & status
        right = ttk.Frame(root, width=420)
        right.pack(side="right", fill="y", padx=10, pady=6)

        # ======Thống kê======
        stats_frame = ttk.LabelFrame(right, text="Thống kê nhanh")
        stats_frame.pack(fill="x", pady=6)
        self.lbl_total_gv = ttk.Label(stats_frame, text="Tổng số giảng viên: 0")
        self.lbl_total_gv.pack(anchor="w", padx=6, pady=2)
        self.lbl_total_mon = ttk.Label(stats_frame, text="Tổng số học phần (dòng): 0")
        self.lbl_total_mon.pack(anchor="w", padx=6, pady=2)
        self.lbl_valid_emails = ttk.Label(stats_frame, text="Email hợp lệ: 0")
        self.lbl_valid_emails.pack(anchor="w", padx=6, pady=2)

        # ====== Xem trước mẫu email  ======
        preview_frame = ttk.LabelFrame(right, text="Xem trước mẫu email")
        preview_frame.pack(fill="both", expand=True, pady=6)
        ttk.Label(preview_frame, text="Chủ đề:").pack(anchor="w", padx=6, pady=(6,0))
        self.subject_var = tk.StringVar(value="Lịch thi các học phần - {GV}")
        ttk.Entry(preview_frame, textvariable=self.subject_var, width=50).pack(padx=6, pady=(0,6))
        ttk.Label(preview_frame, text="Mẫu nội dung (HTML) - sẽ chèn vào bảng lịch bên dưới:").pack(anchor="w", padx=6)
        self.text_preview = tk.Text(preview_frame, height=12, wrap="word")
        default_body = ("<p>Kính gửi Thầy/Cô <b>{GV}</b>,</p>"
                        "<p>Dưới đây là lịch thi các học phần do Thầy/Cô phụ trách:</p>"
                        "{TABLE}"
                        "<p>Trân trọng,<br>Phòng Khảo thí</p>")
        self.text_preview.insert("1.0", default_body)
        self.text_preview.pack(fill="both", expand=True, padx=6, pady=6)

        # Send controls
        send_frame = ttk.LabelFrame(right, text="Gửi mail")
        send_frame.pack(fill="x", pady=6)
        ttk.Button(send_frame, text="Gửi cho tất cả", command=self.confirm_and_send).pack(fill="x", padx=6, pady=6)

        self.progress = ttk.Progressbar(send_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=(0,6))

        # ====== BẢNG TRẠNG THÁI GỬI ======
        status_frame = ttk.LabelFrame(root, text="Trạng thái gửi")
        status_frame.pack(fill="both", padx=10, pady=(0,10), expand=True)
        cols = ("Trạng thái", "Email", "Giảng viên", "Số lớp" )
        self.status_tree = ttk.Treeview(status_frame, columns=cols, show="headings", height=8)
        for c in cols:
            self.status_tree.heading(c, text=c)
            self.status_tree.column(c, anchor="center")
        self.status_tree.pack(fill="both", expand=True, side="left")
        status_v = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_tree.yview)
        self.status_tree.configure(yscroll=status_v.set)
        status_v.pack(side="right", fill="y")

                # ===== Phân luồng dữ liệu mail để gửi =====
        self.df = None
        self.grouped = None
        self.send_queue = queue.Queue()
        self.sending_thread = None

    def choose_file(self):
        path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.filevar.set(path)
            self.load_file()

    def load_file(self):
        path = self.filevar.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Thiếu file", "Hãy chọn file Excel hợp lệ.")
            return
        try:
            # read with pandas
            df = pd.read_excel(path, dtype=str, engine="openpyxl")
            # strip column names
            df.columns = [c.strip() for c in df.columns]
            # fill NaN
            df = df.fillna("")
            self.df = df
            self.populate_preview(df)
            self.update_stats(df)
            self.prepare_groups(df)
        except Exception as e:
            messagebox.showerror("Lỗi đọc file", str(e))

    def populate_preview(self, df):
        # clear tree
        for col in self.tree["columns"]:
            self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        # set headings
        for c in df.columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="center")
        # insert some rows (limiting to 200 for preview)
        for i, row in df.head(200).iterrows():
            vals = [str(row[c]) for c in df.columns]
            self.tree.insert("", "end", values=vals)

    def update_stats(self, df):
        # total giang vien unique by "Giang_vien" or by Email
        if "Giang_vien" in df.columns:
            total_gv = df["Giang_vien"].nunique()
        else:
            total_gv = df["Email"].nunique() if "Email" in df.columns else 0
        total_rows = len(df)
        valid_emails = df["Email"].apply(is_valid_email).sum() if "Email" in df.columns else 0
        self.lbl_total_gv.config(text=f"Tổng số giảng viên: {total_gv}")
        self.lbl_total_mon.config(text=f"Tổng số học phần (dòng): {total_rows}")
        self.lbl_valid_emails.config(text=f"Email hợp lệ: {valid_emails}")

    def prepare_groups(self, df):
    # Đảm bảo tồn tại các cột cần thiết
        required = ["Email", "Giang_vien", "Nganh", "Hoc_phan", "Lop", "Ngay_thi", "Gio_thi"]
        for c in required:
            if c not in df.columns:
                df[c] = ""
        # NHÓM LỊCH LỚP THEO GIẢNG VIÊN
        grouped = df.groupby(["Email", "Giang_vien"], sort=False)
        self.grouped = grouped

        # Làm trống bảng trạng thái
        for i in self.status_tree.get_children():
            self.status_tree.delete(i)

        # Ghi nhận từng giảng viên vào bảng trạng thái
        for (email, gv), group in grouped:
            count = len(group)
            self.status_tree.insert("", "end", values=("🕓 Chưa gửi", email, gv, count))


    def confirm_and_send(self):
        if self.df is None:
            messagebox.showwarning("Thiếu dữ liệu", "Hãy chọn file Excel trước.")
            return
        if EMAIL_USER is None or EMAIL_PASS is None:
            messagebox.showerror("Thiếu cấu hình", "Thiếu EMAIL_USER hoặc EMAIL_PASS trong file .env.")
            return

        # Xác nhận gửi
        if not messagebox.askyesno("Xác nhận", "Bạn có chắc muốn gửi mail cho tất cả giảng viên?"):
            return

        # Chuẩn bị dữ liệu gửi
        items = []
        for (email, gv), group in self.grouped:
            items.append((email, gv, group.copy()))
        if not items:
            messagebox.showinfo("Không có người nhận", "Không tìm thấy người nhận hợp lệ.")
            return

        # Thiết lập tiến trình
        total = len(items)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        # Đặt trạng thái tất cả giảng viên thành “Sẵn sàng”
        for iid in self.status_tree.get_children():
            vals = self.status_tree.item(iid, "values")
            self.status_tree.set(iid, column="Trạng thái", value="🟡 Sẵn sàng")

        # Bắt đầu luồng gửi mail
        self.sending_thread = threading.Thread(target=self._sending_worker, args=(items,), daemon=True)
        self.sending_thread.start()
        self.root.after(200, self._process_queue)


    def _sending_worker(self, items):
        for idx, (email, gv, group_df) in enumerate(items, start=1):
            # Cập nhật trạng thái đang gửi
            self.send_queue.put(("update_status", email, gv, "🟡 Đang gửi..."))

            # Tạo nội dung email
            table_html = build_html_table(group_df)
            body_template = self.text_preview.get("1.0", "end").strip()
            body_html = body_template.replace("{GV}", gv).replace("{TABLE}", table_html)
            subject = self.subject_var.get().strip().replace("{GV}", gv)

            # Gửi mail và xử lý kết quả
            ok, err = send_email(email, subject, body_html) if is_valid_email(email) else (False, "Email không hợp lệ")

            if ok:
                self.send_queue.put(("update_status", email, gv, "🟢 Thành công"))
            else:
                self.send_queue.put(("update_status", email, gv, f"🔴 Lỗi: {err}"))

            # Cập nhật tiến trình
            self.send_queue.put(("progress", idx))

        # Khi hoàn tất
        self.send_queue.put(("done", None))


    def _process_queue(self):
        try:
            while True:
                item = self.send_queue.get_nowait()

                # Cập nhật trạng thái từng giảng viên
                if item[0] == "update_status":
                    _, email, gv, status = item
                    for iid in self.status_tree.get_children():
                        vals = self.status_tree.item(iid, "values")

                        # ⚠️ Sửa logic tìm đúng cột Email và Giảng viên
                        if vals[1] == email and vals[2] == gv:
                            self.status_tree.set(iid, column="Trạng thái", value=status)
                            break

                # Cập nhật tiến trình
                elif item[0] == "progress":
                    _, val = item
                    self.progress["value"] = val

                # Hoàn tất gửi
                elif item[0] == "done":
                    messagebox.showinfo("Hoàn tất", "Quá trình gửi đã kết thúc.")

        except queue.Empty:
            pass

        # Tiếp tục lặp lại nếu luồng gửi vẫn chạy
        if self.sending_thread and self.sending_thread.is_alive():
            self.root.after(200, self._process_queue)
        else:
            self.progress["value"] = self.progress["maximum"]

       

# ---- Run app ----
def main():
    root = tk.Tk()
    app = SendScheduleApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
