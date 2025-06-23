import tkinter as tk
from tkinter import messagebox, ttk
import sqlite3
import smtplib
from email.message import EmailMessage

# ===== 전역 변수 선언 =====
student_management_tree = None
student_management_checked_items = {}
student_management_db_path = "teacher_db.db"
student_management_all_checked = [False]
student_management_id_map = {}  # item_id -> DB id 매핑용

def student_management_init_db():
    conn = sqlite3.connect(student_management_db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_number INTEGER NOT NULL,
            name TEXT NOT NULL,
            seat_number INTEGER NOT NULL,
            email TEXT
        )
    ''')
    conn.commit()
    conn.close()

def student_management_register_student(entry_class, entry_name, entry_seat, entry_email):
    class_num = entry_class.get()
    name = entry_name.get()
    seat_num = entry_seat.get()
    email = entry_email.get()

    if not class_num or not name or not seat_num or not email:
        messagebox.showwarning("입력 오류", "모든 필드를 입력하세요.")
        return

    try:
        class_num = int(class_num)
        seat_num = int(seat_num)
    except ValueError:
        messagebox.showerror("형식 오류", "반 번호와 좌석 번호는 숫자여야 합니다.")
        return

    conn = sqlite3.connect(student_management_db_path)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO students (class_number, name, seat_number, email) VALUES (?, ?, ?, ?)",
                   (class_num, name, seat_num, email))
    conn.commit()
    conn.close()

    entry_class.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_seat.delete(0, tk.END)
    entry_email.delete(0, tk.END)

    student_management_load_students()

def student_management_load_students(filter_class=None):
    global student_management_tree, student_management_checked_items, student_management_id_map

    for row in student_management_tree.get_children():
        student_management_tree.delete(row)
    student_management_checked_items.clear()
    student_management_id_map.clear()

    conn = sqlite3.connect(student_management_db_path)
    cursor = conn.cursor()
    if filter_class is None:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students")
    else:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students WHERE class_number = ?", (filter_class,))
    rows = cursor.fetchall()
    conn.close()

    for row in rows:
        db_id, class_num, seat, name, email = row
        item_id = student_management_tree.insert("", tk.END, values=("☐", class_num, seat, name, email))
        student_management_checked_items[item_id] = False
        student_management_id_map[item_id] = db_id

def student_management_get_class_numbers():
    conn = sqlite3.connect(student_management_db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT class_number FROM students ORDER BY class_number")
    classes = [row[0] for row in cursor.fetchall()]
    conn.close()
    return classes

def student_management_show_class_menu(event):
    menu = tk.Menu(student_management_tree, tearoff=0)
    menu.add_command(label="전체 보기", command=lambda: student_management_load_students())
    for class_num in student_management_get_class_numbers():
        menu.add_command(label=f"{class_num}반 보기", command=lambda c=class_num: student_management_load_students(c))
    menu.post(event.x_root, event.y_root)

def student_management_edit_cell(event):
    row_id = student_management_tree.identify_row(event.y)
    column = student_management_tree.identify_column(event.x)
    if not row_id or not column:
        return

    col_index = int(column.replace("#", ""))
    if col_index < 2:
        return

    x, y, width, height = student_management_tree.bbox(row_id, column)
    value = student_management_tree.set(row_id, column)
    col_name = student_management_tree.heading(column)["text"]

    entry = tk.Entry(student_management_tree)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, value)
    entry.focus()

    def save_edit(event_inner):
        new_value = entry.get().strip()
        if not new_value:
            messagebox.showwarning("입력 오류", f"{col_name}은(는) 빈 값일 수 없습니다.")
            entry.focus()
            return

        if column in ("#2", "#3") and not new_value.isdigit():
            messagebox.showerror("형식 오류", f"{col_name}은(는) 숫자여야 합니다.")
            entry.focus()
            return

        entry.destroy()
        student_management_tree.set(row_id, column, new_value)

        db_id = student_management_id_map[row_id]
        col_map = {
            "#2": "class_number",
            "#3": "seat_number",
            "#4": "name",
            "#5": "email"
        }
        db_col = col_map.get(column)
        if db_col:
            conn = sqlite3.connect(student_management_db_path)
            cursor = conn.cursor()
            cursor.execute(f"UPDATE students SET {db_col} = ? WHERE id = ?", (new_value, db_id))
            conn.commit()
            conn.close()
            student_management_load_students()

    entry.bind("<Return>", save_edit)
    entry.bind("<FocusOut>", lambda e: entry.destroy())

def student_management_on_tree_click(event):
    region = student_management_tree.identify_region(event.x, event.y)
    column = student_management_tree.identify_column(event.x)
    row = student_management_tree.identify_row(event.y)

    if column == "#1":
        if region == "cell" and row:
            student_management_checked_items[row] = not student_management_checked_items.get(row, False)
            student_management_tree.set(row, column="#1", value="☑" if student_management_checked_items[row] else "☐")
        elif region == "heading":
            student_management_all_checked[0] = not student_management_all_checked[0]
            for item_id in student_management_tree.get_children():
                student_management_checked_items[item_id] = student_management_all_checked[0]
                student_management_tree.set(item_id, column="#1", value="☑" if student_management_all_checked[0] else "☐")
    elif region == "heading" and column == "#2":
        student_management_show_class_menu(event)

def student_management_delete_selected():
    conn = sqlite3.connect(student_management_db_path)
    cursor = conn.cursor()
    for item_id, checked in student_management_checked_items.items():
        if checked:
            db_id = student_management_id_map[item_id]
            cursor.execute("DELETE FROM students WHERE id = ?", (db_id,))
    conn.commit()
    conn.close()
    student_management_load_students()

# =========================================== 학생 관리 윈도우 ===========================================
def student_management():
    global student_management_tree

    student_management_init_db()

    window = tk.Tk()
    window.title("선생님 프로그램")

    window.grid_columnconfigure(0, weight=1)
    window.grid_rowconfigure(2, weight=1)

    # =========================================== 네비게이션 프레임 ===========================================
    nav_frame = tk.Frame(window, bg="#f0f0f0")
    nav_frame.grid(row=0, column=0, sticky="ew")
    nav_frame.grid_columnconfigure(1, weight=1)

    btn_home = tk.Button(nav_frame, text="← 홈으로", relief="flat")
    btn_home.grid(row=0, column=0, padx=10, pady=5)

    lbl_title = tk.Label(nav_frame, text="학생 관리", font=("맑은 고딕", 16, "bold"), bg="#f0f0f0")
    lbl_title.grid(row=0, column=1, sticky="w")

    # =========================================== 상단 프레임 ===========================================
    top_frame = tk.LabelFrame(window, text="학생 입력")
    top_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    # top_frame의 열 확장 막기
    # 열 크기 고정: 버튼이나 엔트리가 왼쪽으로만 정렬되도록 함
    top_frame.grid_columnconfigure(0, weight=0)
    top_frame.grid_columnconfigure(1, weight=0)    

    # ===== 학생 등록 버튼 =====
    btn_register = tk.Button(top_frame, text="학생 등록", width=10, height=2,
                             command=lambda: student_management_register_student(entry_class, entry_name, entry_seat, entry_email))
    btn_register.grid(row=0, column=0, padx=5, pady=10, sticky="w")

    # ===== 입력 필드 =====
    tk.Label(top_frame, text="반 번호", anchor="w").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    entry_class = tk.Entry(top_frame, width=20)
    entry_class.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(top_frame, text="좌석 번호").grid(row=2, column=0, padx=5, pady=5, sticky="w")
    entry_seat = tk.Entry(top_frame, width=20)
    entry_seat.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Label(top_frame, text="이름").grid(row=3, column=0, padx=5, pady=5, sticky="w")
    entry_name = tk.Entry(top_frame, width=20)
    entry_name.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    tk.Label(top_frame, text="이메일").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    entry_email = tk.Entry(top_frame, width=20)
    entry_email.grid(row=4, column=1, padx=5, pady=5, sticky="w")

    # =========================================== 중간 프레임 ===========================================
    middle_frame = tk.LabelFrame(window, text="학생 관리")
    middle_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

    # ===== 선택 삭제 버튼 =====
    btn_delete = tk.Button(middle_frame, text="선택된 학생 삭제", height=2,
                           command=student_management_delete_selected)
    btn_delete.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # ===== 트리뷰 =====
    tree_frame = tk.Frame(middle_frame)
    tree_frame.grid(row=1, column=0, sticky="nsew")

    middle_frame.grid_rowconfigure(1, weight=1)
    middle_frame.grid_columnconfigure(0, weight=1)

    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
    scrollbar.grid(row=0, column=1, sticky="ns")

    student_management_tree = ttk.Treeview(
        tree_frame,
        columns=("check", "class", "seat", "name", "email"),
        show="headings",
        yscrollcommand=scrollbar.set
    )
    student_management_tree.grid(row=0, column=0, sticky="nsew")
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)

    scrollbar.config(command=student_management_tree.yview)

    student_management_tree.heading("check", text="선택")
    student_management_tree.column("check", width=40, anchor='center', stretch=False)

    student_management_tree.heading("class", text="반 번호")
    student_management_tree.column("class", width=60, anchor='center', stretch=False)

    student_management_tree.heading("seat", text="좌석 번호")
    student_management_tree.column("seat", width=60, anchor='center', stretch=False)

    student_management_tree.heading("name", text="이름")
    student_management_tree.column("name", width=200, stretch=False)

    student_management_tree.heading("email", text="이메일")
    student_management_tree.column("email", width=400)

    student_management_tree.bind("<Button-1>", student_management_on_tree_click)
    student_management_tree.bind("<Double-1>", student_management_edit_cell)

    student_management_load_students()
    window.mainloop()

# =========================================== 로그인 윈도우 ===========================================
app_password = ""
teacher_email = ""

def login():
    login_window = tk.Tk()
    login_window.title("Gmail 로그인")
    login_window.geometry("400x300")

    tk.Label(login_window, text="선생님 Gmail 주소:", font=("Arial", 11)).grid(row=0, column=0, padx=10, pady=(20, 5), sticky="w")
    teacher_email_entry = tk.Entry(login_window, width=40)
    teacher_email_entry.grid(row=1, column=0, padx=10, pady=5, columnspan=2)

    tk.Label(login_window, text="앱 비밀번호:", font=("Arial", 11)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    password_entry = tk.Entry(login_window, width=40, show="*")
    password_entry.grid(row=3, column=0, padx=10, pady=5, columnspan=2)

    def do_login():
        global app_password, teacher_email

        email = teacher_email_entry.get().strip()
        password = password_entry.get().strip()

        if not email or not password:
            messagebox.showwarning("입력 오류", "선생님 이메일, 앱 비밀번호를 모두 입력하세요.")
            return

        # 실제 로그인 시도
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(email, password)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("로그인 실패", "학생 이메일 또는 앱 비밀번호가 틀렸습니다.")
            return
        except Exception as e:
            messagebox.showerror("오류", f"알 수 없는 오류: {e}")
            return

        # 성공
        app_password = password
        teacher_email = email
        login_window.destroy()
        student_management()

    tk.Button(login_window, text="로그인", command=do_login, width=12).grid(row=6, column=0, pady=20)

    login_window.grid_columnconfigure(0, weight=1)
    login_window.mainloop()

if __name__ == '__main__':
    student_management()
