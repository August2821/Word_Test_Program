import os
import shutil
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import smtplib
from email.message import EmailMessage

# 전역 변수 선언
sender_email = ""  # 선생님 이메일
app_password = "" # 선생님 앱 비밀번호

DB_NAME = 'teacher_db.db' # DB 이름

file_checked_items = {} # 파일 체크 상태 저장용
file_all_checked = False  # 파일 전체 선택 상태
file_id_map = {} # 파일 id 저장용

student_checked_items = {} # 학생 체크 상태 저장용
student_all_checked = False  # 학생 전체 선택 상태
student_id_map = {} # 파일 id 저장용

# DB 초기화
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS word_test_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_path TEXT NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
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

# =========================================== 파일 관리 관련 함수 ===========================================
# DB에서 데이터 읽어 트리뷰에 출력
def file_management_load_word_test_files(tree):
    file_checked_items.clear()
    file_id_map.clear()

    for row in tree.get_children():
        tree.delete(row)

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('SELECT id, file_path FROM word_test_files ORDER BY id')
    rows = cursor.fetchall()
    conn.close()

    for row in rows:
        file_id, file_path = row
        # 경로 전체를 보여줌
        item_id = tree.insert('', 'end', values=('☐', file_path))
        file_checked_items[item_id] = False
        file_id_map[item_id] = file_id

# 파일 업로드 핸들러 (파일 복사만)
def file_management_handle_file_upload():
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return
    try:
        filename = os.path.basename(file_path)

        file_path_var.set(file_path)  # 원래 경로를 그대로 저장
        uploaded_file_label.config(text=f"선택된 파일: {filename}")
    except Exception as e:
        messagebox.showerror("오류", f"파일 선택 중 오류 발생: {e}")

# 등록 버튼 핸들러 (DB 저장)
def file_management_register_file():
    path = file_path_var.get()

    if not path:
        messagebox.showwarning("경고", "먼저 파일을 업로드하세요.")
        return

    filename = os.path.basename(path)

    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM word_test_files WHERE file_path LIKE ?', ('%' + filename,))
        (count,) = cursor.fetchone()
        if count > 0:
            messagebox.showwarning("중복 오류", f"파일 이름 '{filename}'이(가) 이미 등록되어 있습니다.")
            conn.close()
            return

        cursor.execute('INSERT INTO word_test_files (file_path) VALUES (?)', (path,))
        conn.commit()
        conn.close()

        file_management_reset_ui()
        file_management_load_word_test_files(file_tree)
    except Exception as e:
        messagebox.showerror("오류", f"DB 저장 중 오류 발생: {e}")

# 취소 버튼 핸들러 (업로드 파일 삭제 + UI 초기화)
def file_management_cancel_upload():
    path = file_path_var.get()
    if path and os.path.exists(path):
        filename = os.path.basename(path)
        # DB에 같은 파일 이름이 이미 존재하는지 확인
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM word_test_files WHERE file_path LIKE ?", ('%' + filename,))
        (count,) = cursor.fetchone()
        conn.close()

        if count == 0:
            try:
                os.remove(path)
            except Exception as e:
                messagebox.showerror("오류", f"파일 삭제 중 오류 발생: {e}")
        else:
            # DB에 이미 존재하므로 삭제하지 않음
            print(f"파일 '{filename}'은 DB에 이미 존재하므로 삭제하지 않음.")

    file_management_reset_ui()

# UI 리셋 함수
def file_management_reset_ui():
    file_path_var.set("")
    uploaded_file_label.config(text="선택된 파일: 없음")

# 체크박스 토글
def file_management_toggle_checkbox(event):
    global file_all_checked
    region = file_tree.identify('region', event.x, event.y)
    column = file_tree.identify_column(event.x)
    row = file_tree.identify_row(event.y)

    if column == '#1':  # 선택 열
        if region == 'cell' and row:
            file_checked_items[row] = not file_checked_items.get(row, False)
            file_tree.set(row, column, '☑' if file_checked_items[row] else '☐')
        elif region == 'heading':
            file_all_checked = not file_all_checked
            for item in file_tree.get_children():
                file_checked_items[item] = file_all_checked
                file_tree.set(item, column, '☑' if file_all_checked else '☐')

# 선택된 파일 삭제
def file_management_delete_selected_files():
    selected_ids = []
    for item_id, checked in file_checked_items.items():
        if checked:
            file_id = file_id_map.get(item_id)
            selected_ids.append(file_id)

    if not selected_ids:
        messagebox.showinfo("알림", "삭제할 항목을 선택하세요.")
        return

    confirm = messagebox.askyesno("확인", f"{len(selected_ids)}개 항목을 삭제하시겠습니까?")
    if not confirm:
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    for file_id in selected_ids:
        cursor.execute("SELECT file_path FROM word_test_files WHERE id=?", (file_id,))
        row = cursor.fetchone()
        if row and os.path.exists(row[0]):
            try:
                os.remove(row[0])
            except:
                pass
        cursor.execute("DELETE FROM word_test_files WHERE id=?", (file_id,))
    conn.commit()
    conn.close()

    file_management_load_word_test_files(file_tree)

# =========================================== 학생 관련 함수 ===========================================

def file_management_load_students(filter_class=None):
    for row in student_tree.get_children():
        student_tree.delete(row)
    student_checked_items.clear()
    student_id_map.clear()

    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    if filter_class is None:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students")
    else:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students WHERE class_number = ?", (filter_class,))
    rows = cursor.fetchall()
    conn.close()

    for row in rows:
        student_id, class_num, seat, name, email = row
        item_id = student_tree.insert("", tk.END, values=("☐", class_num, seat, name, email))
        student_checked_items[item_id] = False
        student_id_map[item_id] = student_id  # 트리뷰 row → student_id 매핑 저장

def file_management_on_tree_click(event):
    global student_all_checked
    region = student_tree.identify_region(event.x, event.y)
    column = student_tree.identify_column(event.x)
    row = student_tree.identify_row(event.y)

    if column == "#1":
        if region == "cell":
            student_checked_items[row] = not student_checked_items.get(row, False)
            student_tree.set(row, column="#1", value="☑" if student_checked_items[row] else "☐")
        elif region == "heading":
            student_all_checked = not student_all_checked
            for item_id in student_tree.get_children():
                student_checked_items[item_id] = student_all_checked
                student_tree.set(item_id, column="#1", value="☑" if student_all_checked else "☐")

    elif region == "heading" and column == "#2":  # 반 번호
        file_management_show_class_menu(event)

    elif region == "cell":
        col_index = int(column.replace("#", ""))
        if col_index >= 2:  # class, seat, name, email
            file_management_edit_cell(row, column)

def file_management_edit_cell(row_id, column):
    x, y, width, height = student_tree.bbox(row_id, column)
    value = student_tree.set(row_id, column)
    col_name = student_tree.heading(column)["text"]

    entry = tk.Entry(student_tree)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, value)
    entry.focus()

    def save_edit(event):
        new_value = entry.get().strip()
        if not new_value:
            messagebox.showwarning("입력 오류", f"{col_name}은(는) 빈 값일 수 없습니다.")
            entry.focus()
            return

        if column in ("#2", "#3"):
            if not new_value.isdigit():
                messagebox.showerror("형식 오류", f"{col_name}은(는) 숫자여야 합니다.")
                entry.focus()
                return

        entry.destroy()
        student_tree.set(row_id, column, new_value)

        student_id = student_id_map.get(row_id)
        col_map = {
            "#2": "class_number",
            "#3": "seat_number",
            "#4": "name",
            "#5": "email"
        }
        db_col = col_map.get(column)
        if db_col and student_id:
            conn = sqlite3.connect("teacher_db.db")
            cursor = conn.cursor()
            cursor.execute(f"UPDATE students SET {db_col} = ? WHERE id = ?", (new_value, student_id))
            conn.commit()
            conn.close()
            file_management_load_students()

    entry.bind("<Return>", save_edit)
    entry.bind("<FocusOut>", lambda e: entry.destroy())

def file_management_get_class_numbers():
    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT class_number FROM students ORDER BY class_number")
    classes = [row[0] for row in cursor.fetchall()]
    conn.close()
    return classes

def file_management_show_class_menu(event):
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="전체 보기", command=lambda: file_management_load_students())
    for class_num in file_management_get_class_numbers():
        menu.add_command(label=f"{class_num}반 보기", command=lambda c=class_num: file_management_load_students(c))
    menu.post(event.x_root, event.y_root)

# =========================================== 이메일 보내기 ===========================================
def file_management_send_files_to_students():
    selected_files = []
    for item_id, checked in file_checked_items.items():
        if checked:
            file_id = file_tree.item(item_id, 'values')[1]
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("SELECT file_path FROM word_test_files WHERE id=?", (file_id,))
            row = cursor.fetchone()
            conn.close()
            if row:
                selected_files.append(row[0])

    selected_students = []
    for item_id, checked in student_checked_items.items():
        if checked:
            email = student_tree.item(item_id, 'values')[4]
            if email:
                selected_students.append(email)

    if not selected_files:
        messagebox.showwarning("파일 없음", "보낼 파일을 선택하세요.")
        return
    if not selected_students:
        messagebox.showwarning("학생 없음", "보낼 학생을 선택하세요.")
        return

    for email in selected_students:
        try:
            msg = EmailMessage()
            msg['Subject'] = '단어 테스트 파일'
            msg['From'] = sender_email
            msg['To'] = email
            msg.set_content("첨부된 파일을 확인하세요.")

            for path in selected_files:
                with open(path, 'rb') as f:
                    file_data = f.read()
                    filename = os.path.basename(path)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, app_password)
                smtp.send_message(msg)
        except Exception as e:
            messagebox.showerror("전송 오류", f"{email}로 전송 실패: {e}")
            continue

    messagebox.showinfo("완료", "모든 파일을 성공적으로 전송했습니다.")

# =========================================== UI 구성 ===========================================
def file_management_window():
    init_db()

    global root, file_path_var, uploaded_file_label, file_tree, student_tree

    root = tk.Tk()
    root.title("학생 관리")
    #root.geometry("600x400")

    file_path_var = tk.StringVar()

    # ======================= 상단 프레임: 파일 선택 =======================
    top_frame = tk.LabelFrame(root, text="파일 선택")
    top_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    uploaded_file_label = tk.Label(top_frame, text="선택된 파일: 없음")
    uploaded_file_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=20, pady=(10, 0))

    tk.Button(top_frame, text="파일 업로드", command=file_management_handle_file_upload, width=30, height=2)\
        .grid(row=3, column=0, columnspan=2, padx=20, pady=10, sticky="w")

    # 등록 / 취소 버튼 (이제 top_frame 안에 위치)
    button_frame = tk.Frame(top_frame)
    button_frame.grid(row=4, column=0, columnspan=2, pady=10,sticky="we")

    tk.Button(button_frame, text="파일 등록", command=file_management_register_file, width=14, height=2)\
        .grid(row=0, column=0, padx=20)
    tk.Button(button_frame, text="취소", command=file_management_cancel_upload, width=14, height=2)\
        .grid(row=0, column=1, padx=20)

    # ======================= 중간 프레임: 파일 목록 =======================
    middle_frame = tk.LabelFrame(root, text="파일 목록")
    middle_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    # 버튼 프레임 (삭제/전송)
    file_button_frame = tk.Frame(middle_frame)
    file_button_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=(5, 10))

    tk.Button(file_button_frame, text="선택된 파일 항목 삭제", command=file_management_delete_selected_files, height=2)\
        .grid(row=0, column=0, padx=(20, 5), sticky="w")

    tk.Button(file_button_frame, text="파일 보내기", command=file_management_send_files_to_students, height=2)\
        .grid(row=0, column=1, padx=20, sticky="e")

    # 스크롤바 프레임
    file_tree_frame = tk.Frame(middle_frame)
    file_tree_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")

    # 스크롤바
    file_scrollbar = ttk.Scrollbar(file_tree_frame, orient="vertical")
    file_scrollbar.grid(row=0, column=1, sticky="ns")

    # 파일 목록 Treeview
    columns = ('선택', '파일 경로')
    file_tree = ttk.Treeview(
        file_tree_frame,
        columns=columns,
        show='headings',
        yscrollcommand=file_scrollbar.set
    )

    for col in columns:
        file_tree.heading(col, text=col)
        if col == '선택':
            file_tree.column(col, width=40, anchor='center', stretch=False)
        else:
            file_tree.column(col, width=600, anchor='w')

    file_tree.grid(row=0, column=0, sticky="nsew")
    file_scrollbar.config(command=file_tree.yview)
    file_tree.bind('<Button-1>', file_management_toggle_checkbox)

    # Treeview 확장 설정
    file_tree_frame.grid_rowconfigure(0, weight=1)
    file_tree_frame.grid_columnconfigure(0, weight=1)

    middle_frame.grid_rowconfigure(1, weight=1)
    middle_frame.grid_columnconfigure(0, weight=1)

    file_management_load_word_test_files(file_tree)

    # ======================= 하단 프레임: 학생 리스트 =======================
    bottom_frame = tk.LabelFrame(root, text="학생 목록")
    bottom_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

    student_tree_frame = tk.Frame(bottom_frame)
    student_tree_frame.grid(row=0, column=0, sticky="nsew")

    scrollbar = ttk.Scrollbar(student_tree_frame, orient="vertical")
    scrollbar.grid(row=0, column=1, sticky="ns")

    student_tree = ttk.Treeview(
        student_tree_frame,
        columns=("check", "class", "seat", "name", "email"),
        show="headings",
        yscrollcommand=scrollbar.set
    )
    student_tree.grid(row=0, column=0, sticky="nsew")
    student_tree_frame.grid_rowconfigure(0, weight=1)
    student_tree_frame.grid_columnconfigure(0, weight=1)
    scrollbar.config(command=student_tree.yview)

    student_tree.heading("check", text="선택")
    student_tree.column("check", width=40, anchor="center", stretch=False)
    student_tree.heading("class", text="반 번호")
    student_tree.column("class", width=60, anchor="center", stretch=False)
    student_tree.heading("seat", text="좌석 번호")
    student_tree.column("seat", width=60, anchor="center", stretch=False)
    student_tree.heading("name", text="이름")
    student_tree.column("name", width=200, stretch=False)
    student_tree.heading("email", text="이메일")
    student_tree.column("email", width=400)

    student_tree.bind("<Button-1>", file_management_on_tree_click)

    bottom_frame.grid_rowconfigure(0, weight=1)
    bottom_frame.grid_columnconfigure(0, weight=1)

    file_management_load_students()

    root.grid_rowconfigure(1, weight=1)  # 학생 리스트 확장
    root.grid_columnconfigure(0, weight=1)
# ======================================================================================
    root.mainloop()

# =========================================== 로그인 ===========================================
def login_window():
    login_root = tk.Tk()
    login_root.title("Gmail 로그인")
    login_root.geometry("400x250")

    tk.Label(login_root, text="Gmail 주소:", font=("Arial", 11)).grid(row=0, column=0, padx=10, pady=(20, 5), sticky="w")
    email_entry = tk.Entry(login_root, width=40)
    email_entry.grid(row=1, column=0, padx=10, pady=5, columnspan=2)

    tk.Label(login_root, text="앱 비밀번호:", font=("Arial", 11)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    pw_entry = tk.Entry(login_root, width=40, show="*")
    pw_entry.grid(row=3, column=0, padx=10, pady=5, columnspan=2)

    def do_login():
        nonlocal email_entry, pw_entry
        global sender_email, app_password

        email = email_entry.get().strip()
        pw = pw_entry.get().strip()

        if not email or not pw:
            messagebox.showwarning("입력 오류", "이메일과 앱 비밀번호를 모두 입력하세요.")
            return

        # 실제 로그인 시도
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(email, pw)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("로그인 실패", "이메일 또는 앱 비밀번호가 틀렸습니다.")
            return
        except Exception as e:
            messagebox.showerror("오류", f"알 수 없는 오류: {e}")
            return

        # 성공
        sender_email = email
        app_password = pw
        login_root.destroy()
        file_management_window()

    tk.Button(login_root, text="로그인", command=do_login, width=12).grid(row=4, column=0, pady=20)

    login_root.grid_columnconfigure(0, weight=1)
    login_root.mainloop()


if __name__ == '__main__':
    file_management_window()
