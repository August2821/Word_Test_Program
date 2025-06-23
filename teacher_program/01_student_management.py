import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import sqlite3

checked_items = {}
all_checked = False  # 전체 선택 상태 추적

# DB 초기화
def init_db():
    conn = sqlite3.connect("teacher_db.db")
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

# 학생 등록 함수
def register_student():
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

    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO students (class_number, name, seat_number, email) VALUES (?, ?, ?, ?)",
                   (class_num, name, seat_num, email))
    conn.commit()
    conn.close()

    entry_class.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_seat.delete(0, tk.END)
    entry_email.delete(0, tk.END)

    load_students()

# 학생 목록 조회
def load_students(filter_class=None):
    for row in tree.get_children():
        tree.delete(row)
    checked_items.clear()

    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    if filter_class is None:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students")
    else:
        cursor.execute("SELECT id, class_number, seat_number, name, email FROM students WHERE class_number = ?", (filter_class,))
    rows = cursor.fetchall()
    conn.close()

    for row in rows:
        item_id = tree.insert("", tk.END, values=("☐", *row))  # 첫 칼럼은 체크박스 자리
        checked_items[item_id] = False

# 트리뷰 반 번호 헤더, 체크 헤더, 셀 클릭에  바인딩
def on_tree_click(event):
    global all_checked
    region = tree.identify_region(event.x, event.y)
    column = tree.identify_column(event.x)
    row = tree.identify_row(event.y)

    if column == "#1":
        if region == "cell":
            # 개별 체크박스 토글
            checked_items[row] = not checked_items.get(row, False)
            tree.set(row, column="#1", value="☑" if checked_items[row] else "☐")
        elif region == "heading":
            # 헤더 클릭 시 전체 토글
            all_checked = not all_checked
            for item_id in tree.get_children():
                checked_items[item_id] = all_checked
                tree.set(item_id, column="#1", value="☑" if all_checked else "☐")

    elif region == "heading" and column == "#3":  # 반 번호 헤더 메뉴
        show_class_menu(event)
    
    # 클릭한 셀 편집 (체크, ID 칼럼 제외)
    elif region == "cell":
        col_index = int(column.replace("#", ""))  # "#3" -> 3
        if col_index >= 3:  # class, seat, name, email만 수정 허용
            edit_cell(row, column)

# 셀 편집용 Entry 생성
def edit_cell(row_id, column):
    x, y, width, height = tree.bbox(row_id, column)
    value = tree.set(row_id, column)
    col_name = tree.heading(column)["text"]

    entry = tk.Entry(tree)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, value)
    entry.focus()

    def save_edit(event):
        new_value = entry.get().strip()
        if not new_value:
            messagebox.showwarning("입력 오류", f"{col_name}은(는) 빈 값일 수 없습니다.")
            entry.focus()
            return

        # 반 번호, 좌석 번호는 숫자여야 함
        if column in ("#3", "#4"):
            if not new_value.isdigit():
                messagebox.showerror("형식 오류", f"{col_name}은(는) 숫자여야 합니다.")
                entry.focus()
                return

        entry.destroy()
        tree.set(row_id, column, new_value)

        # DB에 반영
        student_id = tree.set(row_id, "#2")  # ID 칼럼
        col_map = {
            "#3": "class_number",
            "#4": "seat_number",
            "#5": "name",
            "#6": "email"
        }
        db_col = col_map.get(column)
        if db_col:
            conn = sqlite3.connect("teacher_db.db")
            cursor = conn.cursor()
            cursor.execute(f"UPDATE students SET {db_col} = ? WHERE id = ?", (new_value, student_id))
            conn.commit()
            conn.close()
            load_students()

    entry.bind("<Return>", save_edit)
    entry.bind("<FocusOut>", lambda e: entry.destroy())


# 반 번호 목록 조회 함수
def get_class_numbers():
    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT class_number FROM students ORDER BY class_number")
    classes = [row[0] for row in cursor.fetchall()]
    conn.close()
    return classes

# 반 번호 헤더 클릭 시 메뉴 표시 함수
def show_class_menu(event):
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="전체 보기", command=lambda: load_students())
    for class_num in get_class_numbers():
        menu.add_command(label=f"{class_num}반 보기", command=lambda c=class_num: load_students(c))
    menu.post(event.x_root, event.y_root)

# 선택된 학생 삭제 함수
def delete_selected():
    conn = sqlite3.connect("teacher_db.db")
    cursor = conn.cursor()
    for item_id, checked in checked_items.items():
        if checked:
            student_id = tree.item(item_id, "values")[1]  # ID는 두 번째 열
            cursor.execute("DELETE FROM students WHERE id = ?", (student_id,))
    conn.commit()
    conn.close()
    load_students()

init_db()

# =========================================== UI 구성 ===========================================
root = tk.Tk()
root.title("선생님 프로그램")

# =========================================== Frame 설정 ===========================================
left_frame = tk.Frame(root)
left_frame.grid(row=0, column=0, padx=10, sticky="n")

right_frame = tk.Frame(root)
right_frame.grid(row=0, column=1, sticky="nsew")

root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)

# =========================================== 왼쪽 입력 폼 ===========================================
tk.Label(left_frame, text="반 번호").grid(row=0, column=0, padx=5, pady=5)
entry_class = tk.Entry(left_frame)
entry_class.grid(row=0, column=1, padx=5, pady=5)

tk.Label(left_frame, text="좌석 번호").grid(row=1, column=0, padx=5, pady=5)
entry_seat = tk.Entry(left_frame)
entry_seat.grid(row=1, column=1, padx=5, pady=5)

tk.Label(left_frame, text="이름").grid(row=2, column=0, padx=5, pady=5)
entry_name = tk.Entry(left_frame)
entry_name.grid(row=2, column=1, padx=5, pady=5)

tk.Label(left_frame, text="이메일").grid(row=3, column=0, padx=5, pady=5)
entry_email = tk.Entry(left_frame)
entry_email.grid(row=3, column=1, padx=5, pady=5)

btn_register = tk.Button(left_frame, text="학생 등록", command=register_student)
btn_register.grid(row=4, column=0, columnspan=2, pady=10)

# =========================================== 오른쪽 트리뷰 + 스크롤 + 삭제 버튼 ===========================================
tree_frame = tk.Frame(right_frame)
tree_frame.grid(row=0, column=0, sticky="nsew")

right_frame.grid_rowconfigure(0, weight=1)
right_frame.grid_columnconfigure(0, weight=1)

scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
scrollbar.grid(row=0, column=1, sticky="ns")

tree = ttk.Treeview(
    tree_frame,
    columns=("check", "id", "class", "seat", "name", "email"),
    show="headings",
    yscrollcommand=scrollbar.set
)
tree.grid(row=0, column=0, sticky="nsew")
tree_frame.grid_rowconfigure(0, weight=1)
tree_frame.grid_columnconfigure(0, weight=1)

scrollbar.config(command=tree.yview)

tree.heading("check", text="선택")
tree.column("check", width=30, anchor='center')

tree.heading("id", text="ID")
tree.column("id", width=40, anchor='center')

tree.heading("class", text="반 번호")
tree.column("class", width=40, anchor='center')

tree.heading("seat", text="좌석 번호")
tree.column("seat", width=60, anchor='center')

tree.heading("name", text="이름")
tree.column("name", width=200)

tree.heading("email", text="이메일")
tree.column("email", width=400)

tree.bind("<Button-1>", on_tree_click)

btn_delete = tk.Button(right_frame, text="선택 삭제", command=delete_selected)
btn_delete.grid(row=1, column=0, pady=10)

# ============================================================================================
load_students()
root.mainloop()