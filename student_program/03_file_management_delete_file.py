import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sqlite3
import os
import openpyxl

DB_NAME = 'student_db.db'
selected_file_path = None
file_checked_items = {}
file_id = {}

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS word_test_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_path TEXT NOT NULL UNIQUE
        )
    ''')
    conn.commit()
    conn.close()

def load_word_test_files():
    file_checked_items.clear()
    file_id.clear()

    for row in file_tree.get_children():
        file_tree.delete(row)

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('SELECT id, file_path FROM word_test_files ORDER BY id')
    rows = cursor.fetchall()
    conn.close()

    for fid, file_path in rows:
        item_id = file_tree.insert('', 'end', values=('☐', file_path))
        file_checked_items[item_id] = False
        file_id[item_id] = fid

def toggle_checkbox(event):
    item_id = file_tree.identify_row(event.y)
    column = file_tree.identify_column(event.x)

    if column == '#1' and item_id:
        checked = file_checked_items.get(item_id, False)
        file_checked_items[item_id] = not checked
        new_icon = '☑' if not checked else '☐'
        file_tree.set(item_id, column='선택', value=new_icon)

def register_file():
    file_path = selected_file_path.get()
    if not file_path:
        messagebox.showwarning("경고", "먼저 파일을 선택하세요.")
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM word_test_files WHERE file_path = ?", (file_path,))
    count = cursor.fetchone()[0]

    if count > 0:
        messagebox.showwarning("중복", "이미 등록된 파일입니다.")
    else:
        cursor.execute("INSERT INTO word_test_files (file_path) VALUES (?)", (file_path,))
        conn.commit()
        messagebox.showinfo("성공", "파일이 등록되었습니다.")
        selected_file_path.set("")
        uploaded_file_label.config(text="선택된 파일: 없음")
        load_word_test_files()

    conn.close()

def file_upload():
    file_path = filedialog.askopenfilename()
    if file_path:
        selected_file_path.set(file_path)
        uploaded_file_label.config(text=f"선택된 파일: {file_path}")

def cancel_upload():
    selected_file_path.set("")
    uploaded_file_label.config(text="선택된 파일: 없음")

# 선택된 파일 삭제
def delete_selected_files():
    to_delete_ids = [file_id[item] for item, checked in file_checked_items.items() if checked]
    if not to_delete_ids:
        messagebox.showwarning("경고", "삭제할 파일을 선택하세요.")
        return

    confirm = messagebox.askyesno("삭제 확인", f"{len(to_delete_ids)}개의 파일을 삭제하시겠습니까?")
    if not confirm:
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.executemany("DELETE FROM word_test_files WHERE id = ?", [(fid,) for fid in to_delete_ids])
    conn.commit()
    conn.close()

    load_word_test_files()

# 파일 관리 셀 선택시 word_test() 실행
def handle_file_path_click(event):
    region = file_tree.identify("region", event.x, event.y)
    if region != "cell":
        return  # 헤더나 빈 공간 클릭 시 무시

    col = file_tree.identify_column(event.x)
    row = file_tree.identify_row(event.y)

    if col == "#2" and row:  # "#2"는 "파일 경로" 열
        file_path = file_tree.item(row, "values")[1]
        word_test(file_path)

# =========================================== 파일 윈도우 ===========================================
def file_management():
    global selected_file_path, uploaded_file_label, file_tree

    init_db()

    file_window = tk.Tk()
    file_window.title("파일 업로드 및 등록")
    file_window.geometry("720x480")

    selected_file_path = tk.StringVar()

    # 네비게이션 프레임
    nav_frame = tk.Frame(file_window)
    nav_frame.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="ew")

    tk.Button(nav_frame, text="홈으로", width=14).grid(row=0, column=0, padx=5)
    tk.Button(nav_frame, text="종료", width=14, command=file_window.destroy).grid(row=0, column=1, padx=5)

    # 상단 프레임
    top_frame = tk.LabelFrame(file_window, text="파일 선택")
    top_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

    uploaded_file_label = tk.Label(top_frame, text="선택된 파일: 없음", anchor="w")
    uploaded_file_label.grid(row=0, column=0, columnspan=2, padx=20, pady=(5, 0), sticky="w")

    tk.Button(top_frame, text="파일 업로드", command=file_upload, width=30, height=2)\
        .grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="w")

    button_frame = tk.Frame(top_frame)
    button_frame.grid(row=2, column=0, columnspan=2, pady=10)

    tk.Button(button_frame, text="파일 등록", command=register_file, width=14, height=2)\
        .grid(row=0, column=0, padx=20)
    tk.Button(button_frame, text="취소", command=cancel_upload, width=14, height=2)\
        .grid(row=0, column=1, padx=20)

    # 중간 프레임
    middle_frame = tk.LabelFrame(file_window, text="등록된 파일 목록")
    middle_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

    # 삭제 버튼
    tk.Button(middle_frame, text="선택된 파일 삭제", command=delete_selected_files, width=20, height=2)\
        .grid(row=0, column=0, padx=10, pady=(10, 10), sticky="w")

    # 트리뷰
    columns = ("선택", "파일 경로")
    file_tree = ttk.Treeview(middle_frame, columns=columns, show="headings", height=10)
    file_tree.heading("선택", text="선택")
    file_tree.heading("파일 경로", text="파일 경로")
    file_tree.column("선택", width=60, anchor="center", stretch=False)
    file_tree.column("파일 경로", width=600, anchor="w", stretch=True)
    file_tree.grid(row=1, column=0, sticky="nsew")

    scrollbar = ttk.Scrollbar(middle_frame, orient="vertical", command=file_tree.yview)
    file_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=1, sticky="ns")

    # 확장 설정
    file_window.grid_rowconfigure(2, weight=1)
    file_window.grid_columnconfigure(0, weight=1)
    middle_frame.grid_rowconfigure(1, weight=1)
    middle_frame.grid_columnconfigure(0, weight=1)

    file_tree.bind("<Button-1>", toggle_checkbox)
    file_tree.bind("<ButtonRelease-1>", handle_file_path_click)

    load_word_test_files()
    file_window.mainloop()

# =========================================== 단어 테스트 윈도우 ===========================================
def word_test(file_path):
    # 새 창 생성
    test_window = tk.Toplevel()
    test_window.title("test_window")
    test_window.geometry("400x300")

    # 엑셀 파일 열기
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        words = [str(cell.value) for cell in sheet['A'] if cell.value is not None]
        wb.close()
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 파일을 열 수 없습니다:\n{e}")
        test_window.destroy()
        return

    # 단어 리스트박스
    tk.Label(test_window, text="단어 목록").pack(pady=5)
    listbox = tk.Listbox(test_window, width=50, height=10)
    listbox.pack(padx=10, pady=5, fill="both", expand=True)

    for w in words:
        listbox.insert(tk.END, w)

    # 단어 선택 시 상세 표시(간단히 단어 텍스트만 표시)
    word_label = tk.Label(test_window, text="", font=("Arial", 14, "bold"))
    word_label.pack(pady=10)

    def on_select(event):
        selection = listbox.curselection()
        if selection:
            index = selection[0]
            word_label.config(text=words[index])

    listbox.bind("<<ListboxSelect>>", on_select)

# =========================================== 실행 ===========================================
if __name__ == '__main__':
    file_management()
