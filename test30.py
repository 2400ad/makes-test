import sqlite3
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

class SQLiteQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SQLite 쿼리 도구")
        self.root.geometry("900x700")
        
        # 데이터베이스 파일 경로
        self.db_file = ""
        self.conn = None
        self.cursor = None
        
        # 하드코딩된 쿼리 상수
        self.QUERY_CONSTANTS = {
            'AAA.TAB1': {'system': 'AAA', 'table': 'TAB1'},
            'BBB.TAB2': {'system': 'BBB', 'table': 'TAB2'}
        }
        
        # UI 구성
        self.create_widgets()
        
    def create_widgets(self):
        # 상단 프레임 (데이터베이스 선택 및 테이블 목록)
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Label(top_frame, text="데이터베이스 파일:").pack(side=tk.LEFT, padx=5)
        
        self.db_path_var = tk.StringVar()
        self.db_path_entry = ttk.Entry(top_frame, textvariable=self.db_path_var, width=50)
        self.db_path_entry.pack(side=tk.LEFT, padx=5)
        
        self.browse_btn = ttk.Button(top_frame, text="찾아보기", command=self.browse_db)
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.connect_btn = ttk.Button(top_frame, text="연결", command=self.connect_db)
        self.connect_btn.pack(side=tk.LEFT, padx=5)
        
        # 중간 프레임 (테이블 목록 및 쿼리 입력)
        mid_frame = ttk.Frame(self.root, padding="10")
        mid_frame.pack(fill=tk.BOTH, expand=True)
        
        # 왼쪽 패널 (테이블 목록)
        left_panel = ttk.LabelFrame(mid_frame, text="테이블 목록", padding="10")
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        
        self.tables_listbox = tk.Listbox(left_panel, width=25, height=15)
        self.tables_listbox.pack(fill=tk.Y, expand=True)
        self.tables_listbox.bind("<Double-1>", self.show_table_structure)
        
        # 오른쪽 패널 (쿼리 입력 및 결과)
        right_panel = ttk.Frame(mid_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 쿼리 입력 영역
        query_frame = ttk.LabelFrame(right_panel, text="SQL 쿼리", padding="10")
        query_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.query_text = ScrolledText(query_frame, height=5)
        self.query_text.pack(fill=tk.X)
        
        # 특수 쿼리 버튼 영역
        special_query_frame = ttk.LabelFrame(right_panel, text="특수 쿼리", padding="10")
        special_query_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.special_query_btn = ttk.Button(special_query_frame, text="AAA.TAB1, BBB.TAB2 쿼리 실행", 
                                           command=self.execute_special_query)
        self.special_query_btn.pack(padx=5, pady=5)
        
        # 쿼리 실행 버튼
        btn_frame = ttk.Frame(right_panel)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.execute_btn = ttk.Button(btn_frame, text="쿼리 실행", command=self.execute_query)
        self.execute_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(btn_frame, text="쿼리 지우기", command=lambda: self.query_text.delete(1.0, tk.END))
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 자주 사용하는 쿼리 버튼들
        self.show_tables_btn = ttk.Button(btn_frame, text="테이블 목록", command=self.show_tables_query)
        self.show_tables_btn.pack(side=tk.LEFT, padx=5)
        
        self.show_schema_btn = ttk.Button(btn_frame, text="스키마 보기", command=self.show_schema_query)
        self.show_schema_btn.pack(side=tk.LEFT, padx=5)
        
        # 결과 영역
        result_frame = ttk.LabelFrame(right_panel, text="쿼리 결과", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # 결과 테이블
        self.result_tree = ttk.Treeview(result_frame)
        self.result_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # 스크롤바 추가
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_x = ttk.Scrollbar(right_panel, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        scrollbar_x.pack(fill=tk.X)
        self.result_tree.configure(xscrollcommand=scrollbar_x.set)
        
        # 상태 표시줄
        self.status_var = tk.StringVar()
        self.status_var.set("준비")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def browse_db(self):
        """데이터베이스 파일 선택 대화상자 표시"""
        file_path = filedialog.askopenfilename(
            title="SQLite 데이터베이스 파일 선택",
            filetypes=[("SQLite 파일", "*.db *.sqlite *.sqlite3"), ("모든 파일", "*.*")]
        )
        if file_path:
            self.db_path_var.set(file_path)
    
    def connect_db(self):
        """선택한 데이터베이스 파일에 연결"""
        db_path = self.db_path_var.get().strip()
        
        if not db_path:
            messagebox.showwarning("경고", "데이터베이스 파일을 선택하세요.")
            return
        
        if not os.path.exists(db_path):
            messagebox.showerror("오류", f"파일이 존재하지 않습니다: {db_path}")
            return
        
        try:
            # 기존 연결 닫기
            if self.conn:
                self.conn.close()
            
            # 새 연결 생성
            self.conn = sqlite3.connect(db_path)
            self.cursor = self.conn.cursor()
            self.db_file = db_path
            
            # 테이블 목록 가져오기
            self.refresh_tables_list()
            
            self.status_var.set(f"데이터베이스 연결됨: {os.path.basename(db_path)}")
            messagebox.showinfo("연결 성공", f"데이터베이스에 성공적으로 연결되었습니다: {os.path.basename(db_path)}")
        
        except sqlite3.Error as e:
            messagebox.showerror("연결 오류", f"데이터베이스 연결 중 오류가 발생했습니다: {e}")
            self.status_var.set("연결 실패")
    
    def refresh_tables_list(self):
        """데이터베이스의 테이블 목록 갱신"""
        if not self.conn:
            return
        
        # 기존 목록 지우기
        self.tables_listbox.delete(0, tk.END)
        
        # 테이블 목록 가져오기
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = self.cursor.fetchall()
        
        for table in tables:
            self.tables_listbox.insert(tk.END, table[0])
    
    def show_table_structure(self, event=None):
        """선택한 테이블의 구조 표시"""
        if not self.conn:
            messagebox.showwarning("경고", "먼저 데이터베이스에 연결하세요.")
            return
        
        # 선택한 테이블 가져오기
        selection = self.tables_listbox.curselection()
        if not selection:
            return
        
        table_name = self.tables_listbox.get(selection[0])
        
        # 테이블 구조 쿼리 생성 및 실행
        query = f"PRAGMA table_info({table_name})"
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(tk.END, query)
        self.execute_query()
    
    def show_tables_query(self):
        """테이블 목록 조회 쿼리 실행"""
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(tk.END, "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        self.execute_query()
    
    def show_schema_query(self):
        """선택한 테이블의 스키마 조회 쿼리 실행"""
        selection = self.tables_listbox.curselection()
        if not selection:
            messagebox.showinfo("정보", "테이블을 먼저 선택하세요.")
            return
        
        table_name = self.tables_listbox.get(selection[0])
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(tk.END, f"PRAGMA table_info({table_name})")
        self.execute_query()
    
    def execute_special_query(self):
        """하드코딩된 특수 쿼리 실행 (AAA.TAB1, BBB.TAB2)"""
        if not self.conn:
            messagebox.showwarning("경고", "먼저 데이터베이스에 연결하세요.")
            return
        
        # 하드코딩된 값 사용
        system1 = 'AAA'
        table1 = 'TAB1'
        system2 = 'BBB'
        table2 = 'TAB2'
        
        # 하드코딩된 컬럼 이름
        col1 = 'COL1X'
        col2 = 'COL2X'
        col3 = 'COL3X'
        col4 = 'COL4X'
        
        # 쿼리 생성
        query = f"""
        SELECT * FROM book2 
        WHERE {col1} = '{system1}' 
        AND {col2} = '{table1}' 
        AND {col3} = '{system2}' 
        AND {col4} = '{table2}'
        """
        
        # 쿼리 표시 및 실행
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(tk.END, query)
        
        try:
            # 쿼리 실행
            self.cursor.execute(query)
            
            # 결과 가져오기
            results = self.cursor.fetchall()
            column_names = [description[0] for description in self.cursor.description]
            
            # 결과 표시
            self.display_formatted_results(results, column_names, system1, table1, system2, table2)
            
            self.status_var.set(f"특수 쿼리 실행 완료: {len(results)}개의 결과")
            
        except sqlite3.Error as e:
            messagebox.showerror("쿼리 오류", f"쿼리 실행 중 오류가 발생했습니다: {e}")
            self.status_var.set("쿼리 실행 실패")
    
    def display_formatted_results(self, results, column_names, system1, table1, system2, table2):
        """특수 쿼리 결과를 보기 좋게 표시"""
        # 기존 테이블 내용 지우기
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # 컬럼 설정
        self.result_tree["columns"] = column_names
        self.result_tree["show"] = "headings"  # 기본 첫 번째 컬럼 숨기기
        
        # 컬럼 헤더 설정 - 보기 좋게 스타일 적용
        for i, col in enumerate(column_names):
            # 중요 컬럼 강조
            if col in ['COL1X', 'COL2X', 'COL3X', 'COL4X']:
                self.result_tree.heading(col, text=f"★ {col} ★")
                self.result_tree.column(col, width=120, anchor='center')
            else:
                self.result_tree.heading(col, text=col)
                self.result_tree.column(col, width=100)
        
        # 결과가 없는 경우
        if not results:
            messagebox.showinfo("정보", f"'{system1}.{table1}, {system2}.{table2}'에 대한 검색 결과가 없습니다.")
            return
        
        # 결과 행 추가 - 보기 좋게 스타일 적용
        for i, row in enumerate(results):
            row_values = list(row)
            
            # 특정 컬럼 값 강조 (예: 시스템 및 테이블 이름)
            for j, val in enumerate(row_values):
                col_name = column_names[j]
                if col_name == 'COL1X' and val == system1:
                    row_values[j] = f"✓ {val}"
                elif col_name == 'COL2X' and val == table1:
                    row_values[j] = f"✓ {val}"
                elif col_name == 'COL3X' and val == system2:
                    row_values[j] = f"✓ {val}"
                elif col_name == 'COL4X' and val == table2:
                    row_values[j] = f"✓ {val}"
            
            # 행 추가
            tag = 'even' if i % 2 == 0 else 'odd'
            self.result_tree.insert("", tk.END, values=row_values, tags=(tag,))
        
        # 행 색상 설정 - 더 보기 좋게
        self.result_tree.tag_configure('even', background='#e6f2ff')  # 연한 파란색
        self.result_tree.tag_configure('odd', background='#ffffff')   # 흰색
        
        # 결과 요약 표시
        messagebox.showinfo("검색 결과", 
                           f"'{system1}.{table1}, {system2}.{table2}'에 대한 검색 결과: {len(results)}개 항목 발견")
    
    def execute_query(self):
        """SQL 쿼리 실행"""
        if not self.conn:
            messagebox.showwarning("경고", "먼저 데이터베이스에 연결하세요.")
            return
        
        query = self.query_text.get(1.0, tk.END).strip()
        if not query:
            messagebox.showinfo("정보", "실행할 쿼리를 입력하세요.")
            return
        
        try:
            # 쿼리 실행
            self.cursor.execute(query)
            
            # 결과가 있는 쿼리인지 확인
            if query.lower().startswith(("select", "pragma", "show")):
                # 결과 가져오기
                results = self.cursor.fetchall()
                column_names = [description[0] for description in self.cursor.description]
                
                # 결과 표시
                self.display_results(results, column_names)
                self.status_var.set(f"쿼리 실행 완료: {len(results)}개의 결과")
            else:
                # 변경 사항 커밋
                self.conn.commit()
                self.status_var.set("쿼리 실행 완료")
                messagebox.showinfo("성공", "쿼리가 성공적으로 실행되었습니다.")
                
                # 테이블 목록 갱신 (테이블이 추가/삭제되었을 수 있음)
                self.refresh_tables_list()
        
        except sqlite3.Error as e:
            messagebox.showerror("쿼리 오류", f"쿼리 실행 중 오류가 발생했습니다: {e}")
            self.status_var.set("쿼리 실행 실패")
    
    def display_results(self, results, column_names):
        """일반 쿼리 결과를 테이블에 표시"""
        # 기존 테이블 내용 지우기
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # 컬럼 설정
        self.result_tree["columns"] = column_names
        self.result_tree["show"] = "headings"  # 기본 첫 번째 컬럼 숨기기
        
        # 컬럼 헤더 설정
        for col in column_names:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=100)  # 기본 너비 설정
        
        # 결과 행 추가
        for i, row in enumerate(results):
            self.result_tree.insert("", tk.END, values=row, tags=('even' if i % 2 == 0 else 'odd',))
        
        # 행 색상 설정
        self.result_tree.tag_configure('even', background='#f0f0f0')
        self.result_tree.tag_configure('odd', background='#ffffff')

def main():
    root = tk.Tk()
    app = SQLiteQueryApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
