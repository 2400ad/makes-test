import openpyxl
import sqlite3
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

def excel_to_sqlite(excel_configs, db_file='test.db', log_callback=None):
    """
    엑셀 파일의 데이터를 SQLite 데이터베이스로 변환합니다.
    
    Args:
        excel_configs (dict): 엑셀 파일 구성 정보를 담은 딕셔너리
        db_file (str): SQLite 데이터베이스 파일 경로
        log_callback (function): 로그 메시지를 표시하기 위한 콜백 함수
    """
    def log(message):
        if log_callback:
            log_callback(message)
        else:
            print(message)
    
    def clean_column_name(name):
        """컬럼 이름에서 공백과 특수문자를 제거하고 SQLite에 적합한 이름으로 변환합니다."""
        if name is None:
            return ""
        
        # 문자열로 변환
        name = str(name).strip()
        
        # 빈 이름 처리
        if not name:
            return ""
        
        # 공백과 특수문자를 언더스코어로 대체
        import re
        name = re.sub(r'[^\w\s]', '_', name)  # 특수문자를 언더스코어로 변환
        name = re.sub(r'\s+', '_', name)      # 공백을 언더스코어로 변환
        
        # SQLite 예약어 처리
        reserved_words = ['add', 'all', 'alter', 'and', 'as', 'autoincrement', 'between', 'case', 'check', 'collate', 
                         'commit', 'constraint', 'create', 'default', 'deferrable', 'delete', 'distinct', 'drop', 
                         'else', 'escape', 'except', 'exists', 'foreign', 'from', 'group', 'having', 'if', 'in', 
                         'index', 'insert', 'intersect', 'into', 'is', 'isnull', 'join', 'limit', 'not', 'notnull', 
                         'null', 'on', 'or', 'order', 'primary', 'references', 'select', 'set', 'table', 'then', 
                         'to', 'transaction', 'union', 'unique', 'update', 'using', 'values', 'when', 'where']
        
        if name.lower() in reserved_words:
            name = name + '_col'
        
        # 숫자로 시작하는 경우 앞에 'col_' 추가
        if name and name[0].isdigit():
            name = 'col_' + name
            
        return name
    
    # SQLite 데이터베이스 연결
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    
    # 각 엑셀 설정에 대해 처리
    for table_name, config in excel_configs.items():
        log(f"테이블 '{table_name}' 처리 중...")
        
        # 엑셀 파일 열기
        try:
            log(f"엑셀 파일 '{config['path']}' 열기...")
            workbook = openpyxl.load_workbook(config['path'], read_only=True, data_only=True)
            sheet = workbook[config['sheet_name']]
            log(f"시트 '{config['sheet_name']}' 로드 완료")
        except Exception as e:
            log(f"엑셀 파일 '{config['path']}' 열기 실패: {e}")
            continue
        
        # 헤더 행 결정 (book2는 두 번째 행을 헤더로 사용)
        header_row_num = 1  # 기본값은 첫 번째 행
        if 'header_row' in config:
            header_row_num = config['header_row']
        
        # 헤더 행에서 컬럼 이름 가져오기
        header_row = next(sheet.iter_rows(min_row=header_row_num, max_row=header_row_num))
        column_names = []
        column_indices = []  # 유효한 컬럼의 인덱스를 저장
        column_name_counts = {}  # 중복 컬럼 이름 처리를 위한 딕셔너리
        
        for idx, cell in enumerate(header_row):
            col_name = clean_column_name(cell.value)
            
            # 빈 컬럼 이름은 무시
            if not col_name:
                continue
                
            # 중복 컬럼 이름 처리
            if col_name in column_name_counts:
                column_name_counts[col_name] += 1
                col_name = f"{col_name}_{column_name_counts[col_name]}"
            else:
                column_name_counts[col_name] = 0
                
            column_names.append(col_name)
            column_indices.append(idx)  # 유효한 컬럼의 인덱스 저장
        
        log(f"컬럼: {', '.join(column_names)}")
        
        # 테이블이 존재하면 삭제
        cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
        
        # 테이블 생성
        create_table_sql = f"CREATE TABLE {table_name} ({', '.join([f'{col} TEXT' for col in column_names])})"
        cursor.execute(create_table_sql)
        log(f"테이블 '{table_name}' 생성 완료")
        
        # 데이터 시작 행 결정
        data_start_row = header_row_num + 1
        
        # 데이터 삽입
        rows = []
        row_count = 0
        for row in sheet.iter_rows(min_row=data_start_row):  # 헤더 다음 행부터 시작
            row_data = []
            for idx in column_indices:  # 유효한 컬럼만 처리
                cell_value = row[idx].value
                row_data.append(str(cell_value) if cell_value is not None else "")
            
            if any(row_data):  # 빈 행은 건너뜀
                rows.append(row_data)
                row_count += 1
                
                # 1000행마다 데이터베이스에 삽입하고 로그 출력
                if row_count % 1000 == 0:
                    log(f"{row_count}행 처리 중...")
        
        # 데이터 일괄 삽입
        placeholders = ', '.join(['?' for _ in column_names])
        insert_sql = f"INSERT INTO {table_name} ({', '.join(column_names)}) VALUES ({placeholders})"
        cursor.executemany(insert_sql, rows)
        log(f"총 {len(rows)}행 삽입 완료")
        
        # 워크북 닫기
        workbook.close()
    
    # 변경사항 저장 및 연결 종료
    conn.commit()
    conn.close()
    
    log(f"데이터베이스 '{db_file}'에 성공적으로 데이터를 저장했습니다.")
    return True


class ExcelToSqliteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to SQLite 변환기")
        self.root.geometry("800x600")
        
        # 기본 설정
        self.excel_configs = {}
        self.db_file = "mgui2.db"
        
        # UI 구성
        self.create_widgets()
        
        # 기본 설정 로드
        self.load_default_config()
        
    def create_widgets(self):
        # 프레임 생성
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        config_frame = ttk.LabelFrame(self.root, text="설정", padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(self.root, text="로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 상단 버튼
        self.convert_btn = ttk.Button(top_frame, text="엑셀 -> SQLite 변환", command=self.convert)
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        self.select_excel_btn = ttk.Button(top_frame, text="엑셀 파일 선택", command=self.select_excel_file)
        self.select_excel_btn.pack(side=tk.LEFT, padx=5)
        
        self.select_db_btn = ttk.Button(top_frame, text="DB 파일 선택", command=self.select_db_file)
        self.select_db_btn.pack(side=tk.LEFT, padx=5)
        
        # 설정 표시 영역
        self.config_text = ScrolledText(config_frame, height=10)
        self.config_text.pack(fill=tk.BOTH, expand=True)
        
        # 로그 표시 영역
        self.log_text = ScrolledText(log_frame, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def load_default_config(self):
        # 기본 설정 로드
        self.excel_configs = { 
            'book1': {
                'path': 'C:/work/doc/1.xlsx',
                'sheet_name': '1',
                'header_row': 1  # 첫 번째 행이 헤더
            },
            'book2': {
                'path': 'C:/work/doc/2.xlsx',
                'sheet_name': '1',
                'header_row': 2  # 두 번째 행이 헤더
            }
        }
        self.update_config_display()
        
    def update_config_display(self):
        # 설정 정보 표시
        self.config_text.delete(1.0, tk.END)
        self.config_text.insert(tk.END, f"데이터베이스 파일: {self.db_file}\n\n")
        
        for table_name, config in self.excel_configs.items():
            self.config_text.insert(tk.END, f"테이블: {table_name}\n")
            self.config_text.insert(tk.END, f"  엑셀 파일: {config['path']}\n")
            self.config_text.insert(tk.END, f"  시트 이름: {config['sheet_name']}\n")
            self.config_text.insert(tk.END, f"  헤더 행: {config.get('header_row', 1)}\n")
            self.config_text.insert(tk.END, "\n")
    
    def log(self, message):
        # 로그 메시지 추가
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 스크롤을 가장 아래로 이동
        self.root.update()  # UI 업데이트
    
    def select_excel_file(self):
        # 엑셀 파일 선택
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )
        
        if file_path:
            # 간단한 구현을 위해 첫 번째 테이블의 경로만 변경
            table_name = list(self.excel_configs.keys())[0]
            self.excel_configs[table_name]['path'] = file_path
            self.update_config_display()
            self.log(f"엑셀 파일 경로가 {file_path}로 변경되었습니다.")
    
    def select_db_file(self):
        # 데이터베이스 파일 선택
        file_path = filedialog.asksaveasfilename(
            title="SQLite 데이터베이스 파일 선택",
            defaultextension=".db",
            filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.db_file = file_path
            self.update_config_display()
            self.log(f"데이터베이스 파일 경로가 {file_path}로 변경되었습니다.")
    
    def convert(self):
        # 변환 시작
        self.log("변환 시작...")
        self.convert_btn.config(state=tk.DISABLED)
        
        try:
            # 변환 실행
            success = excel_to_sqlite(self.excel_configs, self.db_file, self.log)
            
            if success:
                messagebox.showinfo("완료", f"변환이 완료되었습니다.\n데이터베이스 파일: {self.db_file}")
            
        except Exception as e:
            self.log(f"오류 발생: {e}")
            messagebox.showerror("오류", f"변환 중 오류가 발생했습니다: {e}")
        
        finally:
            self.convert_btn.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToSqliteApp(root)
    root.mainloop()