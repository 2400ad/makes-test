import os
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.scrolledtext import ScrolledText
import openpyxl
import re

def excel_to_sqlite(excel_configs, db_file='test.db', log_callback=None):
    """
    엑셀 파일을 SQLite 데이터베이스로 변환합니다.
    
    Args:
        excel_configs (dict): 엑셀 파일 설정 정보
        db_file (str): SQLite 데이터베이스 파일 경로
        log_callback (function): 로그 출력 콜백 함수
    
    Returns:
        bool: 성공 여부
    """
    def log(message):
        if log_callback:
            log_callback(message)
        else:
            print(message)
    
    def clean_column_name(name, used_names=None):
        if used_names is None:
            used_names = set()
        
        # 빈 컬럼명 처리
        if not name or name.strip() == "":
            return None  # 빈 컬럼명은 None 반환하여 처리하지 않음
        
        # 공백 및 특수문자 제거, 소문자로 변환
        name = str(name).strip()
        name = re.sub(r'[^\w\s]', '_', name)  # 특수문자를 언더스코어로 변환
        name = re.sub(r'\s+', '_', name)      # 공백을 언더스코어로 변환
        
        # 숫자로 시작하는 경우 접두어 추가
        if name[0].isdigit():
            name = 'col_' + name
        
        # SQLite 예약어 확인 및 처리
        sqlite_keywords = ['ABORT', 'ACTION', 'ADD', 'AFTER', 'ALL', 'ALTER', 'ANALYZE', 'AND', 'AS', 'ASC', 
                          'ATTACH', 'AUTOINCREMENT', 'BEFORE', 'BEGIN', 'BETWEEN', 'BY', 'CASCADE', 'CASE', 
                          'CAST', 'CHECK', 'COLLATE', 'COLUMN', 'COMMIT', 'CONFLICT', 'CONSTRAINT', 'CREATE', 
                          'CROSS', 'CURRENT_DATE', 'CURRENT_TIME', 'CURRENT_TIMESTAMP', 'DATABASE', 'DEFAULT', 
                          'DEFERRABLE', 'DEFERRED', 'DELETE', 'DESC', 'DETACH', 'DISTINCT', 'DROP', 'EACH', 
                          'ELSE', 'END', 'ESCAPE', 'EXCEPT', 'EXCLUSIVE', 'EXISTS', 'EXPLAIN', 'FAIL', 'FOR', 
                          'FOREIGN', 'FROM', 'FULL', 'GLOB', 'GROUP', 'HAVING', 'IF', 'IGNORE', 'IMMEDIATE', 
                          'IN', 'INDEX', 'INDEXED', 'INITIALLY', 'INNER', 'INSERT', 'INSTEAD', 'INTERSECT', 
                          'INTO', 'IS', 'ISNULL', 'JOIN', 'KEY', 'LEFT', 'LIKE', 'LIMIT', 'MATCH', 'NATURAL', 
                          'NO', 'NOT', 'NOTNULL', 'NULL', 'OF', 'OFFSET', 'ON', 'OR', 'ORDER', 'OUTER', 'PLAN', 
                          'PRAGMA', 'PRIMARY', 'QUERY', 'RAISE', 'RECURSIVE', 'REFERENCES', 'REGEXP', 'REINDEX', 
                          'RELEASE', 'RENAME', 'REPLACE', 'RESTRICT', 'RIGHT', 'ROLLBACK', 'ROW', 'SAVEPOINT', 
                          'SELECT', 'SET', 'TABLE', 'TEMP', 'TEMPORARY', 'THEN', 'TO', 'TRANSACTION', 'TRIGGER', 
                          'UNION', 'UNIQUE', 'UPDATE', 'USING', 'VACUUM', 'VALUES', 'VIEW', 'VIRTUAL', 'WHEN', 
                          'WHERE', 'WITH', 'WITHOUT']
        
        if name.upper() in sqlite_keywords:
            name = name + '_col'
        
        # 중복 컬럼명 처리
        original_name = name
        counter = 1
        while name.lower() in [n.lower() for n in used_names]:
            name = f"{original_name}_{counter}"
            counter += 1
        
        used_names.add(name)
        return name
    
    try:
        # SQLite 데이터베이스 연결
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        for table_name, config in excel_configs.items():
            excel_path = config['path']
            sheet_name = config['sheet_name']
            header_row = config.get('header_row', 1)  # 기본값은 1 (첫 번째 행)
            
            log(f"엑셀 파일 '{excel_path}' 처리 중...")
            
            # 엑셀 파일 로드
            if not os.path.exists(excel_path):
                log(f"엑셀 파일 '{excel_path}'이 존재하지 않습니다.")
                continue
            
            workbook = openpyxl.load_workbook(excel_path, data_only=True)
            
            if sheet_name not in workbook.sheetnames:
                log(f"시트 '{sheet_name}'이 존재하지 않습니다.")
                continue
            
            sheet = workbook[sheet_name]
            
            # 헤더 행 읽기 (컬럼명)
            header_cells = list(sheet.rows)[header_row - 1]
            used_column_names = set()
            columns = []
            
            for cell in header_cells:
                column_name = clean_column_name(cell.value, used_column_names)
                if column_name:  # 빈 컬럼명이 아닌 경우만 추가
                    columns.append(column_name)
            
            if not columns:
                log(f"유효한 컬럼이 없습니다.")
                continue
            
            # 테이블 생성
            columns_sql = ', '.join([f'"{col}" TEXT' for col in columns])
            cursor.execute(f'DROP TABLE IF EXISTS {table_name}')
            cursor.execute(f'CREATE TABLE {table_name} ({columns_sql})')
            
            log(f"테이블 '{table_name}' 생성됨")
            
            # 데이터 삽입
            data_rows = list(sheet.rows)[header_row:]  # 헤더 다음 행부터 데이터
            
            for row in data_rows:
                values = []
                for i, cell in enumerate(row):
                    if i < len(columns):  # 컬럼 수만큼만 처리
                        value = cell.value if cell.value is not None else ""
                        values.append(value)
                
                if len(values) < len(columns):
                    values.extend([""] * (len(columns) - len(values)))
                
                placeholders = ', '.join(['?' for _ in columns])
                cursor.execute(f'INSERT INTO {table_name} VALUES ({placeholders})', values[:len(columns)])
            
            log(f"테이블 '{table_name}'에 {len(data_rows)}개의 행이 삽입됨")
        
        # 변경사항 저장
        conn.commit()
        log("데이터베이스 저장 완료")
        
        # 연결 종료
        conn.close()
        return True
        
    except Exception as e:
        log(f"오류 발생: {e}")
        return False

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
        
        # 상단 버튼 및 검색 UI
        self.convert_btn = ttk.Button(top_frame, text="엑셀 -> SQLite 변환", command=self.convert)
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # 매핑SEQ 검색 UI
        ttk.Label(top_frame, text="매핑SEQ:").pack(side=tk.LEFT, padx=(20, 5))
        
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=15)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind("<Return>", lambda event: self.search_mapping_seq())
        
        self.search_btn = ttk.Button(top_frame, text="Search", command=self.search_mapping_seq)
        self.search_btn.pack(side=tk.LEFT, padx=5)
        
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
        self.config_text.insert(tk.END, f"데이터베이스 파일: {self.db_file}\n")
        
        for table_name, config in self.excel_configs.items():
            self.config_text.insert(tk.END, f"테이블: {table_name}\n")
            self.config_text.insert(tk.END, f"  엑셀 파일: {config['path']}\n")
            self.config_text.insert(tk.END, f"  시트 이름: {config['sheet_name']}\n")
            self.config_text.insert(tk.END, f"  헤더 행: {config.get('header_row', 1)}\n")
    
    def log(self, message):
        # 로그 메시지 추가
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # 스크롤을 가장 아래로 이동
        self.root.update()  # UI 업데이트
    
    def search_mapping_seq(self):
        # 매핑SEQ로 검색하는 함수
        mapping_seq = self.search_var.get().strip()
        if not mapping_seq:
            messagebox.showwarning("검색 오류", "매핑SEQ를 입력하세요.")
            return
        
        try:
            # SQLite 데이터베이스 연결
            if not os.path.exists(self.db_file):
                messagebox.showwarning("검색 오류", f"데이터베이스 파일({self.db_file})이 존재하지 않습니다. 먼저 변환을 실행하세요.")
                return
                
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # 테이블 존재 여부 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='book2'")
            if not cursor.fetchone():
                messagebox.showwarning("검색 오류", "book2 테이블이 존재하지 않습니다. 먼저 변환을 실행하세요.")
                conn.close()
                return
            
            # 컬럼 정보 가져오기
            cursor.execute(f"PRAGMA table_info(book2)")
            columns = [column[1] for column in cursor.fetchall()]
            
            # 매핑SEQ 컬럼 존재 여부 확인
            if "매핑SEQ" not in columns:
                # 컬럼명이 정리되었을 수 있으므로 가능한 변형 확인
                mapping_seq_col = None
                for col in columns:
                    if col.lower() == "매핑seq" or col.lower() == "매핑_seq" or col.lower() == "매핑seq_" or col.lower() == "매핑_seq_":
                        mapping_seq_col = col
                        break
                
                if not mapping_seq_col:
                    messagebox.showwarning("검색 오류", "매핑SEQ 컬럼이 존재하지 않습니다.")
                    conn.close()
                    return
            else:
                mapping_seq_col = "매핑SEQ"
            
            # GroupID와 EventID 컬럼 확인
            group_id_col = None
            event_id_col = None
            for col in columns:
                if "group" in col.lower() and "id" in col.lower():
                    group_id_col = col
                if "event" in col.lower() and "id" in col.lower():
                    event_id_col = col
            
            # 송신 및 수신 관련 컬럼 찾기
            send_task_col = None
            send_qmgr_col = None
            recv_task_col = None
            recv_qmgr_col = None
            
            # SQL 관련 컬럼 찾기
            send_userid_col = None
            send_passwd_col = None
            send_db_col = None
            send_schema_adapter_col = None
            send_table_adapter_col = None
            
            recv_userid_col = None
            recv_passwd_col = None
            recv_db_col = None
            recv_schema_adapter_col = None
            recv_table_adapter_col = None
            
            for col in columns:
                # 송신 관련 컬럼
                if "송신" in col and "업무" in col:
                    send_task_col = col
                if "송신" in col and "qmgr" in col.lower():
                    send_qmgr_col = col
                if "송신" in col and "userid" in col.lower():
                    send_userid_col = col
                if "송신" in col and ("passwd" in col.lower() or "password" in col.lower()):
                    send_passwd_col = col
                if "송신" in col and "db" in col.lower():
                    send_db_col = col
                if "송신" in col and "schema" in col.lower() and "adapter" in col.lower():
                    send_schema_adapter_col = col
                if "송신" in col and "table" in col.lower() and "adapter" in col.lower():
                    send_table_adapter_col = col
                
                # 수신 관련 컬럼
                if "수신" in col and "업무" in col:
                    recv_task_col = col
                if "수신" in col and "qmgr" in col.lower():
                    recv_qmgr_col = col
                if "수신" in col and "userid" in col.lower():
                    recv_userid_col = col
                if "수신" in col and ("passwd" in col.lower() or "password" in col.lower()):
                    recv_passwd_col = col
                if "수신" in col and "db" in col.lower():
                    recv_db_col = col
                if "수신" in col and "schema" in col.lower() and "adapter" in col.lower():
                    recv_schema_adapter_col = col
                if "수신" in col and "table" in col.lower() and "adapter" in col.lower():
                    recv_table_adapter_col = col
            
            # 검색 실행
            query = f"SELECT * FROM book2 WHERE {mapping_seq_col} = ?"
            cursor.execute(query, (mapping_seq,))
            row = cursor.fetchone()
            
            if not row:
                messagebox.showinfo("검색 결과", f"매핑SEQ '{mapping_seq}'에 해당하는 데이터가 없습니다.")
                conn.close()
                return
            
            # 결과 표시
            self.config_text.delete(1.0, tk.END)
            
            # 매핑SEQ 표시
            self.config_text.insert(tk.END, f"[매핑SEQ] {row[columns.index(mapping_seq_col)]}\n")
            
            # INTERFACE ID 표시 (GroupID.EventID)
            interface_id = ""
            if group_id_col and event_id_col:
                group_id = row[columns.index(group_id_col)]
                event_id = row[columns.index(event_id_col)]
                interface_id = f"{group_id}.{event_id}"
            
            self.config_text.insert(tk.END, f"[INTERFACE ID] {interface_id}\n")
            
            # 송신 정보 표시
            send_info = ""
            if send_task_col:
                send_task = row[columns.index(send_task_col)]
                send_info = f"[송신_업무명] {send_task}"
            
            if send_qmgr_col:
                send_qmgr = row[columns.index(send_qmgr_col)]
                if send_info:
                    send_info += f"    [송신_QMGR명] {send_qmgr}"
                else:
                    send_info = f"[송신_QMGR명] {send_qmgr}"
            
            if send_info:
                self.config_text.insert(tk.END, f"\n{send_info}")
            
            # 수신 정보 표시
            recv_info = ""
            if recv_task_col:
                recv_task = row[columns.index(recv_task_col)]
                recv_info = f"[수신_업무명] {recv_task}"
            
            if recv_qmgr_col:
                recv_qmgr = row[columns.index(recv_qmgr_col)]
                if recv_info:
                    recv_info += f"    [수신_QMGR명] {recv_qmgr}"
                else:
                    recv_info = f"[수신_QMGR명] {recv_qmgr}"
            
            if recv_info:
                self.config_text.insert(tk.END, f"\n{recv_info}")
            
            # SQL 정보 추가
            # 송신 SQL
            send_sql = "\n[송신SQL]"
            
            send_userid = row[columns.index(send_userid_col)] if send_userid_col and columns.index(send_userid_col) < len(row) else "?"
            send_passwd = row[columns.index(send_passwd_col)] if send_passwd_col and columns.index(send_passwd_col) < len(row) else "?"
            send_db = row[columns.index(send_db_col)] if send_db_col and columns.index(send_db_col) < len(row) else "?"
            send_schema = row[columns.index(send_schema_adapter_col)] if send_schema_adapter_col and columns.index(send_schema_adapter_col) < len(row) else "?"
            send_table = row[columns.index(send_table_adapter_col)] if send_table_adapter_col and columns.index(send_table_adapter_col) < len(row) else "?"
            
            send_sql += f"\nsqlplus {send_userid}/{send_passwd}@{send_db}"
            send_sql += f"\nselect count(*) from {send_schema}.{send_table} where EAI_TRANSFER_DATE > sysdate-(1/12) and EAI_TRANSFER_FLAG='Y';"
            
            self.config_text.insert(tk.END, send_sql)
            
            # 수신 SQL
            recv_sql = "\n[수신SQL]"
            
            recv_userid = row[columns.index(recv_userid_col)] if recv_userid_col and columns.index(recv_userid_col) < len(row) else "?"
            recv_passwd = row[columns.index(recv_passwd_col)] if recv_passwd_col and columns.index(recv_passwd_col) < len(row) else "?"
            recv_db = row[columns.index(recv_db_col)] if recv_db_col and columns.index(recv_db_col) < len(row) else "?"
            recv_schema = row[columns.index(recv_schema_adapter_col)] if recv_schema_adapter_col and columns.index(recv_schema_adapter_col) < len(row) else "?"
            recv_table = row[columns.index(recv_table_adapter_col)] if recv_table_adapter_col and columns.index(recv_table_adapter_col) < len(row) else "?"
            
            recv_sql += f"\nsqlplus {recv_userid}/{recv_passwd}@{recv_db}"
            recv_sql += f"\nselect count(*) from {recv_schema}.{recv_table} where EAI_TRANSFER_DATE > sysdate-(1/12);"
            
            self.config_text.insert(tk.END, recv_sql)
            
            # 추가 정보 표시 (모든 컬럼 표시)
            self.config_text.insert(tk.END, "\n\n--- 추가 정보 ---")
            for i, col in enumerate(columns):
                if col not in [mapping_seq_col, group_id_col, event_id_col, 
                              send_task_col, send_qmgr_col, recv_task_col, recv_qmgr_col,
                              send_userid_col, send_passwd_col, send_db_col, 
                              send_schema_adapter_col, send_table_adapter_col,
                              recv_userid_col, recv_passwd_col, recv_db_col,
                              recv_schema_adapter_col, recv_table_adapter_col]:
                    value = row[i] if i < len(row) else ""
                    self.config_text.insert(tk.END, f"\n[{col}] {value}")
            
            conn.close()
            self.log(f"매핑SEQ '{mapping_seq}' 검색 완료")
            
        except Exception as e:
            self.log(f"검색 중 오류 발생: {e}")
            messagebox.showerror("검색 오류", f"검색 중 오류가 발생했습니다: {e}")
    
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