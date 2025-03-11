import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from tkinter.scrolledtext import ScrolledText
import openpyxl
import re
import threading
import queue
import time

def excel_to_sqlite(excel_configs, db_file='test.db', log_callback=None):
    if log_callback:
        log_callback("Excel to SQLite 변환 시작...")
    
    # 기존 DB 파일 삭제
    if os.path.exists(db_file):
        os.remove(db_file)
        if log_callback:
            log_callback(f"{db_file} 파일 삭제됨")
    
    # SQLite 연결
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()
    
    def clean_column_name(name):
        # 컬럼명에서 특수문자 제거하고 공백을 언더스코어로 변경
        if name is None:
            return "unnamed"
        
        # 문자열로 변환
        name = str(name).strip()
        
        # 빈 문자열이면 unnamed 반환
        if not name:
            return "unnamed"
        
        # 특수문자 제거하고 공백을 언더스코어로 변경
        name = re.sub(r'[^\w\s]', '', name)
        name = re.sub(r'\s+', '_', name)
        
        return name
    
    # 각 Excel 파일 처리
    for idx, config in enumerate(excel_configs):
        excel_file = config.get('file')
        sheet_name = config.get('sheet')
        table_name = f"book{idx+1}"
        
        if not excel_file or not os.path.exists(excel_file):
            if log_callback:
                log_callback(f"파일이 존재하지 않음: {excel_file}")
            continue
        
        if log_callback:
            log_callback(f"{excel_file} 처리 중...")
        
        # Excel 파일 로드
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        
        # 시트 선택
        if sheet_name and sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            sheet = wb.active
        
        # 데이터 읽기
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)
        
        if not data:
            if log_callback:
                log_callback(f"{excel_file}에 데이터가 없음")
            continue
        
        # 첫 번째 행을 컬럼명으로 사용
        headers = data[0]
        
        # 컬럼명 정리 (중복 제거 및 특수문자 처리)
        clean_headers = []
        header_counts = {}
        
        for h in headers:
            clean_h = clean_column_name(h)
            
            # 중복된 컬럼명 처리
            if clean_h in header_counts:
                header_counts[clean_h] += 1
                clean_h = f"{clean_h}_{header_counts[clean_h]}"
            else:
                header_counts[clean_h] = 0
            
            clean_headers.append(clean_h)
        
        # book2의 경우 첫 번째 행 건너뛰기
        start_idx = 1
        if idx == 1:  # book2
            start_idx = 1
        
        # 테이블 생성
        columns_str = ", ".join([f'"{h}" TEXT' for h in clean_headers if h])
        if columns_str:
            cursor.execute(f'CREATE TABLE IF NOT EXISTS {table_name} ({columns_str})')
            
            # 데이터 삽입
            for row_idx in range(start_idx, len(data)):
                row_data = data[row_idx]
                # 행 데이터가 컬럼 수보다 적을 경우 None으로 채움
                while len(row_data) < len(clean_headers):
                    row_data = row_data + (None,)
                
                # 컬럼명이 비어있지 않은 컬럼만 사용
                valid_headers = [h for h in clean_headers if h]
                valid_data = [row_data[i] for i, h in enumerate(clean_headers) if h]
                
                placeholders = ", ".join(["?" for _ in valid_headers])
                insert_query = f'INSERT INTO {table_name} ({", ".join([f\'"{h}"' for h in valid_headers])}) VALUES ({placeholders})'
                
                cursor.execute(insert_query, valid_data)
        
        if log_callback:
            log_callback(f"{table_name} 테이블 생성 완료")
    
    conn.commit()
    conn.close()
    
    if log_callback:
        log_callback("Excel to SQLite 변환 완료")
    
    return db_file

class ExcelToSqliteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to SQLite Converter")
        self.root.geometry("800x600")
        
        self.excel_configs = []
        self.db_file = "test.db"
        self.log_queue = queue.Queue()
        self.create_widgets()
        self.process_log_queue()
    
    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 상단 프레임 (파일 선택 및 변환 버튼)
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(top_frame, text="Excel 파일 1:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5), pady=5)
        self.excel1_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.excel1_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="찾아보기", command=lambda: self.browse_file(self.excel1_var)).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(top_frame, text="시트명:").grid(row=0, column=3, sticky=tk.W, padx=(10, 5), pady=5)
        self.sheet1_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.sheet1_var, width=15).grid(row=0, column=4, padx=5, pady=5)
        
        ttk.Label(top_frame, text="Excel 파일 2:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=5)
        self.excel2_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.excel2_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="찾아보기", command=lambda: self.browse_file(self.excel2_var)).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Label(top_frame, text="시트명:").grid(row=1, column=3, sticky=tk.W, padx=(10, 5), pady=5)
        self.sheet2_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.sheet2_var, width=15).grid(row=1, column=4, padx=5, pady=5)
        
        ttk.Label(top_frame, text="SQLite DB 파일:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=5)
        self.db_var = tk.StringVar(value="test.db")
        ttk.Entry(top_frame, textvariable=self.db_var, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="찾아보기", command=self.browse_db_file).grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Button(top_frame, text="변환", command=self.convert).grid(row=2, column=4, padx=5, pady=5)
        
        # 중간 프레임 (설정)
        config_frame = ttk.LabelFrame(main_frame, text="설정")
        config_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 설정 내용 (스크롤 텍스트)
        self.config_text = ScrolledText(config_frame, height=20)  # 높이를 2배로 늘림
        self.config_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 검색 프레임
        search_frame = ttk.Frame(config_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        ttk.Label(search_frame, text="매핑SEQ 검색:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.search_var, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="검색", command=self.search_mapping_seq).pack(side=tk.LEFT, padx=5)
        
        # 하단 프레임 (로그)
        log_frame = ttk.LabelFrame(main_frame, text="로그")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 로그 내용 (스크롤 텍스트)
        self.log_text = ScrolledText(log_frame, height=5)  # 높이를 1/3로 줄임
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def browse_file(self, var):
        filename = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
    
    def browse_db_file(self):
        filename = filedialog.asksaveasfilename(
            title="SQLite DB 파일 저장",
            defaultextension=".db",
            filetypes=[("SQLite files", "*.db"), ("All files", "*.*")]
        )
        if filename:
            self.db_var.set(filename)
    
    def convert(self):
        # 설정 가져오기
        excel1 = self.excel1_var.get()
        sheet1 = self.sheet1_var.get()
        excel2 = self.excel2_var.get()
        sheet2 = self.sheet2_var.get()
        db_file = self.db_var.get()
        
        if not excel1 and not excel2:
            messagebox.showerror("오류", "최소한 하나의 Excel 파일을 선택해야 합니다.")
            return
        
        # Excel 설정 구성
        self.excel_configs = []
        if excel1:
            self.excel_configs.append({'file': excel1, 'sheet': sheet1})
        if excel2:
            self.excel_configs.append({'file': excel2, 'sheet': sheet2})
        
        self.db_file = db_file
        
        # 변환 스레드 시작
        threading.Thread(target=self.run_conversion, daemon=True).start()
    
    def run_conversion(self):
        try:
            excel_to_sqlite(self.excel_configs, self.db_file, self.log)
            self.root.after(0, lambda: messagebox.showinfo("완료", "Excel 파일이 SQLite DB로 변환되었습니다."))
        except Exception as e:
            self.log(f"오류 발생: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("오류", f"변환 중 오류가 발생했습니다: {str(e)}"))
    
    def log(self, message):
        self.log_queue.put(message)
    
    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_queue.task_done()
        except queue.Empty:
            self.root.after(100, self.process_log_queue)
    
    def search_mapping_seq(self):
        mapping_seq = self.search_var.get().strip()
        if not mapping_seq:
            messagebox.showinfo("검색", "매핑SEQ를 입력하세요.")
            return
        
        if not os.path.exists(self.db_file):
            messagebox.showinfo("검색", "먼저 Excel 파일을 변환해야 합니다.")
            return
        
        try:
            # SQLite 연결
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # 테이블 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='book2'")
            book2_exists = cursor.fetchone() is not None
            
            if not book2_exists:
                messagebox.showinfo("검색", "book2 테이블이 존재하지 않습니다.")
                conn.close()
                return
            
            # book1 테이블 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='book1'")
            book1_exists = cursor.fetchone() is not None
            
            # 컬럼 목록 가져오기
            cursor.execute("PRAGMA table_info(book2)")
            book2_columns = [column[1] for column in cursor.fetchall()]
            
            book1_columns = []
            book1_mapping_seq_col = None
            interface_name_col = None
            if_type_col = None
            route_col = None
            
            if book1_exists:
                cursor.execute("PRAGMA table_info(book1)")
                book1_columns = [column[1] for column in cursor.fetchall()]
                
                # book1에서 매핑SEQ 컬럼 찾기
                for col in book1_columns:
                    if '매핑seq' in col.lower() or '매핑_seq' in col.lower() or 'seq' in col.lower():
                        book1_mapping_seq_col = col
                        break
                
                # 인터페이스_명 컬럼 찾기
                for col in book1_columns:
                    if '인터페이스_명' in col or '인터페이스명' in col:
                        interface_name_col = col
                        break
                
                # I_F_Type 컬럼 찾기
                for col in book1_columns:
                    if 'i_f_type' in col.lower() or 'if_type' in col.lower() or 'iftype' in col.lower():
                        if_type_col = col
                        break
                
                # Route정의 컬럼 찾기
                for col in book1_columns:
                    if 'route' in col.lower() or '라우트' in col.lower() or '정의' in col.lower():
                        route_col = col
                        break
            
            # 필요한 컬럼 찾기
            mapping_seq_col = None
            group_id_col = None
            event_id_col = None
            send_task_col = None
            send_qmgr_col = None
            recv_task_col = None
            recv_qmgr_col = None
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
            
            for col in book2_columns:
                if '매핑seq' in col.lower() or '매핑_seq' in col.lower() or 'seq' in col.lower():
                    mapping_seq_col = col
                elif 'group' in col.lower() or '그룹' in col.lower():
                    group_id_col = col
                elif 'event' in col.lower() or '이벤트' in col.lower():
                    event_id_col = col
                elif '송신' in col and ('업무' in col or 'task' in col.lower()):
                    send_task_col = col
                elif '송신' in col and ('qmgr' in col.lower() or 'queue' in col.lower()):
                    send_qmgr_col = col
                elif '수신' in col and ('업무' in col or 'task' in col.lower()):
                    recv_task_col = col
                elif '수신' in col and ('qmgr' in col.lower() or 'queue' in col.lower()):
                    recv_qmgr_col = col
                elif '송신' in col and ('userid' in col.lower() or 'id' in col.lower()):
                    send_userid_col = col
                elif '송신' in col and ('passwd' in col.lower() or 'pw' in col.lower() or 'password' in col.lower()):
                    send_passwd_col = col
                elif '송신' in col and ('db' in col.lower() or 'database' in col.lower()):
                    send_db_col = col
                elif '송신' in col and ('schema' in col.lower() or '스키마' in col):
                    send_schema_adapter_col = col
                elif '송신' in col and ('table' in col.lower() or '테이블' in col):
                    send_table_adapter_col = col
                elif '수신' in col and ('userid' in col.lower() or 'id' in col.lower()):
                    recv_userid_col = col
                elif '수신' in col and ('passwd' in col.lower() or 'pw' in col.lower() or 'password' in col.lower()):
                    recv_passwd_col = col
                elif '수신' in col and ('db' in col.lower() or 'database' in col.lower()):
                    recv_db_col = col
                elif '수신' in col and ('schema' in col.lower() or '스키마' in col):
                    recv_schema_adapter_col = col
                elif '수신' in col and ('table' in col.lower() or '테이블' in col):
                    recv_table_adapter_col = col
            
            if not mapping_seq_col:
                messagebox.showinfo("검색", "매핑SEQ 컬럼을 찾을 수 없습니다.")
                conn.close()
                return
            
            # book2에서 검색 실행
            query = f"SELECT * FROM book2 WHERE {mapping_seq_col} = ?"
            cursor.execute(query, (mapping_seq,))
            book2_row = cursor.fetchone()
            
            # book1에서 검색 실행
            book1_row = None
            if book1_exists and book1_mapping_seq_col:
                query = f"SELECT * FROM book1 WHERE {book1_mapping_seq_col} = ?"
                cursor.execute(query, (mapping_seq,))
                book1_row = cursor.fetchone()
            
            # book2에도 없고 book1에도 없으면 메시지 표시 후 종료
            if not book2_row and not book1_row:
                messagebox.showinfo("검색 결과", f"매핑SEQ '{mapping_seq}'에 해당하는 데이터가 없습니다.")
                conn.close()
                return
            
            # 결과 표시
            self.config_text.delete(1.0, tk.END)
            
            # 매핑SEQ 표시
            self.config_text.insert(tk.END, f"[매핑SEQ] {mapping_seq}\n")
            
            # book2 데이터가 있는 경우 해당 정보 표시
            if book2_row:
                # INTERFACE ID 표시 (GroupID.EventID)
                interface_id = ""
                if group_id_col and event_id_col:
                    group_id = book2_row[book2_columns.index(group_id_col)]
                    event_id = book2_row[book2_columns.index(event_id_col)]
                    interface_id = f"{group_id}.{event_id}"
                
                self.config_text.insert(tk.END, f"[INTERFACE ID] {interface_id}\n")
                
                # 송신 정보 표시
                send_info = ""
                if send_task_col:
                    send_task = book2_row[book2_columns.index(send_task_col)]
                    send_info = f"[송신_업무명] {send_task}"
                
                if send_qmgr_col:
                    send_qmgr = book2_row[book2_columns.index(send_qmgr_col)]
                    if send_info:
                        send_info += f"    [송신_QMGR명] {send_qmgr}"
                    else:
                        send_info = f"[송신_QMGR명] {send_qmgr}"
                
                if send_info:
                    self.config_text.insert(tk.END, f"\n{send_info}")
                
                # 수신 정보 표시
                recv_info = ""
                if recv_task_col:
                    recv_task = book2_row[book2_columns.index(recv_task_col)]
                    recv_info = f"[수신_업무명] {recv_task}"
                
                if recv_qmgr_col:
                    recv_qmgr = book2_row[book2_columns.index(recv_qmgr_col)]
                    if recv_info:
                        recv_info += f"    [수신_QMGR명] {recv_qmgr}"
                    else:
                        recv_info = f"[수신_QMGR명] {recv_qmgr}"
                
                if recv_info:
                    self.config_text.insert(tk.END, f"\n{recv_info}")
                
                # SQL 정보 추가
                # 송신 SQL
                send_sql = "\n[송신SQL]"
                
                send_userid = book2_row[book2_columns.index(send_userid_col)] if send_userid_col and book2_columns.index(send_userid_col) < len(book2_row) else "?"
                send_passwd = book2_row[book2_columns.index(send_passwd_col)] if send_passwd_col and book2_columns.index(send_passwd_col) < len(book2_row) else "?"
                send_db = book2_row[book2_columns.index(send_db_col)] if send_db_col and book2_columns.index(send_db_col) < len(book2_row) else "?"
                send_schema = book2_row[book2_columns.index(send_schema_adapter_col)] if send_schema_adapter_col and book2_columns.index(send_schema_adapter_col) < len(book2_row) else "?"
                send_table = book2_row[book2_columns.index(send_table_adapter_col)] if send_table_adapter_col and book2_columns.index(send_table_adapter_col) < len(book2_row) else "?"
                
                send_sql += f"\nsqlplus {send_userid}/{send_passwd}@{send_db}"
                send_sql += f"\nselect count(*) from {send_schema}.{send_table} where EAI_TRANSFER_DATE > sysdate-(1/12) and EAI_TRANSFER_FLAG='Y';"
                
                self.config_text.insert(tk.END, send_sql)
                
                # 수신 SQL
                recv_sql = "\n[수신SQL]"
                
                recv_userid = book2_row[book2_columns.index(recv_userid_col)] if recv_userid_col and book2_columns.index(recv_userid_col) < len(book2_row) else "?"
                recv_passwd = book2_row[book2_columns.index(recv_passwd_col)] if recv_passwd_col and book2_columns.index(recv_passwd_col) < len(book2_row) else "?"
                recv_db = book2_row[book2_columns.index(recv_db_col)] if recv_db_col and book2_columns.index(recv_db_col) < len(book2_row) else "?"
                recv_schema = book2_row[book2_columns.index(recv_schema_adapter_col)] if recv_schema_adapter_col and book2_columns.index(recv_schema_adapter_col) < len(book2_row) else "?"
                recv_table = book2_row[book2_columns.index(recv_table_adapter_col)] if recv_table_adapter_col and book2_columns.index(recv_table_adapter_col) < len(book2_row) else "?"
                
                recv_sql += f"\nsqlplus {recv_userid}/{recv_passwd}@{recv_db}"
                recv_sql += f"\nselect count(*) from {recv_schema}.{recv_table} where EAI_TRANSFER_DATE > sysdate-(1/12);"
                
                self.config_text.insert(tk.END, recv_sql)
            else:
                # book2에는 없지만 book1에는 있는 경우
                self.config_text.insert(tk.END, "[book2 정보 없음]\n")
            
            # book1 데이터가 있는 경우 해당 정보 표시 (book2 존재 여부와 무관)
            if book1_row:
                # 인터페이스 명 표시 (book1 테이블에서)
                if interface_name_col:
                    interface_name = book1_row[book1_columns.index(interface_name_col)]
                    if interface_name:
                        self.config_text.insert(tk.END, f"\n[인터페이스 명] {interface_name}")
                
                # I_F_Type 표시 (book1 테이블에서)
                if if_type_col:
                    if_type = book1_row[book1_columns.index(if_type_col)]
                    if if_type and str(if_type).strip():
                        self.config_text.insert(tk.END, f"\n[I_F_Type] {if_type}")
                
                # Route정의 표시 (book1 테이블에서)
                if route_col:
                    route_def = book1_row[book1_columns.index(route_col)]
                    if route_def:
                        self.config_text.insert(tk.END, f"\n[Route정의] {route_def}")
            
            conn.close()
            self.log(f"매핑SEQ '{mapping_seq}' 검색 완료")
        
        except Exception as e:
            self.log(f"검색 중 오류 발생: {str(e)}")
            messagebox.showerror("오류", f"검색 중 오류가 발생했습니다: {str(e)}")

def main():
    root = tk.Tk()
    app = ExcelToSqliteApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()