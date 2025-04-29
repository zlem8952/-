import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# LibreOffice 실행파일 경로 (환경에 맞게 수정)
SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
# SOFFICE_PATH = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"  # 32비트 설치시

class ExcelToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 폴더 → PDF 일괄 변환기")
        self.root.geometry("700x450")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", font=('맑은 고딕', 11), padding=7)
        style.configure("TLabel", font=('맑은 고딕', 10))
        style.configure("TEntry", font=('맑은 고딕', 10))

        frm = ttk.Frame(root, padding=25)
        frm.pack(expand=True, fill='both')

        ttk.Label(frm, text="엑셀 파일이 들어있는 폴더 선택:").grid(row=0, column=0, sticky='w')
        self.folder_entry = ttk.Entry(frm, width=55, state='readonly')
        self.folder_entry.grid(row=0, column=1, padx=5)
        ttk.Button(frm, text="폴더 선택", command=self.select_folder).grid(row=0, column=2, padx=5)

        self.convert_btn = ttk.Button(frm, text="PDF로 일괄 변환 시작", command=self.convert_files)
        self.convert_btn.grid(row=1, column=0, columnspan=3, pady=18)

        self.progress = ttk.Progressbar(frm, orient='horizontal', length=620, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=3, pady=10)
        self.progress.grid_remove()

        ttk.Label(frm, text="변환 결과 로그:").grid(row=3, column=0, columnspan=3, sticky='w')
        self.log = tk.Text(frm, height=13, width=90, font=('맑은 고딕', 9))
        self.log.grid(row=4, column=0, columnspan=3, pady=5)
        self.log.config(state='disabled')

        self.folder = ""

    def select_folder(self):
        folder = filedialog.askdirectory(title="엑셀 파일이 들어있는 폴더 선택")
        if folder:
            self.folder = folder
            self.folder_entry.config(state='normal')
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            self.folder_entry.config(state='readonly')

    def find_excel_files(self, folder):
        # .xlsx만 변환 (가장 안정적)
        excel_exts = ('.xlsx',)
        file_list = []
        for dirpath, _, files in os.walk(folder):
            for file in files:
                if file.lower().endswith(excel_exts):
                    file_list.append(os.path.join(dirpath, file))
        return file_list

    def convert_files(self):
        if not self.folder or not os.path.exists(self.folder):
            messagebox.showerror("오류", "엑셀 파일이 들어있는 폴더를 선택하세요.")
            return

        excel_files = self.find_excel_files(self.folder)
        if not excel_files:
            messagebox.showinfo("결과", "폴더 내에 변환할 .xlsx 파일이 없습니다.")
            return

        self.progress.grid()
        self.progress['value'] = 0
        self.log.config(state='normal')
        self.log.delete(1.0, tk.END)
        self.root.update_idletasks()

        total = len(excel_files)
        success, failed = 0, 0
        for idx, file_path in enumerate(excel_files):
            try:
                # 파일 경로/이름에 공백, 한글, 특수문자 있을 때 따옴표로 감싸기
                outdir = os.path.dirname(file_path)
                cmd = [
                    SOFFICE_PATH,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', outdir,
                    file_path
                ]
                # Windows에서 경로에 공백/한글 있을 때는 shell=True로 실행
                subprocess.run(' '.join(f'"{c}"' if ' ' in c or '(' in c or ')' in c or '한글' in c else c for c in cmd),
                               shell=True, check=True)
                # 변환 후 PDF가 실제로 생성됐는지 확인
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                if os.path.exists(pdf_path):
                    self.log.insert(tk.END, f"성공: {file_path}\n")
                    success += 1
                else:
                    self.log.insert(tk.END, f"실패(파일 생성 안됨): {file_path}\n")
                    failed += 1
            except Exception as e:
                self.log.insert(tk.END, f"실패: {file_path} ({e})\n")
                failed += 1
            self.progress['value'] = (idx+1)/total*100
            self.root.update_idletasks()

        self.progress.grid_remove()
        self.log.config(state='disabled')
        messagebox.showinfo("완료", f"총 {total}개 중 {success}개 변환 성공, {failed}개 실패\n"
                                   "※ 파일이 누락될 경우, 파일을 닫고 경로/이름을 영문으로 바꿔 재시도하세요.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToPDFApp(root)
    root.mainloop()
