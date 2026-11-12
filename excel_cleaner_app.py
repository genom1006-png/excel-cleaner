import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pythoncom
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32com.client


def clean_files(file_paths):
  valid_files = [
      f for f in file_paths if f.lower().endswith((".xlsx", ".xlsm", ".xls"))
  ]

  if not valid_files:
    messagebox.showwarning(
        "파일 오류", "엑셀 파일(.xlsx, .xlsm, .xls)만 처리할 수 있습니다."
    )
    return

  success_count = 0
  error_files = []

  # 보존할 엑셀 기본 시스템 이름 목록
  system_prefixes = (
      "_xlnm.",
      "Print_Area",
      "Print_Titles",
      "_FilterDatabase",
      "Consolidate_Area",
      "Criteria",
      "Database",
      "Extract",
      "Sheet_Title",
  )

  pythoncom.CoInitialize()
  excel = None

  try:
    # 백그라운드에서 엑셀 실행 (화면 깜빡임 및 팝업 차단)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    for file_path in valid_files:
      wb = None
      try:
        # 절대 경로로 변환 (COM 객체 필수)
        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False)

        # 1. 이름 관리자 삭제 (기본 시스템 이름은 유지)
        # 뒤에서부터 역순으로 지워야 인덱스 오류 방지
        for i in range(wb.Names.Count, 0, -1):
          try:
            nm = wb.Names.Item(i)
            # 시트명이 포함된 경우(예: 'Sheet1!Print_Area')를 대비해 느낌표 뒷부분 추출
            short_name = nm.Name.split("!")[-1]
            if not short_name.startswith(system_prefixes):
              nm.Delete()
          except Exception:
            pass

        # 2. 외부 통합문서 연결(External Links) 모두 끊기 (xlExcelLinks = 1)
        try:
          links = wb.LinkSources(1)
          if links:
            for lnk in links:
              wb.BreakLink(Name=lnk, Type=1)
        except Exception:
          pass

        # 원본 파일명 뒤에 '_정리완료'를 붙여 저장
        dir_name, base_name = os.path.split(abs_path)
        name_only, ext = os.path.splitext(base_name)
        new_path = os.path.join(dir_name, f"{name_only}_정리완료{ext}")

        wb.SaveAs(new_path)
        wb.Close(SaveChanges=False)
        success_count += 1

      except Exception as e:
        if wb:
          try:
            wb.Close(SaveChanges=False)
          except Exception:
            pass
        error_files.append(f"{os.path.basename(file_path)} ({e})")

  finally:
    if excel:
      excel.Quit()
    pythoncom.CoUninitialize()

  result_msg = f"총 {success_count}개의 파일 정리가 완료되었습니다.\n(원본 폴더에 '_정리완료' 파일 생성)"
  if error_files:
    result_msg += "\n\n[실패한 파일]\n" + "\n".join(error_files)

  messagebox.showinfo("작업 완료", result_msg)


def on_drop(event):
  raw_data = event.data
  file_paths = root.tk.splitlist(raw_data)
  clean_files(file_paths)


def on_button_click():
  file_paths = filedialog.askopenfilenames(
      title="정리할 엑셀 파일을 선택하세요",
      filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls")],
  )
  if file_paths:
    clean_files(file_paths)


root = TkinterDnD.Tk()
root.title("엑셀 불필요 연결/이름 제거기 (DRM 대응)")
root.geometry("420x250")
root.resizable(False, False)

drop_frame = tk.LabelFrame(
    root,
    text="여기에 엑셀 파일을 끌어다 놓으세요 (여러 개 가능)",
    font=("맑은 고딕", 10, "bold"),
    fg="#217346",
    padx=10,
    pady=10,
)
drop_frame.pack(fill="both", expand=True, padx=20, pady=(15, 10))

drop_label = tk.Label(
    drop_frame,
    text=(
        "📂 파일을 드래그 앤 드롭하면 즉시 정리됩니다.\n\n(외부 연결 제거 및"
        " 사용자 정의 이름 삭제\n* 사내 문서보안(DRM) 및 구버전(.xls) 파일 완벽"
        " 지원)"
    ),
    font=("맑은 고딕", 9),
    fg="#333333",
    justify="center",
)
drop_label.pack(expand=True, fill="both")

drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind("<<Drop>>", on_drop)
drop_label.drop_target_register(DND_FILES)
drop_label.dnd_bind("<<Drop>>", on_drop)

btn = tk.Button(
    root,
    text="또는 클릭하여 파일 직접 선택하기",
    command=on_button_click,
    bg="#217346",
    fg="white",
    font=("맑은 고딕", 10, "bold"),
    padx=15,
    pady=6,
)
btn.pack(pady=(0, 15))

root.mainloop()
