import os
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from tkinterdnd2 import DND_FILES, TkinterDnD


def clean_files(file_paths):
  # 엑셀 파일(.xlsx, .xlsm)만 걸러내기
  valid_files = [
      f for f in file_paths if f.lower().endswith((".xlsx", ".xlsm"))
  ]

  if not valid_files:
    messagebox.showwarning(
        "파일 오류", "엑셀 파일(.xlsx, .xlsm)만 처리할 수 있습니다."
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

  for file_path in valid_files:
    try:
      is_macro = file_path.lower().endswith(".xlsm")
      wb = openpyxl.load_workbook(file_path, keep_vba=is_macro)

      # 1. 이름 관리자 삭제 (기본 시스템 이름은 유지)
      for name in list(wb.defined_names.keys()):
        if not name.startswith(system_prefixes):
          del wb.defined_names[name]

      # 2. 외부 통합문서 연결(External Links) 모두 제거
      wb._external_links.clear()

      # 원본 파일명 뒤에 '_정리완료'를 붙여 저장
      dir_name, base_name = os.path.split(file_path)
      name_only, ext = os.path.splitext(base_name)
      new_path = os.path.join(dir_name, f"{name_only}_정리완료{ext}")

      wb.save(new_path)
      success_count += 1
    except Exception as e:
      error_files.append(f"{os.path.basename(file_path)} ({e})")

  # 결과 알림창 띄우기
  result_msg = f"총 {success_count}개의 파일 정리가 완료되었습니다.\n(원본 폴더에 '_정리완료' 파일 생성)"
  if error_files:
    result_msg += "\n\n[실패한 파일]\n" + "\n".join(error_files)

  messagebox.showinfo("작업 완료", result_msg)


def on_drop(event):
  # 드래그 앤 드롭된 파일 경로 문자열을 리스트로 분리 (공백/중괄호 처리)
  raw_data = event.data
  file_paths = root.tk.splitlist(raw_data)
  clean_files(file_paths)


def on_button_click():
  file_paths = filedialog.askopenfilenames(
      title="정리할 엑셀 파일을 선택하세요",
      filetypes=[("Excel 파일", "*.xlsx *.xlsm")],
  )
  if file_paths:
    clean_files(file_paths)


# 드래그 앤 드롭이 가능한 메인 윈도우 생성
root = TkinterDnD.Tk()
root.title("엑셀 불필요 연결/이름 제거기")
root.geometry("420x250")
root.resizable(False, False)

# 드래그 앤 드롭 영역(프레임) 구성
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
        " 사용자 정의 이름 삭제\n* 인쇄 영역 등 기본 시스템 이름은 유지됨)"
    ),
    font=("맑은 고딕", 9),
    fg="#333333",
    justify="center",
)
drop_label.pack(expand=True, fill="both")

# 드롭 이벤트 연결
drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind("<<Drop>>", on_drop)
drop_label.drop_target_register(DND_FILES)
drop_label.dnd_bind("<<Drop>>", on_drop)

# 기존 파일 선택 버튼도 하단에 유지
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
