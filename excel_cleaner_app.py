import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pythoncom
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32com.client

queued_files = []

# 엑셀 내부에 주입해 1초 만에 실행할 초고속 VBA 매크로 코드
FAST_VBA_CODE = """
Sub FastCleanWorkbook(targetWb As Workbook)
    Dim nm As Name
    Dim lnk As Variant
    Dim linkArr As Variant
    Dim nmName As String
    Dim shortName As String
    Dim exclPos As Long
    Dim i As Long
    
    On Error Resume Next
    
    ' 1. 외부 통합문서 연결 끊기
    linkArr = targetWb.LinkSources(1)
    If Not IsEmpty(linkArr) Then
        For Each lnk In linkArr
            targetWb.BreakLink Name:=lnk, Type:=1
        Next lnk
    End If
    
    ' 2. 이름 관리자 초고속 일괄 삭제 (기본 시스템 이름 보존)
    For i = targetWb.Names.Count To 1 Step -1
        Set nm = targetWb.Names(i)
        nmName = nm.Name
        exclPos = InStrRev(nmName, "!")
        If exclPos > 0 Then
            shortName = Mid$(nmName, exclPos + 1)
        Else
            shortName = nmName
        End If
        
        If Left$(shortName, 6) <> "_xlnm." And _
           Left$(shortName, 10) <> "Print_Area" And _
           Left$(shortName, 12) <> "Print_Titles" And _
           Left$(shortName, 15) <> "_FilterDatabase" And _
           Left$(shortName, 16) <> "Consolidate_Area" And _
           Left$(shortName, 8) <> "Criteria" And _
           Left$(shortName, 8) <> "Database" And _
           Left$(shortName, 7) <> "Extract" And _
           Left$(shortName, 11) <> "Sheet_Title" Then
            nm.Delete
        End If
    Next i
End Sub
"""


def add_files(file_paths):
  valid_exts = (".xlsx", ".xlsm", ".xls")
  added = 0
  for f in file_paths:
    if f.lower().endswith(valid_exts) and f not in queued_files:
      queued_files.append(f)
      file_listbox.insert(tk.END, f"📄 {os.path.basename(f)}")
      added += 1

  if added > 0:
    status_label.config(
        text=(
            f"총 {len(queued_files)}개 파일 대기 중 ('작업 시작' 버튼을 누르세요)"
        ),
        fg="#217346",
    )
  elif file_paths:
    messagebox.showwarning(
        "파일 확인",
        "추가할 수 있는 새 엑셀 파일(.xlsx, .xlsm, .xls)이 없습니다.",
    )


def clear_list():
  queued_files.clear()
  file_listbox.delete(0, tk.END)
  progress_bar["value"] = 0
  status_label.config(text="대기 목록이 비워졌습니다.", fg="#555555")


def on_drop(event):
  file_paths = root.tk.splitlist(event.data)
  root.after(10, lambda: add_files(file_paths))


def on_add_click():
  file_paths = filedialog.askopenfilenames(
      title="정리할 엑셀 파일을 선택하세요",
      filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls")],
  )
  if file_paths:
    add_files(file_paths)


def set_ui_state(is_running):
  state = tk.DISABLED if is_running else tk.NORMAL
  btn_add.config(state=state)
  btn_clear.config(state=state)
  btn_start.config(state=state)


def start_processing():
  if not queued_files:
    messagebox.showinfo(
        "알림", "먼저 정리할 엑셀 파일을 목록에 추가해 주세요."
    )
    return

  set_ui_state(True)
  progress_bar["maximum"] = len(queued_files)
  progress_bar["value"] = 0

  thread = threading.Thread(target=process_worker, daemon=True)
  thread.start()


def process_worker():
  success_count = 0
  error_files = []
  total = len(queued_files)

  pythoncom.CoInitialize()
  excel = None
  helper_wb = None
  created_new_excel = False

  try:
    root.after(
        0,
        lambda: status_label.config(
            text="⚡ 초고속 매크로 엔진 준비 중...", fg="#D97706"
        ),
    )

    try:
      excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
      excel = win32com.client.DispatchEx("Excel.Application")
      created_new_excel = True

    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    orig_calc = excel.Calculation
    excel.Calculation = -4135  # 수동 계산

    # 매크로 실행용 임시 통합 문서 생성 후 VBA 코드 주입
    helper_wb = excel.Workbooks.Add()
    vb_module = helper_wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    vb_module.CodeModule.AddFromString(FAST_VBA_CODE)

    for idx, file_path in enumerate(queued_files, start=1):
      base_name = os.path.basename(file_path)
      wb = None
      try:
        root.after(
            0,
            lambda i=idx, name=base_name: status_label.config(
                text=f"🚀 초고속 정리 중 ({i}/{total}): {name}", fg="#0052CC"
            ),
        )

        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False)

        # 주입된 VBA 매크로를 엑셀 내부에서 직접 고속 실행
        excel.Run("FastCleanWorkbook", wb)

        dir_name = os.path.dirname(abs_path)
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
        error_files.append(f"{base_name} ({e})")

      root.after(0, lambda v=idx: progress_bar.configure(value=v))

  except Exception as e:
    error_files.append(f"엔진 오류: {e}")

  finally:
    if helper_wb:
      try:
        helper_wb.Close(SaveChanges=False)
      except Exception:
        pass
    if excel:
      try:
        excel.Calculation = orig_calc
        excel.EnableEvents = True
        excel.ScreenUpdating = True
        if created_new_excel:
          excel.Quit()
      except Exception:
        pass
    pythoncom.CoUninitialize()

  def on_complete():
    set_ui_state(False)
    status_label.config(
        text=f"✅ 완료! 총 {success_count}개 파일이 정리되었습니다.",
        fg="#217346",
    )

    result_msg = (
        f"총 {success_count}개의 파일 정리가 완료되었습니다.\n(원본 폴더에"
        " '_정리완료' 파일 생성)"
    )
    if error_files:
      result_msg += "\n\n[실패한 파일]\n" + "\n".join(error_files)

    messagebox.showinfo("작업 완료", result_msg)
    clear_list()

  root.after(0, on_complete)


# --- 메인 윈도우 GUI 구성 ---
root = TkinterDnD.Tk()
root.title("엑셀 초고속 연결/이름 제거기")
root.geometry("460x380")
root.resizable(False, False)
root.attributes("-topmost", True)

list_frame = tk.LabelFrame(
    root,
    text="📂 여기에 엑셀 파일을 끌어다 놓거나 추가하세요",
    font=("맑은 고딕", 10, "bold"),
    fg="#217346",
    padx=10,
    pady=8,
)
list_frame.pack(fill="both", expand=True, padx=15, pady=(12, 5))

scrollbar = tk.Scrollbar(list_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

file_listbox = tk.Listbox(
    list_frame,
    yscrollcommand=scrollbar.set,
    font=("맑은 고딕", 9),
    selectmode=tk.EXTENDED,
    bg="#F9FAFB",
    relief="flat",
)
file_listbox.pack(fill="both", expand=True)
scrollbar.config(command=file_listbox.yview)

list_frame.drop_target_register(DND_FILES)
list_frame.dnd_bind("<<Drop>>", on_drop)
file_listbox.drop_target_register(DND_FILES)
file_listbox.dnd_bind("<<Drop>>", on_drop)

status_label = tk.Label(
    root,
    text="파일을 추가한 뒤 아래 '작업 시작' 버튼을 눌러주세요.",
    font=("맑은 고딕", 9, "bold"),
    fg="#555555",
)
status_label.pack(pady=(5, 2))

progress_bar = ttk.Progressbar(
    root, orient="horizontal", length=430, mode="determinate"
)
progress_bar.pack(pady=(2, 10))

btn_frame = tk.Frame(root)
btn_frame.pack(fill="x", padx=15, pady=(0, 15))

btn_add = tk.Button(
    btn_frame,
    text="+ 파일 추가",
    command=on_add_click,
    font=("맑은 고딕", 9),
    padx=10,
    pady=6,
)
btn_add.pack(side=tk.LEFT, padx=(0, 5))

btn_clear = tk.Button(
    btn_frame,
    text="목록 비우기",
    command=clear_list,
    font=("맑은 고딕", 9),
    padx=10,
    pady=6,
)
btn_clear.pack(side=tk.LEFT)

btn_start = tk.Button(
    btn_frame,
    text="▶ 작업 시작",
    command=start_processing,
    bg="#217346",
    fg="white",
    font=("맑은 고딕", 10, "bold"),
    padx=20,
    pady=6,
)
btn_start.pack(side=tk.RIGHT)

root.mainloop()
