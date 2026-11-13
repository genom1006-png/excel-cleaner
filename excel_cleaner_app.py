import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pythoncom
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32com.client

queued_files = []


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
  root.update()


def update_status(msg, color="#0052CC"):
  status_label.config(text=msg, fg=color)
  root.update()


def start_processing():
  if not queued_files:
    messagebox.showinfo(
        "알림", "먼저 정리할 엑셀 파일을 목록에 추가해 주세요."
    )
    return

  set_ui_state(True)
  total = len(queued_files)
  progress_bar["maximum"] = total
  progress_bar["value"] = 0
  root.update()

  success_count = 0
  error_files = []

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
  created_new_excel = False

  try:
    update_status("⚡ 엑셀 연결 중...", "#D97706")
    try:
      excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
      excel = win32com.client.DispatchEx("Excel.Application")
      created_new_excel = True

    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3

    for idx, file_path in enumerate(queued_files, start=1):
      base_name = os.path.basename(file_path)
      wb = None
      try:
        update_status(f"[{idx}/{total}] 1/4 파일 여는 중: {base_name}")
        abs_path = os.path.abspath(file_path)

        # UpdateLinks=0(업데이트 안함), ReadOnly=False, CorruptLoad=0
        wb = excel.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
        )

        # 열리자마자 자동 계산 끄기
        orig_calc = excel.Calculation
        excel.Calculation = -4135  # 수동 계산
        excel.ScreenUpdating = False

        try:
          # 2단계: 외부 연결 끊기
          update_status(f"[{idx}/{total}] 2/4 외부 연결 제거 중: {base_name}")
          links = wb.LinkSources(1)
          if links:
            for lnk in links:
              try:
                wb.BreakLink(Name=lnk, Type=1)
              except Exception:
                pass

          # 3단계: 이름 관리자 삭제
          total_names = wb.Names.Count
          update_status(
              f"[{idx}/{total}] 3/4 이름 정리 중 (총 {total_names}개)..."
          )

          for i in range(total_names, 0, -1):
            if i % 200 == 0:
              update_status(
                  f"[{idx}/{total}] 3/4 이름 정리 중... ({total_names - i} /"
                  f" {total_names})"
              )
            try:
              nm = wb.Names.Item(i)
              short_name = nm.Name.split("!")[-1]
              if not short_name.startswith(system_prefixes):
                nm.Delete()
            except Exception:
              pass

        finally:
          excel.Calculation = orig_calc
          excel.ScreenUpdating = True

        # 4단계: 저장
        update_status(f"[{idx}/{total}] 4/4 저장 중: {base_name}", "#217346")
        dir_name = os.path.dirname(abs_path)
        name_only, ext = os.path.splitext(base_name)
        new_path = os.path.join(dir_name, f"{name_only}_정리완료{ext}")

        wb.SaveAs(new_path)
        wb.Close(SaveChanges=False)
        success_count += 1

      except Exception as e:
        excel.ScreenUpdating = True
        if wb:
          try:
            wb.Close(SaveChanges=False)
          except Exception:
            pass
        error_files.append(f"{base_name} ({e})")

      progress_bar["value"] = idx
      root.update()

  finally:
    if excel and created_new_excel:
      try:
        excel.Quit()
      except Exception:
        pass
    pythoncom.CoUninitialize()

  set_ui_state(False)
  update_status(f"✅ 완료! 총 {success_count}개 파일 정리됨", "#217346")

  result_msg = f"총 {success_count}개의 파일 정리가 완료되었습니다."
  if error_files:
    result_msg += "\n\n[실패한 파일]\n" + "\n".join(error_files)

  messagebox.showinfo("작업 완료", result_msg)
  clear_list()


# --- 메인 윈도우 GUI 구성 ---
root = TkinterDnD.Tk()
root.title("엑셀 연결/이름 제거기")
root.geometry("460x380")
root.resizable(False, False)

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
