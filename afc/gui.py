"""금융기관조회서 변환기 — 비전문가용 데스크톱 GUI (tkinter).

회계사가 파이썬/명령어 없이 쓸 수 있도록: 조회서 ZIP 선택 → [변환 시작] → 결과 폴더 열기.
PyInstaller로 단일 .exe 패키징 가능(build_exe.bat 참조).

실행: python -m afc.gui
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from afc.run import run_pipeline

APP_TITLE = "금융기관조회서 변환기"
DEFAULT_OUT = Path.home() / "Desktop" / "금융조회_결과"


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.zip_path = tk.StringVar()
        self.pbc_path = tk.StringVar()
        self.out_dir = tk.StringVar(value=str(DEFAULT_OUT))
        self.msg_q: queue.Queue = queue.Queue()
        self._build()
        self.root.after(100, self._poll)

    # ── UI ──────────────────────────────────────────────────────────────
    def _build(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry("720x560")
        self.root.minsize(640, 480)
        pad = {"padx": 10, "pady": 6}

        ttk.Label(self.root, text=APP_TITLE, font=("맑은 고딕", 16, "bold")).pack(anchor="w", **pad)
        ttk.Label(
            self.root,
            text="은행에서 받은 '금융기관조회서(전자) ZIP' 을 선택하고 [변환 시작]을 누르세요.\n"
                 "결과 엑셀이 아래 저장 폴더에 만들어집니다.",
            foreground="#444",
        ).pack(anchor="w", padx=10)

        frm = ttk.Frame(self.root)
        frm.pack(fill="x", **pad)
        frm.columnconfigure(1, weight=1)

        self._file_row(frm, 0, "① 조회서 ZIP 파일", self.zip_path, self._pick_zip, "필수")
        self._file_row(frm, 1, "② 회사 예금명세 (선택)", self.pbc_path, self._pick_pbc, "선택")
        self._file_row(frm, 2, "③ 결과 저장 폴더", self.out_dir, self._pick_out, "폴더")

        self.run_btn = ttk.Button(self.root, text="변환 시작 ▶", command=self._start)
        self.run_btn.pack(anchor="w", padx=10, pady=8)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        ttk.Label(self.root, text="진행 상황", foreground="#444").pack(anchor="w", padx=10, pady=(8, 0))
        self.log = scrolledtext.ScrolledText(self.root, height=14, state="disabled",
                                             font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _file_row(self, frm, row, label, var, cmd, tag) -> None:
        ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(frm, textvariable=var).grid(row=row, column=1, sticky="ew", padx=4)
        ttk.Button(frm, text="선택…", command=cmd, width=8).grid(row=row, column=2, padx=4)

    # ── 파일 선택 ────────────────────────────────────────────────────────
    def _pick_zip(self) -> None:
        p = filedialog.askopenfilename(title="조회서 ZIP 선택",
                                       filetypes=[("ZIP 파일", "*.zip"), ("모든 파일", "*.*")])
        if p:
            self.zip_path.set(p)

    def _pick_pbc(self) -> None:
        p = filedialog.askopenfilename(title="회사 예금명세 엑셀 선택 (선택)",
                                       filetypes=[("엑셀 파일", "*.xlsx"), ("모든 파일", "*.*")])
        if p:
            self.pbc_path.set(p)

    def _pick_out(self) -> None:
        p = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if p:
            self.out_dir.set(p)

    # ── 실행 ────────────────────────────────────────────────────────────
    def _start(self) -> None:
        zip_path = self.zip_path.get().strip()
        if not zip_path or not Path(zip_path).exists():
            messagebox.showwarning(APP_TITLE, "조회서 ZIP 파일을 먼저 선택하세요.")
            return
        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self._clear_log()
        pbc = self.pbc_path.get().strip() or None
        out_base = Path(self.out_dir.get().strip() or DEFAULT_OUT)
        threading.Thread(
            target=self._worker, args=(Path(zip_path), pbc, out_base), daemon=True
        ).start()

    def _worker(self, zip_path: Path, pbc, out_base: Path) -> None:
        try:
            out_dir, written = run_pipeline(
                zip_path, pbc=Path(pbc) if pbc else None, out_base=out_base,
                log=lambda m: self.msg_q.put(("log", m)),
            )
            self.msg_q.put(("done", (out_dir, len(written))))
        except Exception:  # noqa: BLE001 — 사용자에게 친절히 표시
            self.msg_q.put(("error", traceback.format_exc()))

    # ── 메시지 펌프 (스레드 → UI) ─────────────────────────────────────────
    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "log":
                    self._append(payload)
                elif kind == "done":
                    out_dir, n = payload
                    self._finish_ok(out_dir, n)
                elif kind == "error":
                    self._finish_err(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _finish_ok(self, out_dir: Path, n: int) -> None:
        self.progress.stop()
        self.run_btn.config(state="normal")
        if messagebox.askyesno(APP_TITLE, f"변환 완료! 파일 {n}개가 만들어졌습니다.\n\n결과 폴더를 열까요?"):
            try:
                os.startfile(str(out_dir))  # Windows
            except Exception:
                messagebox.showinfo(APP_TITLE, f"결과 위치:\n{out_dir}")

    def _finish_err(self, detail: str) -> None:
        self.progress.stop()
        self.run_btn.config(state="normal")
        self._append("\n[오류]\n" + detail)
        messagebox.showerror(
            APP_TITLE,
            "변환 중 문제가 발생했습니다.\n\n"
            "· 올바른 '금융기관조회서(전자) ZIP'인지 확인하세요.\n"
            "· 파일이 다른 프로그램(엑셀 등)에서 열려 있으면 닫고 다시 시도하세요.\n\n"
            "자세한 내용은 아래 '진행 상황'에 표시했습니다.",
        )

    # ── 로그 ────────────────────────────────────────────────────────────
    def _append(self, text: str) -> None:
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self) -> None:
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


def main() -> None:
    # 명령행 모드: zip 경로를 인자로 주면 GUI 없이 변환 (스크립트·검증용).
    #   금융기관조회서변환기.exe "조회서.zip" ["회사명세.xlsx"]
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        try:  # 콘솔 인코딩(cp949)로 한글 로그가 깨지지 않게
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
        pbc = Path(args[1]) if len(args) > 1 else None
        out_dir, written = run_pipeline(
            Path(args[0]), pbc=pbc, out_base=Path.home() / "Desktop" / "금융조회_결과", log=print
        )
        print(f"파일 {len(written)}개 생성: {out_dir}")
        return

    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")  # Windows 기본 테마
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
