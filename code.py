#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chuyển cụm (nội dung footnote) thành footnote Word và xóa ngoặc đơn.
Giữ nguyên dấu chấm câu sau ngoặc (.,;:!?). Không ghi đè file nguồn — luôn lưu file mới.

Yêu cầu môi trường: Windows + Microsoft Word đã cài.
Anaconda: conda install pywin32  hoặc  pip install pywin32

Ví dụ:
  python code.py                           # giao diện
  python code.py "bao_cao.docx"            # → bao_cao_footnote.docx
  python code.py trong.docx -o ten_khac.docx
"""

from __future__ import annotations

import argparse
import re
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


# Nội dung trong (...) phải kết thúc bằng một trong các dấu sau (được giữ lại trong thân văn bản).
PAREN_FOOTNOTE = re.compile(r"\(([^)]*)\)([.,;:!?])")


def default_output_path(input_path: Path) -> Path:
    """Luôn sinh file mới cạnh file gốc: ten_footnote.docx /.doc"""
    suf = input_path.suffix if input_path.suffix.lower() in (".docx", ".doc") else ".docx"
    return input_path.with_name(f"{input_path.stem}_footnote{suf}")


def resolve_output_path(input_path: Path, output_path: Path | None) -> Path:
    """
    Không ghi đè file nguồn: nếu không chỉ định hoặc trùng đường dẫn nguồn thì dùng default_output_path.
    """
    if output_path is None:
        return default_output_path(input_path)
    try:
        if output_path.resolve() == input_path.resolve():
            return default_output_path(input_path)
    except OSError:
        pass
    return output_path


def _wd_save_format(win32, dest: Path) -> int:
    """
    WdSaveFormat cho SaveAs2 — bắt buộc khi đuôi file đích khác định dạng tài liệu đang mở
    (ví dụ mở .doc nhưng lưu .docx → lỗi Incompatible file type and file extension).
    """
    ext = dest.suffix.lower()
    lit = {
        ".doc": 0,
        ".docx": 12,
        ".docm": 13,
    }.get(ext, 12)
    name_by_ext = {
        ".doc": "wdFormatDocument",
        ".docx": "wdFormatXMLDocument",
        ".docm": "wdFormatXMLDocumentMacroEnabled",
    }
    nm = name_by_ext.get(ext, "wdFormatXMLDocument")
    c = getattr(win32, "constants", None)
    if c is None:
        return lit
    return int(getattr(c, nm, lit))


def _ensure_word():
    try:
        import win32com.client as win32  # type: ignore
    except ImportError as e:
        raise ImportError(
            "Thiếu pywin32. Cài: conda install pywin32  hoặc  pip install pywin32"
        ) from e
    return win32


def parenthetical_ranges_to_footnotes(doc, matches_desc: list[re.Match[str]]) -> int:
    """
    matches_desc: các match theo thứ tự m.start() giảm dần để chỉnh sửa không làm lệch offset.
    Trả về số footnote đã chèn.
    """
    done = 0
    punct_len = lambda m: len(m.group(2))

    for m in matches_desc:
        inner = (m.group(1) or "").strip()
        if not inner:
            continue

        py_start = m.start()
        py_end_exclusive = m.end() - punct_len(m)

        word_start = py_start + 1
        word_end = py_end_exclusive + 1

        # Chèn dấu footnote tại điểm cuối ngoặc (trước dấu câu),
        # sau đó xóa toàn bộ "(...)" khỏi thân bài.
        delete_rng = doc.Range(word_start, word_end)
        if not delete_rng.Text or delete_rng.Text.strip() == "":
            continue

        insert_rng = doc.Range(word_end, word_end)
        fn = doc.Footnotes.Add(Range=insert_rng)
        fn.Range.Text = inner
        delete_rng.Text = ""

        done += 1

    return done


def convert_by_paragraph_scan(doc) -> int:
    """
    Quét từng Paragraph, căn chỉnh Range qua Paragraph.Range.Start + offset.
    Không dùng WordFind wildcard — tránh lỗi "Pattern Match expression which is not valid".
    """
    inserted = 0
    pc = int(doc.Paragraphs.Count)

    # Duyệt đoạn từ cuối lên: an toàn hơn khi chỉnh sửa tài liệu.
    for i in range(pc, 0, -1):
        para = doc.Paragraphs(i)
        pr = para.Range
        txt = pr.Text or ""
        if "(" not in txt:
            continue

        ms = list(PAREN_FOOTNOTE.finditer(txt))
        if not ms:
            continue

        base = int(pr.Start)
        for m in sorted(ms, key=lambda x: x.start(), reverse=True):
            inner = (m.group(1) or "").strip()
            if not inner:
                continue
            lp = len(m.group(2))
            w_s = base + m.start()
            w_e = base + m.end() - lp
            delete_rng = doc.Range(w_s, w_e)  # "(...)" cần xóa
            if not (delete_rng.Text or "").strip():
                continue
            insert_rng = doc.Range(w_e, w_e)  # chèn ngay sau ")" (trước dấu câu)
            fn = doc.Footnotes.Add(Range=insert_rng)
            fn.Range.Text = inner
            delete_rng.Text = ""
            inserted += 1

    return inserted


def convert_doc(input_path: Path, output_path: Path | None) -> int:
    win32 = _ensure_word()
    wd_do_not_save = getattr(win32.constants, "wdDoNotSaveChanges", 0)

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    out = resolve_output_path(input_path, output_path)
    abs_in = str(input_path.resolve())
    abs_out = str(out.resolve())

    doc = None
    try:
        doc = word.Documents.Open(abs_in)

        inserted = convert_by_paragraph_scan(doc)
        if inserted == 0:
            raw = doc.Content.Text
            all_matches = list(PAREN_FOOTNOTE.finditer(raw))
            ordered = sorted(all_matches, key=lambda x: x.start(), reverse=True)
            inserted = parenthetical_ranges_to_footnotes(doc, ordered)

        filefmt = _wd_save_format(win32, out)
        doc.SaveAs2(abs_out, filefmt)

        return inserted
    finally:
        if doc is not None:
            doc.Close(wd_do_not_save)
        word.Quit(wd_do_not_save)


def run_gui() -> None:
    root = tk.Tk()
    root.title("Footnote tự động — Word")
    root.minsize(560, 280)
    root.geometry("620x320")

    in_var = tk.StringVar()
    out_var = tk.StringVar()

    pad = {"padx": 10, "pady": 6}

    frm = ttk.Frame(root, padding=12)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="File Word:").grid(row=0, column=0, sticky="w", **pad)
    ent_in = ttk.Entry(frm, textvariable=in_var, width=60)
    ent_in.grid(row=0, column=1, sticky="ew", **pad)

    def browse_in() -> None:
        p = filedialog.askopenfilename(
            title="Chọn file Word",
            filetypes=[
                ("Word", "*.docx *.doc"),
                ("Tất cả", "*.*"),
            ],
        )
        if p:
            in_var.set(p)
            out_var.set(str(default_output_path(Path(p))))

    ttk.Button(frm, text="Duyệt…", command=browse_in).grid(
        row=0, column=2, sticky="e", **pad
    )

    hint = ttk.Label(
        frm,
        text="Luôn tạo file mới (mặc định *_footnote.docx trong cùng thư mục).",
        font=("Segoe UI", 8),
    )
    hint.grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 4))

    ttk.Label(frm, text="File lưu:").grid(row=2, column=0, sticky="w", **pad)
    ent_out = ttk.Entry(frm, textvariable=out_var, width=60)
    ent_out.grid(row=2, column=1, sticky="ew", **pad)

    def browse_out() -> None:
        start = in_var.get().strip() or None
        p = filedialog.asksaveasfilename(
            title="Lưu file Word",
            defaultextension=".docx",
            filetypes=[
                ("Word (.docx)", "*.docx"),
                ("Word (.doc)", "*.doc"),
                ("Tất cả", "*.*"),
            ],
            initialfile=Path(start).stem + "_footnote.docx" if start else "",
        )
        if p:
            out_var.set(p)

    btn_out = ttk.Button(frm, text="Duyệt…", command=browse_out)
    btn_out.grid(row=2, column=2, sticky="e", **pad)

    log = tk.Text(frm, height=8, wrap="word", state="disabled", font=("Segoe UI", 9))
    log.grid(row=3, column=0, columnspan=3, sticky="nsew", padx=10, pady=8)
    frm.rowconfigure(3, weight=1)
    sb = ttk.Scrollbar(frm, orient="vertical", command=log.yview)
    sb.grid(row=3, column=3, sticky="ns", pady=8)
    log.configure(yscrollcommand=sb.set)

    def set_log(text: str) -> None:
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.insert("1.0", text.rstrip() + ("\n" if text and not text.endswith("\n") else ""))
        log.configure(state="disabled")

    btn_run = ttk.Button(frm, text="Xử lý", width=18)

    def set_busy(on: bool) -> None:
        btn_run.configure(state="disabled" if on else "normal")
        ent_in.configure(state="disabled" if on else "normal")
        ent_out.configure(state="disabled" if on else "normal")
        btn_out.configure(state="disabled" if on else "normal")

    def finish_ok(n: int, out_path: str) -> None:
        set_busy(False)
        msg = f"Đã chèn {n} footnote.\nLưu tại:\n{out_path}"
        set_log(msg)
        messagebox.showinfo("Xong", f"Đã chèn {n} footnote.\n{out_path}")

    def finish_err(err: str) -> None:
        set_busy(False)
        set_log(f"Lỗi:\n{err}")
        messagebox.showerror("Lỗi", err)

    def do_convert() -> None:
        inp_s = in_var.get().strip()
        if not inp_s:
            messagebox.showwarning("Thiếu file", "Hãy chọn file Word.")
            return
        inp = Path(inp_s)
        if not inp.is_file():
            messagebox.showerror("Lỗi", f"Không tìm thấy file:\n{inp}")
            return

        o = out_var.get().strip()
        out = Path(o) if o else None

        set_busy(True)
        set_log("Đang xử lý (Word có thể mở nền, vài giây tới vài phút)…")

        def worker() -> None:
            import pythoncom

            pythoncom.CoInitialize()
            try:
                n = convert_doc(inp, out)
                outp = str(resolve_output_path(inp, out).resolve())
                root.after(0, lambda nn=n, op=outp: finish_ok(nn, op))
            except Exception as e:
                err_txt = str(e)
                root.after(0, lambda t=err_txt: finish_err(t))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=worker, daemon=True).start()

    btn_run.configure(command=do_convert)
    btn_run.grid(row=4, column=1, pady=12, sticky="w")

    root.mainloop()


def main() -> None:
    p = argparse.ArgumentParser(
        description="Chuyen (noi dung) thanh footnote Word, giu dau cham cau sau ngoac."
    )
    p.add_argument(
        "input_docx",
        nargs="?",
        type=Path,
        default=None,
        help="Duong dan file .doc/.docx (bo qua de mo giao dien)",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="File ghi ra (mac dinh: ten_footnote.docx cung thu muc, khong ghi de nguon)",
    )
    args = p.parse_args()

    if args.input_docx is None:
        run_gui()
        return

    inp = args.input_docx
    if not inp.is_file():
        sys.stderr.write(f"Khong tim thay file: {inp}\n")
        sys.exit(1)

    try:
        _ensure_word()
    except ImportError as e:
        sys.stderr.write(f"{e}\n")
        sys.exit(1)

    try:
        n = convert_doc(inp, args.output)
    except Exception as e:
        sys.stderr.write(f"Loi: {e}\n")
        sys.exit(1)

    out_p = resolve_output_path(inp, args.output)
    print(f"Da chen {n} footnote.")
    print(f"Luu tai: {out_p.resolve()}")


if __name__ == "__main__":
    main()
