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


# Nội dung trong (...) rồi đến dấu câu (được giữ lại trong thân văn bản).
# Cho phép khoảng trắng giữa ")" và dấu câu; nhận cả dấu câu fullwidth.
PAREN_FOOTNOTE = re.compile(r"\(([^)]*)\)(\s*)([.,;:!?。；：！？])")

# Ký tự cấu trúc Word: không được xóa (ô bảng, ngắt trang, dấu footnote, …).
_WD_STRUCT = frozenset("\x01\x02\x03\x04\x05\x07\x08\x0b\x0c")


def default_output_path(input_path: Path) -> Path:
    """Luôn sinh file mới cạnh file gốc (ưu tiên .docx — ổn định hơn .doc)."""
    return input_path.with_name(f"{input_path.stem}_footnote.docx")


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


def _clean_fn_text(s: str) -> str:
    for ch in _WD_STRUCT:
        s = s.replace(ch, "")
    return s.replace("\r", " ").replace("\n", " ").strip()


def _prepare_doc(doc) -> None:
    """Gỡ khóa / tắt track changes để Word cho phép xóa và chèn footnote."""
    try:
        wd_none = -1  # wdNoProtection
        if int(doc.ProtectionType) != wd_none:
            try:
                doc.Unprotect()
            except Exception:
                pass
    except Exception:
        pass
    try:
        doc.TrackRevisions = False
    except Exception:
        pass


def _try_delete_span(doc, start: int, end: int) -> bool:
    """
    Xóa [start, end) trên truyện chính. Range COM không được giữ qua thao tác khác.
    Trả về False nếu không xóa được gì (ô bảng / nội dung khóa).
    """
    if end <= start:
        return False
    start, end = int(start), int(end)

    rng = doc.Range(start, end)
    raw = rng.Text or ""
    if not raw:
        return False
    if any(ch in _WD_STRUCT for ch in raw):
        return _delete_chars_skip_struct(doc, start, end)

    try:
        rng.Delete()
        return True
    except Exception:
        pass
    try:
        rng = doc.Range(start, end)
        rng.Text = ""
        return True
    except Exception:
        return _delete_chars_skip_struct(doc, start, end)


def _delete_chars_skip_struct(doc, start: int, end: int) -> bool:
    n_ok = 0
    for pos in range(end - 1, start - 1, -1):
        cr = doc.Range(pos, pos + 1)
        t = cr.Text or ""
        if not t or t[0] in _WD_STRUCT or t[0] == "\r":
            continue
        try:
            cr.Delete()
            n_ok += 1
        except Exception:
            try:
                cr.Text = ""
                n_ok += 1
            except Exception:
                continue
    return n_ok > 0


def _replace_parens_with_footnote(doc, w_s: int, w_e: int, inner: str) -> bool:
    """
    Xóa "(...)" trước, rồi chèn footnote tại chỗ đó.
    Không giữ Range COM qua Footnotes.Add — tránh lỗi 'The range cannot be deleted'
    (Range cũ bị phình ra gồm cả dấu footnote, Word từ chối xóa).
    """
    inner = _clean_fn_text(inner)
    if not inner:
        return False

    probe = doc.Range(int(w_s), int(w_e))
    raw = probe.Text or ""
    if "(" not in raw or ")" not in raw:
        return False

    if not _try_delete_span(doc, w_s, w_e):
        return False

    insert_rng = doc.Range(int(w_s), int(w_s))
    fn = doc.Footnotes.Add(Range=insert_rng)
    fn.Range.Text = inner
    return True


def parenthetical_ranges_to_footnotes(
    doc, matches_desc: list[re.Match[str]], story_start: int = 0
) -> tuple[int, int]:
    """
    matches_desc: match theo m.start() giảm dần.
    story_start: doc.Content.Start (thường 0). Offset Python phải khớp ký tự Word
    (đã chuẩn hóa \\r\\n → \\r).
    Trả về (số footnote đã chèn, số vị trí bỏ qua).
    """
    done = 0
    skipped = 0

    for m in matches_desc:
        inner = (m.group(1) or "").strip()
        if not inner:
            continue

        # group(2)=khoảng trắng sau ")", group(3)=dấu câu — không xóa.
        keep = len(m.group(2) or "") + len(m.group(3) or "")
        w_s = int(story_start) + m.start()
        w_e = int(story_start) + m.end() - keep

        try:
            if _replace_parens_with_footnote(doc, w_s, w_e, inner):
                done += 1
            else:
                skipped += 1
        except Exception:
            skipped += 1

    return done, skipped


def convert_by_paragraph_scan(doc) -> tuple[int, int]:
    """
    Quét từng Paragraph, căn Range qua Paragraph.Range.Start + offset.
    Không dùng Word Find wildcard — tránh lỗi pattern.
    Trả về (đã chèn, bỏ qua).
    """
    inserted = 0
    skipped = 0
    pc = int(doc.Paragraphs.Count)

    for i in range(pc, 0, -1):
        para = doc.Paragraphs(i)
        pr = para.Range
        txt = (pr.Text or "").replace("\r\n", "\r")
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
            keep = len(m.group(2) or "") + len(m.group(3) or "")
            w_s = base + m.start()
            w_e = base + m.end() - keep
            try:
                if _replace_parens_with_footnote(doc, w_s, w_e, inner):
                    inserted += 1
                else:
                    skipped += 1
            except Exception:
                skipped += 1

    return inserted, skipped


def convert_doc(input_path: Path, output_path: Path | None) -> tuple[int, int]:
    win32 = _ensure_word()
    wd_do_not_save = getattr(win32.constants, "wdDoNotSaveChanges", 0)

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = False
    except Exception:
        pass

    out = resolve_output_path(input_path, output_path)
    abs_in = str(input_path.resolve())
    abs_out = str(out.resolve())
    if abs_out.lower() == abs_in.lower():
        out = default_output_path(input_path)
        abs_out = str(out.resolve())

    doc = None
    try:
        doc = word.Documents.Open(abs_in, ReadOnly=False, AddToRecentFiles=False)
        _prepare_doc(doc)

        inserted, skipped = convert_by_paragraph_scan(doc)
        if inserted == 0 and skipped == 0:
            raw = (doc.Content.Text or "").replace("\r\n", "\r")
            all_matches = list(PAREN_FOOTNOTE.finditer(raw))
            ordered = sorted(all_matches, key=lambda x: x.start(), reverse=True)
            story0 = int(doc.Content.Start)
            inserted, skipped = parenthetical_ranges_to_footnotes(doc, ordered, story0)

        filefmt = _wd_save_format(win32, out)
        doc.SaveAs2(abs_out, filefmt)

        return inserted, skipped
    finally:
        try:
            word.ScreenUpdating = True
        except Exception:
            pass
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

    def finish_ok(n: int, skipped: int, out_path: str) -> None:
        set_busy(False)
        extra = ""
        if skipped:
            extra = (
                f"\nBỏ qua {skipped} vị trí (bảng/ô khóa hoặc Word không cho xóa)."
            )
        msg = f"Đã chèn {n} footnote.{extra}\nLưu tại:\n{out_path}"
        set_log(msg)
        messagebox.showinfo("Xong", f"Đã chèn {n} footnote.{extra}\n{out_path}")

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
                n, skipped = convert_doc(inp, out)
                outp = str(resolve_output_path(inp, out).resolve())
                root.after(
                    0,
                    lambda nn=n, sk=skipped, op=outp: finish_ok(nn, sk, op),
                )
            except Exception as e:
                err_txt = str(e)
                if "range cannot be deleted" in err_txt.lower():
                    err_txt = (
                        "Word không xóa được đoạn chữ (thường do bảng, ô, "
                        "hoặc file đang được Word khác mở).\n"
                        "Hãy đóng file trong Word rồi thử lại; hoặc lưu thành .docx."
                    )
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
        n, skipped = convert_doc(inp, args.output)
    except Exception as e:
        sys.stderr.write(f"Loi: {e}\n")
        sys.exit(1)

    out_p = resolve_output_path(inp, args.output)
    print(f"Da chen {n} footnote.")
    if skipped:
        print(f"Bo qua {skipped} vi tri.")
    print(f"Luu tai: {out_p.resolve()}")


if __name__ == "__main__":
    main()
