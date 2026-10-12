"""
Project 1 - Quiz/Exam Improvement
JOSHUA KURT M. MANZANO
BSIT 3-3
"""

import csv
import os
import re
import sys
from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher
from itertools import combinations

# =====================================================
#  EDIT ONLY THIS SECTION
# =====================================================
# Reports are saved in this folder (created automatically)
PROJECT_FOLDER = os.path.join(os.path.expanduser("~"), "Documents", "Quiz_Exam_Project")

# Cheating-check settings
MIN_SHARED_WRONG = 3        # shared identical WRONG answers needed to raise a flag
SHARED_WRONG_RATIO = 0.7    # ...and at least this share of the weaker student's wrong answers
TEXT_SIMILARITY = 0.85      # written answers this similar are flagged
MIN_TEXT_LENGTH = 40        # only compare written answers at least this long
CLOSE_SUBMIT_SECONDS = 60   # submitted within this many seconds = supporting evidence
# =====================================================
#  EDIT ONLY THIS SECTION
# ===================================================















# ---------- CHEATING CHECKER ----------

IDENTITY_COLS = ("email address", "email", "name", "full name", "student name", "student id")
SKIP_COLS = ("timestamp", "score")


def norm(text):
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def read_table(path):
    """Read a .csv or .xlsx file into a list of dicts (first row = headers)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise RuntimeError("Excel files need one extra install. Open a terminal and run:\n"
                               "    pip install openpyxl\n"
                               "Or save your sheet as .csv and use that instead.")
        ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]

        def cell(v):
            if v is None:
                return ""
            if isinstance(v, datetime):
                return v.strftime("%Y-%m-%d %H:%M:%S")
            return str(v)

        return [{h: cell(v) for h, v in zip(headers, r) if h} for r in rows[1:]
                if any(c is not None and str(c).strip() for c in r)]
    if ext == ".xls":
        raise RuntimeError("Old .xls files are not supported. In Excel, use Save As > .xlsx or .csv.")
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def parse_time(text):
    text = re.sub(r"\s*GMT.*$", "", (text or "").strip())
    for fmt in ("%Y/%m/%d %I:%M:%S %p", "%Y/%m/%d %H:%M:%S", "%m/%d/%Y %H:%M:%S",
                "%m/%d/%Y %I:%M:%S %p", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def analyze(responses, key_row=None):
    """Return (findings, notes). Findings are sorted by points, highest first."""
    if len(responses) < 2:
        return [], ["Need at least 2 responses to compare."]

    cols = list(responses[0].keys())
    id_col = next((c for c in cols if norm(c) in IDENTITY_COLS), None)
    ts_col = next((c for c in cols if norm(c) == "timestamp"), None)
    questions = [c for c in cols if norm(c) not in IDENTITY_COLS + SKIP_COLS]

    names = [((responses[i].get(id_col) or "").strip() if id_col else "") or f"Row {i + 2}"
             for i in range(len(responses))]
    notes = []

    # Work out the "correct" answer for each question
    correct = {}
    if key_row:
        for q in questions:
            if norm(key_row.get(q)):
                correct[q] = norm(key_row[q])
        notes.append(f"Using your answer key ({len(correct)} questions with a key).")
    else:
        for q in questions:
            counts = Counter(norm(r.get(q)) for r in responses if norm(r.get(q)))
            if counts:
                correct[q] = counts.most_common(1)[0][0]
        notes.append("No answer key given: the most common answer to each question is "
                     "treated as 'correct'. Results are less reliable. Provide a key for best accuracy.")

    # Which questions are long written answers?
    def is_text_q(q):
        return sum(len(norm(r.get(q))) >= MIN_TEXT_LENGTH for r in responses) >= 2

    text_qs = [q for q in questions if is_text_q(q)]
    choice_qs = [q for q in questions if q in correct and q not in text_qs]

    wrong = []
    for r in responses:
        wrong.append({q: norm(r.get(q)) for q in choice_qs
                      if norm(r.get(q)) and norm(r.get(q)) != correct[q]})

    times = [parse_time(r.get(ts_col, "")) if ts_col else None for r in responses]
    findings = []

    # Pairwise checks
    for i, j in combinations(range(len(responses)), 2):
        points, reasons = 0, []

        shared = [q for q in wrong[i] if q in wrong[j] and wrong[i][q] == wrong[j][q]]
        weaker = min(len(wrong[i]), len(wrong[j]))
        if len(shared) >= MIN_SHARED_WRONG and weaker and len(shared) / weaker >= SHARED_WRONG_RATIO:
            points += len(shared)
            reasons.append(f"Same WRONG answer on {len(shared)} questions "
                           f"({len(shared)}/{weaker} of the weaker student's wrong answers)")

        for q in text_qs:
            a, b = norm(responses[i].get(q)), norm(responses[j].get(q))
            if len(a) >= MIN_TEXT_LENGTH and len(b) >= MIN_TEXT_LENGTH:
                ratio = SequenceMatcher(None, a, b).ratio()
                if ratio >= TEXT_SIMILARITY:
                    points += 3
                    reasons.append(f"Written answer {ratio:.0%} similar: \"{q[:50]}\"")

        if reasons and times[i] and times[j]:
            gap = abs((times[i] - times[j]).total_seconds())
            if gap <= CLOSE_SUBMIT_SECONDS:
                points += 2
                reasons.append(f"Submitted {int(gap)} seconds apart")

        if points >= 4:
            findings.append({"people": (names[i], names[j]), "points": points, "reasons": reasons,
                             "level": "HIGH" if points >= 8 else "REVIEW"})

    # Duplicate submissions by the same person
    if id_col:
        counts = Counter(norm(r.get(id_col)) for r in responses if norm(r.get(id_col)))
        for ident, n in counts.items():
            if n > 1:
                findings.append({"people": (ident,), "points": 3, "level": "REVIEW",
                                 "reasons": [f"Submitted {n} times (possible retakes or shared account)"]})

    findings.sort(key=lambda f: f["points"], reverse=True)
    return findings, notes


def build_report(findings, notes, total):
    lines = ["EXAM INTEGRITY REPORT", "=" * 50,
             f"Generated: {datetime.now():%Y-%m-%d %H:%M}", f"Responses checked: {total}", ""]
    lines += [f"Note: {n}" for n in notes]
    lines.append("")
    if not findings:
        lines.append("No suspicious patterns found with the current settings.")
    for f in findings:
        lines.append(f"[{f['level']}] {' <-> '.join(f['people'])}  (score {f['points']})")
        lines += [f"    - {r}" for r in f["reasons"]]
        lines.append("")
    lines += ["-" * 50,
              "IMPORTANT: These are flags for human review, NOT proof of cheating.",
              "Students who studied together or share a teacher can have similar answers.",
              "Talk to the student and look at the actual answers before deciding anything."]
    return "\n".join(lines)


def run_check(responses_path, key_path=None):
    responses = read_table(responses_path)
    key_row = read_table(key_path)[0] if key_path else None
    findings, notes = analyze(responses, key_row)
    report = build_report(findings, notes, len(responses))
    os.makedirs(PROJECT_FOLDER, exist_ok=True)
    out = os.path.join(PROJECT_FOLDER, f"Cheating_Report_{datetime.now():%Y%m%d_%H%M}.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(report)
    return report, out











# ---------- APP Graphics --------

def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Project 1 - Quiz/Exam Improvement")
    root.geometry("720x520")

    tk.Label(root, text="Quiz/Exam Improvement", font=("Segoe UI", 16, "bold")).pack(pady=(14, 2))
    tk.Label(root, text="Check exam responses for signs of cheating", fg="#555").pack()

    out = scrolledtext.ScrolledText(root, wrap="word", font=("Consolas", 10))

    def show(text):
        out.delete("1.0", "end")
        out.insert("end", text)

    def on_check():
        resp = filedialog.askopenfilename(title="Choose responses file (.csv or .xlsx)",
                                          filetypes=[("Spreadsheets", "*.csv *.xlsx *.xlsm")])
        if not resp:
            return
        key = None
        if messagebox.askyesno("Answer key", "Do you have an answer key file?\n\n"
                               "(One header row + one row of correct answers.)\n"
                               "It makes the check much more accurate."):
            key = filedialog.askopenfilename(title="Choose answer key file (.csv or .xlsx)",
                                             filetypes=[("Spreadsheets", "*.csv *.xlsx *.xlsm")]) or None
        try:
            report, path = run_check(resp, key)
            show(report + f"\n\nSaved to: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read that file:\n{e}")

    bar = tk.Frame(root)
    bar.pack(pady=10)
    tk.Button(bar, text="Check for Cheating", width=26, height=2, command=on_check).pack()
    out.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    show("How to export responses:\n"
         "Google Form > Responses tab > 3-dot menu > Download responses (.csv)\n"
         "Or from Google Sheets: File > Download > Microsoft Excel (.xlsx)\n")
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "check":
        rep, saved = run_check(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        print(rep)
        print(f"\nSaved to: {saved}")
    else:
        gui()