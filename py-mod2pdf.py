import os, sys, subprocess, webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
# source: https://github.com/zeittresor/py_mod2pdf
period_to_note = {
    1712: "C-0", 1616: "C#0", 1524: "D-0", 1440: "D#0", 1356: "E-0", 1280: "F-0",
    1208: "F#0", 1140: "G-0", 1076: "G#0", 1016: "A-0", 960: "A#0", 906: "B-0",
    856: "C-1", 808: "C#1", 762: "D-1", 720: "D#1", 678: "E-1", 640: "F-1",
    604: "F#1", 570: "G-1", 538: "G#1", 508: "A-1", 480: "A#1", 453: "B-1",
    428: "C-2", 404: "C#2", 381: "D-2", 360: "D#2", 339: "E-2", 320: "F-2",
    302: "F#2", 285: "G-2", 269: "G#2", 254: "A-2", 240: "A#2", 226: "B-2",
    214: "C-3", 202: "C#3", 190: "D-3", 180: "D#3", 170: "E-3", 160: "F-3",
    151: "F#3", 143: "G-3", 135: "G#3", 127: "A-3", 120: "A#3", 113: "B-3"
}

def parse_mod_data(data):
    if len(data) < 20:
        raise ValueError("File too short to be a valid MOD")
    title = data[0:20].decode('ascii', errors='ignore').rstrip('\x00')
    inst_count = 31
    channels = 4
    if len(data) >= 1084:
        tag_bytes = data[1080:1084]
        if all(32 <= c <= 126 for c in tag_bytes):
            tag = tag_bytes.decode('ascii', errors='ignore')
            if tag in ("M.K.", "M!K!", "4CHN", "FLT4"):
                channels = 4
            elif tag == "6CHN":
                channels = 6
            elif tag in ("8CHN", "FLT8"):
                channels = 8
            elif tag.endswith("CH"):
                try:
                    channels = int(tag[:-2])
                except:
                    channels = 4
            elif tag.endswith("CHN"):
                try:
                    channels = int(tag[:-3])
                except:
                    channels = 4
            inst_count = 31
        else:
            inst_count = 15
            channels = 4
    else:
        inst_count = 15
        channels = 4
    song_length_offset = 20 + inst_count * 30
    song_length = data[song_length_offset]
    order_list = list(data[song_length_offset+2 : song_length_offset+2+128])
    used_orders = order_list[:song_length]
    pattern_count = max(used_orders) + 1 if used_orders else 0
    pattern_data_offset = song_length_offset + 2 + 128
    if inst_count == 31:
        pattern_data_offset += 4
    patterns = []
    for pat in range(pattern_count):
        pat_rows = []
        offset = pattern_data_offset + pat * (64 * channels * 4)
        if offset + 64 * channels * 4 > len(data):
            break
        for row in range(64):
            row_cells = []
            for ch in range(channels):
                i = offset + (row * channels + ch) * 4
                b0 = data[i]; b1 = data[i+1]; b2 = data[i+2]; b3 = data[i+3]
                period = ((b0 & 0x0F) << 8) | b1
                instrument = (b0 & 0xF0) | ((b2 & 0xF0) >> 4)
                effect = b2 & 0x0F
                param = b3
                row_cells.append((period, instrument, effect, param))
            pat_rows.append(row_cells)
        patterns.append(pat_rows)
    return {
        "title": title,
        "order_list": used_orders,
        "patterns": patterns,
        "channels": channels
    }

from fpdf import FPDF
def patterns_to_pdf(mod_info, pdf_path):
    patterns = mod_info['patterns']
    order_list = mod_info['order_list']
    channels = mod_info['channels']
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)
    for pos, pat_num in enumerate(order_list):
        if pat_num >= len(patterns):
            continue
        pat_data = patterns[pat_num]
        pdf.add_page()
        pdf.set_font("Courier", style="B", size=12)
        pdf.cell(0, 6, f"Pattern {pat_num}", ln=1)
        pdf.set_font("Courier", size=9)
        col_width = 22
        row_width = 7
        line_height = 4
        for r, row in enumerate(pat_data):
            pdf.cell(row_width, line_height, f"{r:02d}", border=1)
            for ch in range(channels):
                period, instrument, effect, param = row[ch]
                if period != 0:
                    note = period_to_note.get(period, f"{period:03d}")
                else:
                    note = "---"
                instr_str = f"{instrument:02X}"
                eff_str = f"{effect:X}{param:02X}"
                cell_text = f"{note} {instr_str} {eff_str}"
                pdf.cell(col_width, line_height, cell_text, border=1)
            pdf.ln(line_height)
    pdf.output(pdf_path)
    return True

selected_file = ""
script_dir = os.path.abspath(os.path.dirname(__file__)) if '__file__' in globals() else os.getcwd()
output_dir = os.path.join(script_dir, "output")

def select_file():
    global selected_file
    file_path = filedialog.askopenfilename(filetypes=[("MOD files", "*.mod"), ("All files", "*.*")])
    if file_path:
        selected_file = file_path
        file_label.config(text=f"Selected file: {os.path.basename(file_path)}")

def save_to_pdf():
    if not selected_file:
        messagebox.showwarning("No file selected", "Please select a .mod file first.")
        return
    os.makedirs(output_dir, exist_ok=True)
    try:
        with open(selected_file, "rb") as f:
            data = f.read()
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file:\n{e}")
        return
    try:
        mod_info = parse_mod_data(data)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to parse MOD file:\n{e}")
        return
    pdf_name = os.path.splitext(os.path.basename(selected_file))[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    try:
        patterns_to_pdf(mod_info, pdf_path)
        messagebox.showinfo("Success", f"Patterns saved in PDF:\n{pdf_name}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save PDF:\n{e}")

def open_output_folder():
    if not os.path.isdir(output_dir):
        messagebox.showinfo("Info", "No output folder available.")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(output_dir)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", output_dir])
        else:
            subprocess.Popen(["xdg-open", output_dir])
    except Exception as e:
        try:
            webbrowser.open(f"file://{output_dir}")
        except:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

root = tk.Tk()
root.title("MOD Pattern Extractor")
root.resizable(False, False)

file_label = tk.Label(root, text="Selected file: None", anchor="w")
select_btn = tk.Button(root, text="Select .mod file", command=select_file)
save_btn = tk.Button(root, text="Save Patterns as PDF", command=save_to_pdf)
open_btn = tk.Button(root, text="Open Output Folder", command=open_output_folder)

select_btn.pack(padx=10, pady=5, fill="x")
save_btn.pack(padx=10, pady=5, fill="x")
open_btn.pack(padx=10, pady=5, fill="x")
file_label.pack(padx=10, pady=5, fill="x")

root.mainloop()
