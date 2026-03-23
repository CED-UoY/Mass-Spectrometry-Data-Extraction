# -*- coding: utf-8 -*-
"""
Created on Mon Mar 23 10:43:04 2026

@author: wsl530
"""

import tkinter as tk
from tkinter import filedialog
import os
from pathlib import Path
import pandas as pd
import numpy as np
from collections import defaultdict
import math
import openpyxl
from openpyxl.styles import Font

def get_max_runtime(file_paths, mode):
    """Quickly scans all selected files to find the highest run time for hints."""
    max_time = 0.0
    for f_path in file_paths:
        try:
            with open(f_path, 'r', encoding='utf-8', errors='ignore') as f:
                if mode == 'THERMO':  
                    for line in f:
                        if line.startswith('I') and 'RTime' in line:
                            try:
                                t = float(line.split()[2])
                                if t > max_time: max_time = t
                            except: pass
                elif mode == 'BRUKER':  
                    for line in f:
                        parts = line.strip().split(',')
                        if len(parts) >= 9 and 'ESI' in parts:
                            try:
                                t = float(parts[0])
                                if t > max_time: max_time = t
                            except: pass
        except Exception: pass
    return max_time

def parse_ms_data(file_path, mode, num_segments, seg_time):
    """Parses data into discrete segments regardless of instrument."""
    bins = defaultdict(lambda: defaultdict(list))
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            if mode == 'THERMO':  
                rtime = -1.0
                for line in f:
                    if line.startswith('I') and 'RTime' in line:
                        try: rtime = float(line.split()[2])
                        except: pass
                    elif rtime >= 0 and line and line[0].isdigit():
                        seg_idx = int(rtime // seg_time)
                        if seg_idx >= num_segments: continue
                        parts = line.split(maxsplit=2)
                        mz = int(round(float(parts[0])))
                        bins[seg_idx][mz].append(float(parts[1]))
            elif mode == 'BRUKER':  
                for line in f:
                    parts = line.strip().split(',')
                    if len(parts) >= 9 and 'ESI' in parts:
                        try:
                            time = float(parts[0])
                            seg_idx = int(time // seg_time)
                            if seg_idx >= num_segments: continue 
                            for pair in parts[8:]:
                                if ' ' in pair:
                                    mz_str, int_str = pair.strip().split()
                                    mz = int(round(float(mz_str)))
                                    bins[seg_idx][mz].append(float(int_str))
                        except: pass
    except Exception as e:
        print(f"⚠️ Error reading file {Path(file_path).name}: {e}")
        return None

    if not bins: return None

    df_rows = []
    all_mz_set = set()
    for seg_idx in range(num_segments):
        if seg_idx in bins:
            avg_dict = {mz: np.mean(vals) for mz, vals in bins[seg_idx].items()}
            avg_dict['Segment'] = seg_idx + 1
            all_mz_set.update(bins[seg_idx].keys())
            df_rows.append(avg_dict)
        else:
            df_rows.append({'Segment': seg_idx + 1})
            
    df = pd.DataFrame(df_rows).fillna(0.0)
    return df[['Segment'] + sorted(list(all_mz_set))]

# ---------------------------------------------------------
# UI HELPER FUNCTIONS
# ---------------------------------------------------------

def custom_msgbox(parent, title, message, is_error=False):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("550x250")
    win.attributes('-topmost', True)
    win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), parent.quit()))
    color = "red" if is_error else "darkorange"
    header = "⚠️ ERROR" if is_error else "⚠️ WARNING"
    tk.Label(win, text=header, font=("Calibri", 18, "bold"), fg=color).pack(pady=10)
    tk.Label(win, text=message, font=("Calibri", 14), wraplength=500, justify="center").pack(pady=15)
    tk.Button(win, text="OK", font=("Calibri", 14, "bold"), command=lambda: (win.destroy(), parent.quit()), width=15).pack(pady=10)
    parent.mainloop()

def custom_askyesno(parent, title, message):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("450x220")
    win.attributes('-topmost', True)
    result = tk.BooleanVar(value=False)
    def set_result(answer):
        result.set(answer)
        win.destroy()
        parent.quit()
    win.protocol("WM_DELETE_WINDOW", lambda: set_result(False))
    tk.Label(win, text=title, font=("Calibri", 16, "bold"), fg="#0066cc").pack(pady=15)
    tk.Label(win, text=message, font=("Calibri", 14), wraplength=400, justify="center").pack(pady=10)
    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=20)
    tk.Button(btn_frame, text="Yes", font=("Calibri", 14, "bold"), width=10, bg="green", fg="white", command=lambda: set_result(True)).pack(side="left", padx=15)
    tk.Button(btn_frame, text="No", font=("Calibri", 14, "bold"), width=10, bg="red", fg="white", command=lambda: set_result(False)).pack(side="left", padx=15)
    parent.mainloop()
    return result.get()

# ---------------------------------------------------------
# MAIN SCRIPT
# ---------------------------------------------------------

def main():
    root = tk.Tk()
    root.withdraw()
    
    font_title = ("Calibri", 16, "bold")
    font_label = ("Calibri", 14)
    font_entry = ("Calibri", 14)
    font_btn = ("Calibri", 14, "bold")
    font_tooltip = ("Calibri", 12, "italic")

    # =========================================================
    # STEP 1: Select Master Track
    # =========================================================
    track_choice = tk.StringVar(value="")
    
    win1 = tk.Toplevel(root)
    win1.title("Step 1: Select Instrument")
    win1.geometry("500x220") 
    win1.attributes('-topmost', True)
    win1.protocol("WM_DELETE_WINDOW", lambda: (win1.destroy(), root.quit())) 
    
    tk.Label(win1, text="Which instrument data are you processing?", font=font_title).pack(pady=20)
    
    def set_track(val):
        track_choice.set(val)
        win1.destroy()
        root.quit() 
        
    tk.Button(win1, text="Thermo HCD / TRMS (.ms1 / .ms2)", font=font_label, command=lambda: set_track('THERMO'), width=40, bg="#e6f2ff").pack(pady=5)
    tk.Button(win1, text="Bruker CID / TRMS (.ascii)", font=font_label, command=lambda: set_track('BRUKER'), width=40).pack(pady=5)
    
    root.mainloop() 
    
    mode = track_choice.get()
    if not mode: return root.destroy()

    # =========================================================
    # STEP 2: File Selection & Warnings
    # =========================================================
    if mode == 'THERMO':
        custom_msgbox(root, "Thermo Data", "Ensure your files are converted from .raw to .ms1 or .ms2 using MSConvert.")
        file_types = [("Thermo MS Files", "*.ms1 *.ms2"), ("All files", "*.*")]
        title = "Select Thermo (.ms1 / .ms2) Files"
    else:
        custom_msgbox(root, "Bruker Data", "Ensure you have run 'Export_data_seg_times' to generate .ascii files.")
        file_types = [("Bruker ASCII Files", "*.ascii"), ("All files", "*.*")]
        title = "Select Bruker (.ascii) Files"

    raw_selection = filedialog.askopenfilenames(parent=root, title=title, filetypes=file_types)
    file_paths = list(raw_selection) if raw_selection else []
    if not file_paths: return root.destroy()

    while True:
        more = custom_askyesno(root, "Add More Files?", f"You currently have {len(file_paths)} file(s) selected.\n\nDo you want to add more?")
        if more:
            extra_files = filedialog.askopenfilenames(parent=root, title="Select Additional Files", filetypes=file_types)
            if extra_files: file_paths.extend(list(extra_files))
            else: break  
        else: break  

    # =========================================================
    # STEP 3: Unified Parameter Dialog
    # =========================================================
    settings = {}
    win2 = tk.Toplevel(root)
    win2.title("Step 3: Extraction Settings")
    win2.geometry("650x480") 
    win2.attributes('-topmost', True)
    win2.protocol("WM_DELETE_WINDOW", lambda: (win2.destroy(), root.quit()))

    def add_row(parent_win, label_text, var, side_text="", tooltip=""):
        frame = tk.Frame(parent_win)
        frame.pack(fill='x', padx=20, pady=8)
        tk.Label(frame, text=label_text, width=26, anchor='w', font=font_label).pack(side='left')
        tk.Entry(frame, textvariable=var, font=font_entry, width=12).pack(side='left')
        if side_text: tk.Label(frame, text=side_text, font=("Calibri", 12, "bold"), fg="#0066cc").pack(side='left', padx=10)
        if tooltip: tk.Label(parent_win, text=tooltip, font=font_tooltip, fg="gray").pack(anchor='w', padx=20)

    # UI Variables
    v_seg_count = tk.StringVar(value="")
    v_seg_time = tk.StringVar(value="")
    v_par = tk.StringVar(value="")
    v_num = tk.StringVar(value="")
    v_frg = tk.StringVar(value="")

    print(f"Scanning for max runtime...")
    detected_max_time = get_max_runtime(file_paths, mode)
    detected_max_str = f"← Detected Max Run: ~{math.ceil(detected_max_time)} min" if detected_max_time > 0 else ""
    
    add_row(win2, "Number of Segments:", v_seg_count, side_text=detected_max_str)
    add_row(win2, "Time per Segment (min):", v_seg_time)
    add_row(win2, "Precursor m/z:", v_par, tooltip="Leave blank to auto-detect from Run 1.")
    add_row(win2, "Number of Expected Fragments:", v_num)
    add_row(win2, "Expected m/z fragments:", v_frg, tooltip="Comma-separated. Leave blank for auto top peaks.")

    def submit_params():
        try:
            settings['num_segments'] = int(v_seg_count.get())
            settings['segment_time'] = float(v_seg_time.get())
            settings['parent_mz'] = v_par.get().strip()
            settings['num_frag'] = int(v_num.get())
            settings['frag_list'] = v_frg.get().strip()
                
            win2.destroy()
            root.quit()
        except ValueError:
            custom_msgbox(win2, "Input Error", "Missing or Invalid Input. Ensure all numeric fields contain numbers.", is_error=True)

    btn_frame = tk.Frame(win2)
    btn_frame.pack(pady=25)
    tk.Button(btn_frame, text="Submit", font=font_btn, command=submit_params, width=15, bg="green", fg="white").pack(side='left', padx=15)
    tk.Button(btn_frame, text="Cancel", font=font_btn, command=lambda: (win2.destroy(), root.quit()), width=15).pack(side='left', padx=15)
    root.mainloop()

    if not settings: return root.destroy()

    # Parse common variables
    global_parent = int(settings["parent_mz"]) if settings["parent_mz"] else None
    p_num = settings["num_frag"]
    global_products = [int(x.strip()) for x in settings["frag_list"].split(',')] if settings["frag_list"] else []
    target_header = 'Segment'

    # STRICT ENFORCER: Ensure manual list doesn't exceed the requested number right away
    global_products = global_products[:p_num]

    # =========================================================
    # STEP 4: Save & Process
    # =========================================================
    output_path = filedialog.asksaveasfilename(parent=root, defaultextension=".xlsx", title="Step 4: Save Excel As", filetypes=[("Excel files", "*.xlsx")])
    root.destroy() 
    if not output_path: return

    print(f"\n--- Starting Extraction ---")
    all_data = [] 
    
    for idx, f_path in enumerate(file_paths, 1):
        filename = Path(f_path).name
        print(f"\n📂 Processing Run {idx}: {filename}...")
        
        df = parse_ms_data(f_path, mode, settings['num_segments'], settings['segment_time'])
        if df is None: continue

        # Auto-detect precursor if not supplied (using total intensity sum)
        if global_parent is None:
            mz_sums = df.drop(columns=[target_header]).sum()
            if not mz_sums.empty and mz_sums.max() > 0:
                global_parent = mz_sums.idxmax()
                print(f"  ✅ Auto-detected precursor: {global_parent}")
            else: continue

        # Auto-detect fragments by overall intensity SUM (not max spike)
        if len(global_products) < p_num:
            # Calculate total intensity of each m/z across the entire run
            sum_ints = df.drop(columns=[target_header]).drop(labels=[global_parent], errors='ignore').sum()
            
            # Sort by the highest total intensity and grab the top peaks
            for mz in sum_ints.sort_values(ascending=False).index:
                # Extra check: Ignore it if the total intensity is practically 0
                if sum_ints[mz] <= 0: continue 
                
                if mz not in global_products and mz != global_parent:
                    global_products.append(mz)
                if len(global_products) >= p_num: 
                    break

        # STRICT ENFORCER: Double-check that we never exceed p_num fragments
        global_products = global_products[:p_num]
        
        if not settings["frag_list"] and idx == 1: 
            print(f"  ✅ High-intensity fragments locked: {global_products}")

        # Map out the exact columns we will allow in the final Excel file
        target_cols = [target_header]
        if global_parent not in target_cols: target_cols.append(global_parent)
        for mz in global_products:
            if mz not in target_cols: target_cols.append(mz)

        # Ensure all required columns exist (so pandas doesn't crash if a peak was missing)
        for col in target_cols:
            if col not in df.columns: df[col] = 0.0

        # Create a clean dataframe containing ONLY our strictly defined columns
        export_df = df[target_cols].copy()
        
        # Create formatting rows (Blank, Title, Header, Data)
        if idx > 1:
            all_data.append(pd.DataFrame([{col: "" for col in target_cols}]))
            
        all_data.append(pd.DataFrame([{col: f"Run {idx}: {filename}" if col == target_header else "" for col in target_cols}]))
        all_data.append(pd.DataFrame([{col: str(col) for col in target_cols}]))
        all_data.append(export_df)
        print(f"  ✅ Extraction for Run {idx} complete.")

    if all_data:
        try:
            pd.concat(all_data, ignore_index=True).to_excel(output_path, sheet_name=f"Extracted Data", index=False, header=False)
            
            # --- MAKE HEADERS BOLD ---
            wb = openpyxl.load_workbook(output_path)
            ws = wb.active
            bold_font = Font(bold=True)
            for row in ws.iter_rows():
                val = str(row[0].value) if row[0].value else ""
                if val == target_header or val.startswith("Run"):
                    for cell in row: cell.font = bold_font
            wb.save(output_path)
            
            print(f"\n✅ SUCCESS! Saved {len(file_paths)} files to:\n   {output_path}")
        except Exception as e: print(f"\n❌ Error saving file: {e}")

if __name__ == "__main__":
    main()