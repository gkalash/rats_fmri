import os
import shutil
import subprocess
import glob
import json
import pandas as pd
import numpy as np
import openpyxl

# ================= CONFIGURATION =================
# Path to the folder containing 'rat15', 'rat12', etc.
DATA_ROOT = "/neurospin/micromri/Georges/rats_white_matter/data/ratsvisualstimulation"
# Path to your Excel file
EXCEL_PATH = os.path.join(DATA_ROOT, "VisRat_metadata_NeuroSpin.xlsx")
# Output BIDS directory
BIDS_ROOT = "/neurospin/micromri/Georges/rats_white_matter/data/bids_output"

# Tolerance for matching TR (Time of Repetition)
# We use a small tolerance because float numbers can vary slightly (e.g. 1.0 vs 0.9999)
TR_TOLERANCE = 0.05 
# =================================================

def setup_bids():
    if not os.path.exists(BIDS_ROOT):
        os.makedirs(BIDS_ROOT)
    
    desc = {
        "Name": "Rat Visual Stimulation Dataset",
        "BIDSVersion": "1.8.0",
        "DatasetType": "raw",
        "Authors": ["NeuroSpin Team"]
    }
    with open(os.path.join(BIDS_ROOT, "dataset_description.json"), 'w') as f:
        json.dump(desc, f, indent=4)

def clean_id(val):
    """Normalize ID: 'Rat 7' -> 'rat7'"""
    return str(val).lower().strip().replace(' ', '')

def convert_dataset():
    # 1. Load and Clean Excel Metadata
    try:
        df = pd.read_excel(EXCEL_PATH)
        
        # --- FIX 1: FILL DOWN MISSING VALUES ---
        # If Rat6 has empty TRs, copy from Rat7 above it.
        df[['anat.TR', 'func.TR']] = df[['anat.TR', 'func.TR']].ffill()
        
        # --- FIX 2: NORMALIZE IDs ---
        # Create a matching column that is all lowercase, no spaces
        df['clean_id'] = df['rat.ID'].apply(clean_id)
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    setup_bids()
    
    # Generate participants.tsv
    bids_df = pd.DataFrame()
    bids_df['participant_id'] = df['clean_id'].apply(lambda x: f"sub-{x}")
    # Map sex if column exists and isn't empty
    if 'rat.gender' in df.columns:
        bids_df['sex'] = df['rat.gender'].fillna('n/a')
    bids_df.to_csv(os.path.join(BIDS_ROOT, "participants.tsv"), sep='\t', index=False)

    # 2. Iterate through Rat Folders
    subject_dirs = [d for d in os.listdir(DATA_ROOT) if os.path.isdir(os.path.join(DATA_ROOT, d)) and d.lower().startswith('rat')]

    for sub_dir in subject_dirs:
        # Clean the folder name to match the Excel (e.g., rat7 -> rat7)
        clean_sub_name = clean_id(sub_dir)
        
        print(f"--- Processing {sub_dir} (ID: {clean_sub_name}) ---")
        
        # Find corresponding row in Excel
        row = df[df['clean_id'] == clean_sub_name]
        
        if row.empty:
            print(f"  Warning: '{clean_sub_name}' found on disk but NOT in Excel 'rat.ID' column. Skipping.")
            continue
        
        # --- FIX 3: UNITS ---
        # Excel is in SECONDS. We use values directly.
        try:
            anat_tr_target = float(row['anat.TR'].values[0])
            func_tr_target = float(row['func.TR'].values[0])
            print(f"  Target TRs -> Anat: {anat_tr_target}s, Func: {func_tr_target}s")
        except:
            print(f"  Error: TR values missing or not numbers for {clean_sub_name}")
            continue

        # Temp folder
        temp_out = os.path.join(DATA_ROOT, sub_dir, "temp_bids")
        os.makedirs(temp_out, exist_ok=True)

        # Run dcm2niix
        cmd = ["dcm2niix", "-b", "y", "-z", "y", "-f", "%s_%p", "-o", temp_out, os.path.join(DATA_ROOT, sub_dir)]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # Organize Files
        json_files = glob.glob(os.path.join(temp_out, "*.json"))
        
        if not json_files:
            print("  No NIfTI files created. Check if folder contains raw Bruker data.")
        
        for json_file in json_files:
            nii_file = json_file.replace(".json", ".nii.gz")
            if not os.path.exists(nii_file): continue

            # Get TR from the converted image JSON
            try:
                with open(json_file, 'r') as f:
                    data = json.load(f)
                    actual_tr = data.get('RepetitionTime', -1)
            except:
                actual_tr = -1

            dest_folder = None
            new_name = None
            
            # --- MATCHING LOGIC ---
            if abs(actual_tr - func_tr_target) < TR_TOLERANCE:
                dest_folder = "func"
                # You can pull task name from Excel if you want, e.g. row['func.sensory.stimulation']
                new_name = f"sub-{clean_sub_name}_ses-01_task-visual_bold"
                
                # Add TaskName for BIDS compliance
                with open(json_file, 'r+') as f:
                    jdata = json.load(f)
                    jdata['TaskName'] = "visual"
                    f.seek(0)
                    json.dump(jdata, f, indent=4)
                    f.truncate()

            elif abs(actual_tr - anat_tr_target) < TR_TOLERANCE:
                dest_folder = "anat"
                new_name = f"sub-{clean_sub_name}_ses-01_T2w"
            
            else:
                # Optional: Uncomment to see what is being skipped
                # print(f"  Skipping file with TR={actual_tr}s (Expected A:{anat_tr_target}, F:{func_tr_target})")
                continue

            # Move
            final_dir = os.path.join(BIDS_ROOT, f"sub-{clean_sub_name}", "ses-01", dest_folder)
            os.makedirs(final_dir, exist_ok=True)
            
            shutil.move(nii_file, os.path.join(final_dir, new_name + ".nii.gz"))
            shutil.move(json_file, os.path.join(final_dir, new_name + ".json"))
            print(f"  + Converted: {dest_folder}/{new_name}")

        shutil.rmtree(temp_out)

if __name__ == "__main__":
    convert_dataset()
    print("BIDS Conversion Finished!")