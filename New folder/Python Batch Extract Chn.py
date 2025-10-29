import os
import pandas as pd
from datetime import datetime

# 📁 CONFIGURATION
root_dir = r'C:\Users\kason.wang\OneDrive - 大廣國際廣告股份有限公司\桌面\DATA\Maybe Python'
target_channel = '台視'  # 🔍 Change this to the channel you want to search
safe_target = target_channel.replace(' ', '').replace('/', '_').replace('\\', '_')
timestamp = datetime.now().strftime('%Y%m%d_%H%M')
output_file = os.path.join(
    root_dir,
    f"{safe_target}_channel_summary_{timestamp}.xlsx"
)

# 📊 Desired output columns
desired_columns = [
    'Source_File', 'Sheet',
    'No.', 'Date', 'Time', 'Chn',
    'Name', '節目類型', 'Rating',
    '總來電數', '仿冒來電', '有效來電數'
]

# 📦 Collect matching rows
all_matches = []

# 🔍 Traverse folders and files
for foldername, subfolders, filenames in os.walk(root_dir):
    for filename in filenames:
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            file_path = os.path.join(foldername, filename)
            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name)

                    # 🧼 Normalize column names
                    df.columns = df.columns.astype(str).str.replace('\n', '', regex=True).str.strip()

                    # 🔍 Detect channel column
                    chn_col = None
                    for col in df.columns:
                        if 'Chn' in col or '頻道' in col:
                            chn_col = col
                            break

                    if chn_col:
                        df[chn_col] = df[chn_col].astype(str).str.strip()
                        filtered = df[df[chn_col].str.contains(target_channel, na=False)]
                        if not filtered.empty:
                            filtered = filtered.copy()
                            filtered.loc[:, 'Source_File'] = filename
                            filtered.loc[:, 'Sheet'] = sheet_name
                            all_matches.append(filtered)
                    else:
                        print(f"⚠️ No recognizable channel column in {filename} / {sheet_name}")
            except Exception as e:
                print(f"❌ Error reading {file_path}: {e}")

# 🧾 Combine and export
if all_matches:
    result_df = pd.concat(all_matches, ignore_index=True)

    # 🔍 Inspect available columns
    print("🧾 Columns in result_df:", result_df.columns.tolist())

    # ✅ Select only desired columns that exist
    available_columns = [col for col in desired_columns if col in result_df.columns]
    result_df = result_df[available_columns]

    # 💾 Export to Excel
    result_df.to_excel(output_file, index=False)
    print(f"✅ Extracted {len(result_df)} rows to {output_file}")
else:
    print("⚠️ No matching rows found.")