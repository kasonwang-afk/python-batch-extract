import os
import pandas as pd
from datetime import datetime
# 📁 CONFIGURATION
root_dir = r'C:\Users\kason.wang\OneDrive - 大廣國際廣告股份有限公司\桌面\DATA\Maybe Python'
target_show = '好運來'
safe_target = target_show.replace(' ', '').replace('/', '_').replace('\\', '_')
timestamp = datetime.now().strftime('%Y%m%d_%H%M')

# ✅ NEW OUTPUT LOCATION
output_dir = r'C:\Users\kason.wang\OneDrive - 大廣國際廣告股份有限公司\桌面\Search Result'
output_file = os.path.join(
    output_dir,
    f"{safe_target}_summary_{timestamp}.xlsx"
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
                    df.columns = df.columns.astype(str).str.replace('\n', '', regex=True).str.strip()

                    # 🔍 Detect show name column
                    name_col = None
                    for col in df.columns:
                        if 'Name' in col or '節目名稱' in col or '節目' in col:
                            name_col = col
                            break

                    if name_col:
                        df[name_col] = df[name_col].astype(str).str.strip()
                        filtered = df[df[name_col].str.contains(target_show, na=False)]
                        if not filtered.empty:
                            filtered = filtered.copy()
                            filtered.loc[:, 'Source_File'] = filename
                            filtered.loc[:, 'Sheet'] = sheet_name
                            all_matches.append(filtered)
                    else:
                        print(f"⚠️ No recognizable show name column in {filename} / {sheet_name}")
            except Exception as e:
                print(f"❌ Error reading {file_path}: {e}")

# 🧾 Combine and export
if all_matches:
    result_df = pd.concat(all_matches, ignore_index=True)
    print("🧾 Columns in result_df:", result_df.columns.tolist())

    # 🔍 Inspect available columns
    print("🧾 Available columns:", result_df.columns.tolist())

    # ✅ Select only desired columns that exist
    available_columns = [col for col in desired_columns if col in result_df.columns]
    result_df = result_df[available_columns]

    # 💾 Export to Excel
    result_df.to_excel(output_file, index=False)
    print(f"✅ Extracted {len(result_df)} rows to {output_file}")
else:
    print("⚠️ No matching rows found.")