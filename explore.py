import pandas as pd, glob
pd.set_option('display.width', 250); pd.set_option('display.max_columns', 60)
f = glob.glob("*.xlsx")[0]
print("FILE:", f)
xl = pd.ExcelFile(f)
print("SHEETS:", xl.sheet_names)
for s in xl.sheet_names:
    df = xl.parse(s)
    print("\n" + "="*100)
    print(f"SHEET: {s} | shape={df.shape}")
    print("COLUMNS:", list(df.columns))
    print("DTYPES:\n", df.dtypes.to_string())
    print("HEAD:\n", df.head(5).to_string())
