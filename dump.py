import pandas as pd, glob
pd.set_option('display.width', 300); pd.set_option('display.max_columns', 60); pd.set_option('display.max_rows', 100)
f = glob.glob("*.xlsx")[0]
for s in ['نمونه پنبه','نمونه پلی استر']:
    df = pd.read_excel(f, sheet_name=s)
    cols=['کد نمونه','پارچه','نمره نخ قرره\nپنبه (ne)','کشش','تکرار','کرنش اولیه دوخت','درصد آبرفت کل در شستشو اول','درصد آبرفت کل در شستشو دوم','شاخص پسماند آزاد شده 1','شاخص پسماند آزاد شده 2','شاخص تفاضل شستشو 2']
    print("="*110); print("SHEET:", s, df.shape)
    print(df[cols].to_string())
    print("unique کشش:", df['کشش'].dropna().unique())
    print("unique نمره:", df['نمره نخ قرره\nپنبه (ne)'].dropna().unique())
