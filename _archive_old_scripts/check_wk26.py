import pandas as pd

df = pd.read_excel(r"D:\ASUS\News & Report\News\News.xlsx", sheet_name="News").head(60).fillna("")
valid = df[df["Topic"].astype(str).str.strip() != ""]
yr = int(float(valid.iloc[0]["Year"]))
wk = int(float(valid.iloc[0]["Week"]))
print(f"Year={yr}, Week={wk}")

def match(r):
    try:
        return int(float(r["Year"])) == yr and int(float(r["Week"])) == wk
    except:
        return False

cur = df[df.apply(match, axis=1)]
print(f"Total rows for wk{wk}: {len(cur)}")
print("---Categories---")
print(cur["Category"].value_counts().to_string())
print("---Topics---")
for _, r in cur.iterrows():
    cat = r["Category"]
    topic = str(r["Topic"])[:70]
    content = str(r["Content"])[:80]
    impact = str(r.get("對於筆電市場/華碩的影響", ""))[:80]
    print(f"  [{cat}] {topic}")
    print(f"    摘要: {content}")
    print(f"    影響: {impact}")
    print()
