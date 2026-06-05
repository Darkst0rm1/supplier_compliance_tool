import sys, io
sys.path.insert(0, ".")
from src.fill_rate_engine import load_fill_rate

with open(r"C:\Users\melgh\Downloads\Delivery.xlsx", "rb") as f:
    data = f.read()

df_clean, _ = load_fill_rate(io.BytesIO(data), threshold=1000.0)

print("=== ALL COLUMNS WITH POSITION (1-based Excel letter) ===")
import string

def col_letter(n):  # n is 0-based index
    n += 1  # to 1-based
    result = ""
    while n:
        n, r = divmod(n - 1, 26)
        result = string.ascii_uppercase[r] + result
    return result

for i, col in enumerate(df_clean.columns):
    sample = df_clean[col].dropna().iloc[0] if not df_clean[col].dropna().empty else "N/A"
    print(f"  Excel col {col_letter(i)} (index {i}): [{col}]  |  sample: {sample}")

print()
print(f"Column C = [{df_clean.columns[2]}]")
print(f"Column N = [{df_clean.columns[13]}]")
