import json
import pandas as pd

# Baca file JSON
with open("KevinFataZacky_V3925046.json", "r", encoding="utf-8") as f:
    data = json.load(f)

print("Daftar key yang ada di file JSON:")
print(list(data.keys()))  # biar kelihatan key apa aja yang tersedia

# Buat file Excel
output_file = "KevinFataZacky_V3925046.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for key, value in data.items():
        try:
            df = pd.DataFrame(value)
        except Exception:
            df = pd.DataFrame([value])
        df.to_excel(writer, sheet_name=str(key)[:30], index=False)

print(f"Selesai! File Excel berhasil dibuat: {output_file}")
