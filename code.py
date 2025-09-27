import json
import pandas as pd

# Baca file JSON
with open("TI-A_KevinFataZacky_V3925046", "r", encoding="utf-8") as f:
    data = json.load(f)

# Ubah setiap bagian JSON jadi DataFrame
df_cctv = pd.DataFrame(data["cctv_real_time"])
df_pemerintah = pd.DataFrame(data["informasi_pemerintah"])
df_laporan = pd.DataFrame(data["laporan_masyarakat"])
df_patroli = pd.DataFrame(data["patroli_polisi"])

# Tampilkan tabel
print("=== Data CCTV Real Time ===")
print(df_cctv.head(), "\n")

print("=== Data Informasi Pemerintah ===")
print(df_pemerintah.head(), "\n")

print("=== Data Laporan Masyarakat ===")
print(df_laporan.head(), "\n")

print("=== Data Patroli Polisi ===")
print(df_patroli.head(), "\n")

# Simpan ke Excel
with pd.ExcelWriter("TI-A_KevinFataZacky_V3925046.xlsx") as writer:
    df_cctv.to_excel(writer, sheet_name="CCTV Real Time", index=False)
    df_pemerintah.to_excel(writer, sheet_name="Informasi Pemerintah", index=False)
    df_laporan.to_excel(writer, sheet_name="Laporan Masyarakat", index=False)
    df_patroli.to_excel(writer, sheet_name="Patroli Polisi", index=False)

print("✅ Data berhasil disimpan ke TI-A_KevinFataZacky_V3925046.xlsx")
