
from pathlib import Path
from openpyxl import load_workbook

# =====================================
# 1. MEMBACA FILE EXCEL
# =====================================

lokasi_excel = (
    Path(__file__).parent
    / "TUGAS PYTHON - CBR Otomatis.xlsx"
)

if not lokasi_excel.exists():
    print("ERROR: File Excel tidak ditemukan!")
    print("Periksa nama dan lokasi file Excel.")
    raise SystemExit(1)

workbook = load_workbook(lokasi_excel, data_only=True)
sheet = workbook["Sheet1"]

print("=" * 50)
print(" SISTEM PAKAR REKOMENDASI METODE BELAJAR")
print(" METODE CASE-BASED REASONING (CBR)")
print("=" * 50)

print("\nFile Excel berhasil dibaca!")


# =====================================
# 2. MEMBACA 13 METODE BELAJAR
# =====================================

metode_belajar = {}

print("\n===== DAFTAR METODE BELAJAR =====")

for baris in range(8, 21):
    id_metode = sheet[f"B{baris}"].value
    nama_metode = sheet[f"C{baris}"].value

    if id_metode is not None and nama_metode is not None:
        id_metode = str(id_metode).strip()
        nama_metode = str(nama_metode).strip()

        metode_belajar[id_metode] = nama_metode
        print(f"{id_metode} - {nama_metode}")

print("Jumlah metode belajar:", len(metode_belajar))


# =====================================
# 3. MEMBACA 12 KONDISI MAHASISWA
# =====================================

kondisi_mahasiswa = {}

print("\n===== DAFTAR KONDISI MAHASISWA =====")

for baris in range(28, 40):
    id_kondisi = sheet[f"B{baris}"].value
    nama_kondisi = sheet[f"C{baris}"].value

    if id_kondisi is not None and nama_kondisi is not None:
        id_kondisi = str(id_kondisi).strip()
        nama_kondisi = str(nama_kondisi).strip()

        kondisi_mahasiswa[id_kondisi] = nama_kondisi
        print(f"{id_kondisi} - {nama_kondisi}")

print("Jumlah kondisi mahasiswa:", len(kondisi_mahasiswa))


# =====================================
# 4. MEMBACA 20 BASIS KASUS
# =====================================

basis_kasus = {}

print("\n===== DAFTAR BASIS KASUS =====")

for baris in range(47, 67):
    no_kasus = sheet[f"B{baris}"].value
    id_kondisi = sheet[f"C{baris}"].value
    id_metode = sheet[f"D{baris}"].value

    if all(v is not None for v in
           (no_kasus, id_kondisi, id_metode)):

        no_kasus = str(no_kasus).strip()
        id_metode = str(id_metode).strip()

        daftar_kondisi = [
            kode.strip()
            for kode in str(id_kondisi).split(",")
            if kode.strip()
        ]

        basis_kasus[no_kasus] = {
            "kondisi": daftar_kondisi,
            "metode": id_metode
        }

        nama_metode = metode_belajar.get(
            id_metode, "Metode tidak ditemukan"
        )

        print(f"\nKasus: {no_kasus}")
        print(f"Kondisi: {daftar_kondisi}")
        print(f"Metode: {id_metode} - {nama_metode}")

print("\nJumlah basis kasus:", len(basis_kasus))


# =====================================
# 5. MEMBACA RELASI BASIS PENGETAHUAN
# =====================================

relasi_pengetahuan = {}
relasi_metode = {}

print("\n===== RELASI BASIS PENGETAHUAN =====")

jumlah_relasi = 0

for baris in range(8, 59):
    no_kasus = sheet[f"H{baris}"].value
    id_metode = sheet[f"I{baris}"].value
    id_kondisi = sheet[f"J{baris}"].value

    if all(v is not None for v in
           (no_kasus, id_metode, id_kondisi)):

        no_kasus = str(no_kasus).strip()
        id_metode = str(id_metode).strip()
        id_kondisi = str(id_kondisi).strip()

        if no_kasus not in relasi_pengetahuan:
            relasi_pengetahuan[no_kasus] = []
            relasi_metode[no_kasus] = set()

        relasi_pengetahuan[no_kasus].append(id_kondisi)
        relasi_metode[no_kasus].add(id_metode)

        jumlah_relasi += 1

        print(f"{no_kasus} | {id_metode} | {id_kondisi}")

print("\nJumlah kasus relasi:", len(relasi_pengetahuan))
print("Jumlah relasi:", jumlah_relasi)


# =====================================
# 6. VALIDASI BASIS PENGETAHUAN
# =====================================

print("\n===== VALIDASI BASIS PENGETAHUAN =====")

kesalahan = []

if len(metode_belajar) != 13:
    kesalahan.append("Jumlah metode bukan 13.")

if len(kondisi_mahasiswa) != 12:
    kesalahan.append("Jumlah kondisi bukan 12.")

if len(basis_kasus) != 20:
    kesalahan.append("Jumlah basis kasus bukan 20.")

if jumlah_relasi != 51:
    kesalahan.append("Jumlah relasi bukan 51.")

for no_kasus, data in basis_kasus.items():

    kondisi_utama = set(data["kondisi"])
    kondisi_relasi = set(
        relasi_pengetahuan.get(no_kasus, [])
    )

    if kondisi_utama != kondisi_relasi:
        kesalahan.append(
            f"{no_kasus}: Kondisi utama dan relasi berbeda."
        )

    if len(data["kondisi"]) != len(kondisi_utama):
        kesalahan.append(
            f"{no_kasus}: Ada kondisi duplikat."
        )

    if len(relasi_pengetahuan.get(no_kasus, [])) != len(
        kondisi_relasi
    ):
        kesalahan.append(
            f"{no_kasus}: Ada kondisi relasi duplikat."
        )

    if data["metode"] not in metode_belajar:
        kesalahan.append(
            f"{no_kasus}: ID metode tidak ditemukan."
        )

    if relasi_metode.get(no_kasus) != {data["metode"]}:
        kesalahan.append(
            f"{no_kasus}: ID metode pada relasi berbeda."
        )

    for kode in kondisi_utama | kondisi_relasi:
        if kode not in kondisi_mahasiswa:
            kesalahan.append(
                f"{no_kasus}: Kondisi {kode} tidak dikenal."
            )

for no_kasus in relasi_pengetahuan:
    if no_kasus not in basis_kasus:
        kesalahan.append(
            f"{no_kasus}: Ada di relasi tetapi tidak di basis kasus."
        )

if kesalahan:
    print("Ditemukan masalah pada basis pengetahuan:")

    for masalah in kesalahan:
        print("-", masalah)

    print("\nPerbaiki data Excel sebelum melanjutkan.")
    raise SystemExit(1)

print("Validasi berhasil!")
print("Seluruh basis kasus dan relasi sesuai.")


# =====================================
# 7. INPUT KONDISI MAHASISWA
# =====================================

print("\n===== INPUT KONDISI MAHASISWA =====")

jawaban_mahasiswa = {}
kondisi_dipilih = []

for id_kondisi, nama_kondisi in kondisi_mahasiswa.items():

    while True:
        jawaban = input(
            f"Apakah mengalami {nama_kondisi} (Y/T)? "
        ).strip().lower()

        if jawaban in ("y", "t"):
            break

        print("Input tidak valid! Masukkan Y atau T.")

    mengalami = jawaban == "y"

    jawaban_mahasiswa[id_kondisi] = mengalami

    if mengalami:
        kondisi_dipilih.append(id_kondisi)

print("\n===== HASIL INPUT MAHASISWA =====")

print("Kondisi Dipilih:", list(jawaban_mahasiswa.values()))
print("ID Kondisi Dipilih:", kondisi_dipilih)
print("Jumlah Kondisi Dipilih:", len(kondisi_dipilih))


# =====================================
# 8. PERHITUNGAN SIMILARITY CBR
# =====================================

print("\n===== PERHITUNGAN SIMILARITY CBR =====")

hasil_similarity = []

kondisi_user = set(kondisi_dipilih)

for no_kasus, data_kasus in basis_kasus.items():

    id_metode = data_kasus["metode"]
    nama_metode = metode_belajar[id_metode]

    kondisi_kasus = set(relasi_pengetahuan[no_kasus])

    jumlah_dipilih = len(kondisi_user)
    jumlah_kasus = len(kondisi_kasus)

    kondisi_cocok = kondisi_user & kondisi_kasus
    jumlah_cocok = len(kondisi_cocok)

    pembagi = max(jumlah_dipilih, jumlah_kasus)

    similarity = (
        jumlah_cocok / pembagi
        if pembagi > 0 else 0
    )

    print(f"\n===== KASUS {no_kasus} =====")
    print(f"Metode Belajar: {nama_metode}")
    print(f"Jumlah Kondisi Dipilih: {jumlah_dipilih}")
    print(f"Jumlah Kondisi di Kasus: {jumlah_kasus}")
    print(f"Jumlah Kondisi Cocok: {jumlah_cocok}")
    print(
        f"Pembagi: Max({jumlah_dipilih}, "
        f"{jumlah_kasus}) = {pembagi}"
    )
    print(
        f"Hasil: {jumlah_cocok} / "
        f"{pembagi} = {similarity:.4f}"
    )

    hasil_similarity.append({
        "kasus": no_kasus,
        "id_metode": id_metode,
        "metode": nama_metode,
        "kondisi_cocok": sorted(kondisi_cocok),
        "similarity": similarity
    })

print("\nJumlah kasus dihitung:", len(hasil_similarity))


# =====================================
# 9. RANKING KASUS
# =====================================

print("\n===== RANKING KASUS =====")

ranking = sorted(
    hasil_similarity,
    key=lambda data: data["similarity"],
    reverse=True
)

for nomor, data in enumerate(ranking, start=1):
    print(
        f"{nomor}. {data['kasus']} | "
        f"{data['metode']} | "
        f"Similarity: {data['similarity']:.2%}"
    )

# Menampilkan daftar ranking seperti contoh dosen
print("\nRanking Kasus:",
      [data["kasus"] for data in ranking])

print("\nRanking Metode:",
      [data["metode"] for data in ranking])

print("\nRanking Hasil:",
      [round(data["similarity"], 4) for data in ranking])


# =====================================
# 10. HASIL REKOMENDASI
# =====================================

nilai_tertinggi = ranking[0]["similarity"]

kasus_terbaik = [
    data for data in ranking
    if abs(data["similarity"] - nilai_tertinggi) < 1e-9
]

print("\n" + "=" * 50)
print(" HASIL REKOMENDASI METODE BELAJAR")
print("=" * 50)

print(f"Similarity Tertinggi: {nilai_tertinggi:.2%}")
print(f"Jumlah Kasus Terbaik: {len(kasus_terbaik)}")

if nilai_tertinggi == 0:
    print("\nTidak ditemukan kondisi yang cocok.")
    print("Rekomendasi belum dapat ditentukan.")

else:
    print("\nKasus dengan similarity tertinggi:")

    for data in kasus_terbaik:
        print(
            f"- {data['kasus']}: "
            f"{data['metode']} "
            f"({data['similarity']:.2%})"
        )

    rekomendasi = list(dict.fromkeys(
        data["metode"] for data in kasus_terbaik
    ))

    print("\nMetode belajar yang direkomendasikan:")

    for metode in rekomendasi:
        print(f"- {metode}")

    print("\nKondisi yang cocok pada kasus terbaik:")

    for data in kasus_terbaik:
        print(
            f"{data['kasus']}: "
            f"{', '.join(data['kondisi_cocok'])}"
        )

    # Catatan jika tingkat kemiripan rendah
    if nilai_tertinggi < 0.5:
        print("\nCATATAN:")
        print("Tingkat kemiripan kasus masih rendah.")
        print("Rekomendasi perlu dipertimbangkan kembali.")

print("\nProgram selesai dijalankan.")
