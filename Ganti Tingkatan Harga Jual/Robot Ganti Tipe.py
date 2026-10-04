import time
import pyautogui

pyautogui.FAILSAFE = True

pyautogui.PAUSE = 0.4


def cari_dan_klik(
    nama_gambar, offset_x=0, offset_y=0, confidence=0.8, timeout=5
):
    
    waktu_mulai = time.time()
    while time.time() - waktu_mulai < timeout:
        try:
            lokasi = pyautogui.locateCenterOnScreen(
                nama_gambar, confidence=confidence
            )
            if lokasi is not None:
                pyautogui.click(lokasi.x + offset_x, lokasi.y + offset_y)
                return True
        except pyautogui.ImageNotFoundException:
            pass
        time.sleep(0.3)

    print(f"[WARNING] Gambar '{nama_gambar}' tidak ditemukan di layar!")
    return False


def jalankan_otomatisasi_accurate(jumlah_perulangan=10):

    print("=========================================================")
    print("   OTOMATISASI ACCURATE (BERBASIS PENCARIAN GAMBAR)     ")
    print("=========================================================")
    print("Persiapan: Buka dan fokuskan layar ke aplikasi Accurate!")
    
    for detik in range(5, 0, -1):
        print(f"Otomatisasi dimulai dalam {detik} detik...")
        time.sleep(1)

    print("[INFO] Proses otomatisasi berjalan...")

    for i in range(1, jumlah_perulangan + 1):
        print(f"--- Memproses Data Ke-{i} dari {jumlah_perulangan} ---")
        
        print("1. Menekan Enter untuk membuka detail data...")
        pyautogui.press("enter")
        time.sleep(1.0)
        
        print("2. Mencari & mengklik tab 'Penjualan'...")
        if not cari_dan_klik("penjualan.png", confidence=0.8, timeout=5):
            print("[GAGAL] Tab Penjualan tidak ditemukan. Lewati baris ini.")
            continue
        time.sleep(0.5)
        
        print("3. Mencari 'Tingkatan Harga jual' & mengklik kolom pilihan...")
        if not cari_dan_klik(
            "tingkatan_harga_jual.png",
            offset_x=180,
            offset_y=0,
            confidence=0.8,
            timeout=5,
        ):
            print(
                "[GAGAL] Kolom Tingkatan Harga jual tidak ditemukan. Lewati..."
            )
            continue
        time.sleep(0.3)
        
        print("4. Menekan panah arah bawah 5 kali...")
        for _ in range(5):
            pyautogui.press("down")
            time.sleep(0.03)
        time.sleep(0.3)
        
        print("5. Mencari & mengklik tombol 'OK'...")
        if not cari_dan_klik("ok.png", confidence=0.8, timeout=0.3):
            print("[GAGAL] Tombol OK tidak ditemukan.")
            continue

        if not cari_dan_klik("ok.png", confidence=0.8, timeout=0.3):
            print("[GAGAL] Tombol OK tidak ditemukan.")
            continue
        
        time.sleep(0.5)
        
        print("6. Mengarahkan panah bawah 1x untuk data berikutnya...")
        pyautogui.press("down")
        time.sleep(0.3)

        print(f"[INFO] Data ke-{i} selesai diperbarui!")

    print("=========================================================")
    print("   SELESAI: Seluruh proses otomatisasi telah berhasil!   ")
    print("=========================================================")


if __name__ == "__main__":
    JUMLAH_DATA = 100
    jalankan_otomatisasi_accurate(jumlah_perulangan=JUMLAH_DATA)