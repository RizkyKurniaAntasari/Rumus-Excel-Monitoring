# Rumus-Excel-Monitoring

Hal yang perlu dilakukan sebelum menggunakan rumus
1. download data Kegiatan_Assurance_Berjenjang_2025_RO_Bandar_Lampung_Responses.xlsx
2. dan copy semua data menggunakan
--------------------------------------------
   CTRL + A dan dilanjut dengan CTRL + C
--------------------------------------------
3. Klik 2x CELL E6.

<img width="1886" height="458" alt="image" src="https://github.com/user-attachments/assets/dda089f3-8483-400e-a74b-230ed19a283b" />

Copy Rumus berikut:
------------------------------------------------------------------------

=COUNTIFS(
  Sumber!$B:$B, ">="&DATE(2025,12,1),
  Sumber!$B:$B, "<="&DATE(2025,12,7),
  Sumber!$C:$C, "Supervisor (UKO)",
  Sumber!$E:$E, $B6,
  Sumber!$H:$H,
  "*" &
  IF(E$5="CS","Customer Service",
  IF(E$5="Teller","Teller",
  IF(E$5="Satpam","Satpam",
  IF(E$5="Banking Hall","Banking Hall",
  IF(E$5="Gallery ATM","Gallery E-Channel",
  IF(E$5="Fasad","Fasad Gedung",
  IF(E$5="Toilet","Toilet",
  IF(E$5="Brimen","Brimen, Ruang Kerja, Pantry, Ruang Server & Gudang",
  IF(E$5="SPV","SPV","")))))))))
  & "*"
)

------------------------------------------------------------------------
Letakan Rumus di E6

Terdapat 2 hal yang bisa di ubah yaitu: 
2 baris pertama dalam rumus
yaitu ->
-----------------------------
  Sumber!$B:$B, ">="&DATE(2025,12,1),
  Sumber!$B:$B, "<="&DATE(2025,12,7),
-----------------------------

artinya 2025 desember tanggal 1 sampai 
2025 desember tanggal 8
jadi data dari 1 desember sampai 8 desember akan di ambil.

  Sumber!$C:$C, "Supervisor (UKO)",
baris ke 3 ini bisa di ganti-ganti sesuai yang ada di sumber
seperti:
1. MOL / AMOL
2. Supervisor (UKO)
3. Regional Office
contoh saya ingin mengganti ke MOL / AMOL maka penulisannya
-----------------------------
  Sumber!$C:$C, "MOL / AMOL",
-----------------------------

==================================================================================
=============================== S T U D Y  C A S E ===============================
==================================================================================

Jadi sebagai contoh. saya mau data bulan
Januari tahun 2026 dan saya mau merekap data "MOL / AMOL", minggu pertama maka rumusnya menjadi

=COUNTIFS(
  Sumber!$B:$B, ">="&DATE(2026,1,1),
  Sumber!$B:$B, "<="&DATE(2026,1,7),
  Sumber!$C:$C, "MOL / AMOL",
  Sumber!$E:$E, $B6,
  Sumber!$H:$H,
  "*" &
  IF(E$5="CS","Customer Service",
  IF(E$5="Teller","Teller",
  IF(E$5="Satpam","Satpam",
  IF(E$5="Banking Hall","Banking Hall",
  IF(E$5="Gallery ATM","Gallery E-Channel",
  IF(E$5="Fasad","Fasad Gedung",
  IF(E$5="Toilet","Toilet",
  IF(E$5="Brimen","Brimen, Ruang Kerja, Pantry, Ruang Server & Gudang",
  IF(E$5="SPV","SPV","")))))))))
  & "*"
)

Terlihat pada 3 baris pertama yang berubah

yaitu
-----------------------------
  Sumber!$B:$B, ">="&DATE(2026,1,1),
  Sumber!$B:$B, "<="&DATE(2026,1,7),
  Sumber!$C:$C, "MOL / AMOL",
-----------------------------.
