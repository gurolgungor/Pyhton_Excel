#-------------------------------------------------------------------------------
# Name:        Excelinsert
# Purpose:
#
# Author:      Gürol Güngör
#
# Created:     25.09.2022
# Copyright:   (c) Gürol Güngör 2022
# Licence:     GNU - General Public Licence
#-------------------------------------------------------------------------------

#Excel kitapligini yukluyoruz
import xlsxwriter

#Excel dosyası yeri ve adi tanımlanır
ExcelDosyasi = r"C:\Database\rehber.xlsx"

# workbook ve worksheet olusturyoruz
workbook = xlsxwriter.Workbook(ExcelDosyasi)
worksheet = workbook.add_worksheet()

# Excel kaydedilecek kayitlari olusturyoruz.
tum_kayitlar = [
    ('Ahmet','Güngör'),
    ('Mehmet','Güngör'),
    ('Nihat','Güngör'),
]

# Excel icerisinde hagi satir,sutundan baslanacagini seciyoruz.
row = 0
col = 0

# kayitlar sayfaya yaziliyor.
for Ad, Soyad in (tum_kayitlar):
    worksheet.write(row, col,     Ad)
    worksheet.write(row, col + 1, Soyad)
    row += 1


# workbook cikmadan once kapatiyoruz.
workbook.close()

# Program kapatılır.
sys.exit()


