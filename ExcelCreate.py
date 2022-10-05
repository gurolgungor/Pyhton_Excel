#-------------------------------------------------------------------------------
# Name:        ExcelCreate
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
import sys

#Excel dosyası yeri ve adi tanımlanır
ExcelDosyasi = r"C:\Database\rehber.xlsx"

# workbook ve worksheet olusturyoruz
workbook = xlsxwriter.Workbook(ExcelDosyasi)
worksheet = workbook.add_worksheet()

#Ekrana bos bir Excel dosyası olusturulduğu yazılır
print(ExcelDosyasi+" isminde bos bir Excel olusturuldu.")

# workbook cikmadan once kapatiyoruz.
workbook.close()

# Program kapatılır.
sys.exit()

