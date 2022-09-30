#-------------------------------------------------------------------------------
# Name:        ExcelUpdate
# Purpose:
#
# Author:      Gürol Güngör
#
# Created:     25.09.2022
# Copyright:   (c) Gürol Güngör 2022
# Licence:     GNU - General Public Licence
#-------------------------------------------------------------------------------

#Excel kitapligini yukluyoruz
import openpyxl
import os
import sys

#Excel dosyası yeri ve adi tanımlanır
ExcelDosyasi = r"C:\Database\rehber.xlsx"

#Excel dosyası cağırılır
workbook = openpyxl.load_workbook(ExcelDosyasi)

#Excel Sayfası açılır
sheet = workbook.active

#A kolonu 1 satır güncellenir
sheet["A1"].value = "Niyazi"


#Excel dosyası kayıtedilir
workbook.save(ExcelDosyasi)

# workbook cikmadan once kapatiyoruz.
workbook.close()

# Program kapatılır.
sys.exit()


