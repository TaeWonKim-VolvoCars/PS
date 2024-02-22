import os
import openpyxl as op
from openpyxl import Workbook
from datetime import datetime, timedelta

workbook = Workbook()
active_sheet = workbook.active
active_sheet.title = "find supplier invoices"

new_sheet = workbook.create_sheet("payable aging")
workbook.active = new_sheet

current_time = datetime.now()
file_location = "C:/Users/Public/"
file_name = f"{file_location}kr21_ptp_aging_{current_time.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
workbook.save(file_name)
workbook.close()
