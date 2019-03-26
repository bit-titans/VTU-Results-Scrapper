import xlsxwriter
work_book = xlsxwriter.Workbook("Results.xlsx")
worksheet=work_book.add_worksheet("Results")
merge_format = work_book.add_format({
    'align': 'center'})

def write_to_excel(result):
    worksheet.write(0, 0, "Student USN")
    worksheet.write(0, 1, "Student Name")
    worksheet.merge_range('C1:F1', '15CS71', merge_format)
    worksheet.merge_range('G1:J1', '15CS72', merge_format)
    worksheet.merge_range('K1:N1', '15CS73', merge_format)
    worksheet.merge_range('O1:R1', '15CS754', merge_format)
    worksheet.merge_range('S1:V1', '15CS744', merge_format)
    worksheet.merge_range('W1:Z1', '15CSL76', merge_format)
    worksheet.merge_range('AA1:AD1', '15CSL77', merge_format)
    worksheet.merge_range('AE1:AH1', '15CSP78', merge_format)
    j = 2
    for i in range(0, 8):
        worksheet.write(1, j, "IA")
        worksheet.write(1, j + 1, "EA")
        worksheet.write(1, j + 2, "Total")
        worksheet.write(1, j + 3, "Result")
        j += 4
    row = 2
    for obj in result:
        worksheet.write(row, 0, obj["USN"])
        worksheet.write(row, 1, obj["Name"])
        worksheet.write(row, 2, obj["15CS71"][0])
        worksheet.write(row, 3, obj["15CS71"][1])
        worksheet.write(row, 4, obj["15CS71"][2])
        worksheet.write(row, 5, obj["15CS71"][3])
        worksheet.write(row, 6, obj["15CS72"][0])
        worksheet.write(row, 7, obj["15CS72"][1])
        worksheet.write(row, 8, obj["15CS72"][2])
        worksheet.write(row, 9, obj["15CS72"][3])
        worksheet.write(row, 10, obj["15CS73"][0])
        worksheet.write(row, 11, obj["15CS73"][1])
        worksheet.write(row, 12, obj["15CS73"][2])
        worksheet.write(row, 13, obj["15CS73"][3])
        worksheet.write(row, 14, obj["15CS754"][0])
        worksheet.write(row, 15, obj["15CS754"][1])
        worksheet.write(row, 16, obj["15CS754"][2])
        worksheet.write(row, 17, obj["15CS754"][3])
        worksheet.write(row, 18, obj["15CS744"][0])
        worksheet.write(row, 19, obj["15CS744"][1])
        worksheet.write(row, 20, obj["15CS744"][2])
        worksheet.write(row, 21, obj["15CS744"][3])
        worksheet.write(row, 22, obj["15CSL76"][0])
        worksheet.write(row, 23, obj["15CSL76"][1])
        worksheet.write(row, 24, obj["15CSL76"][2])
        worksheet.write(row, 25, obj["15CSL76"][3])
        worksheet.write(row, 26, obj["15CSL77"][0])
        worksheet.write(row, 27, obj["15CSL77"][1])
        worksheet.write(row, 28, obj["15CSL77"][2])
        worksheet.write(row, 29, obj["15CSL77"][3])
        worksheet.write(row, 30, obj["15CSP78"][0])
        worksheet.write(row, 31, obj["15CSP78"][1])
        worksheet.write(row, 32, obj["15CSP78"][2])
        worksheet.write(row, 33, obj["15CSP78"][3])
        row += 1
    return


def close():
    work_book.close()
