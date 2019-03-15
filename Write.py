import xlsxwriter
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 01/02/2019
 """



work_book = xlsxwriter.Workbook("Results.xlsx")
worksheet=work_book.add_worksheet("Results")
merge_format = work_book.add_format({
    'align': 'center'})

def write_to_excel(usn, name, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8, total1, total2, total3, total4, total5, total6, total7, total8, r1, r2, r3, r4, r5, r6, r7, r8):
    worksheet.write(0, 0, "Student USN")
    worksheet.write(0, 1, "Student Name")
    worksheet.merge_range('C1:F1', 'WEB TECHNOLOGY AND ITS APPLICATIONS', merge_format)
    worksheet.merge_range('G1:J1', 'ADVANCED COMPUTER ARCHITECTURES', merge_format)
    worksheet.merge_range('K1:N1', 'MACHINE LEARNING', merge_format)
    worksheet.merge_range('O1:R1', 'ELECTIVE 1', merge_format)
    worksheet.merge_range('S1:V1', 'ELECTIVE 2', merge_format)
    worksheet.merge_range('W1:Z1', 'MACHINE LEARNING LABORATORY', merge_format)
    worksheet.merge_range('AA1:AD1', 'WEB TECHNOLOGY LABORATORY WITH MINI PROJECT', merge_format)
    worksheet.merge_range('AE1:AH1', 'PROJECT PHASE 1 + SEMINAR', merge_format)
    j = 2
    for i in range(0, 8):
        worksheet.write(1, j, "IA")
        worksheet.write(1, j + 1, "EA")
        worksheet.write(1, j + 2, "Total")
        worksheet.write(1, j + 3, "Result")
        j += 4
    length = len(ia1)
    row = 2
    for i in range(0,length):
        worksheet.write(row, 0, usn[i])
        worksheet.write(row, 1, name[i])
        worksheet.write(row, 2, ia1[i])
        worksheet.write(row, 3, ea1[i])
        worksheet.write(row, 4, total1[i])
        worksheet.write(row, 5, r1[i])
        worksheet.write(row, 6, ia2[i])
        worksheet.write(row, 7, ea2[i])
        worksheet.write(row, 8, total2[i])
        worksheet.write(row, 9, r2[i])
        worksheet.write(row, 10, ia3[i])
        worksheet.write(row, 11, ea3[i])
        worksheet.write(row, 12, total3[i])
        worksheet.write(row, 13, r3[i])
        worksheet.write(row, 14, ia4[i])
        worksheet.write(row, 15, ea4[i])
        worksheet.write(row, 16, total4[i])
        worksheet.write(row, 17, r4[i])
        worksheet.write(row, 18, ia5[i])
        worksheet.write(row, 19, ea5[i])
        worksheet.write(row, 20, total5[i])
        worksheet.write(row, 21, r5[i])
        worksheet.write(row, 22, ia6[i])
        worksheet.write(row, 23, ea6[i])
        worksheet.write(row, 24, total6[i])
        worksheet.write(row, 25, r6[i])
        worksheet.write(row, 26, ia7[i])
        worksheet.write(row, 27, ea7[i])
        worksheet.write(row, 28, total7[i])
        worksheet.write(row, 29, r7[i])
        worksheet.write(row, 30, ia8[i])
        worksheet.write(row, 31, ea8[i])
        worksheet.write(row, 32, total8[i])
        worksheet.write(row, 33, r8[i])
        row += 1
    return


def close():
    work_book.close()
