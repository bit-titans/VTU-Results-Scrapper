import xlsxwriter
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 01/02/2019
 """



work_book = xlsxwriter.Workbook("Results.xlsx")
work_sheet_internal = work_book.add_worksheet("Internal Marks")
work_sheet_external = work_book.add_worksheet("External Marks")
work_sheet_total = work_book.add_worksheet("Total Marks")
work_sheet_result = work_book.add_worksheet("Result")


def write_to_excel_internal(usn, name, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8):
    work_sheet_internal.write(0, 0, "Student USN")
    work_sheet_internal.write(0, 1, "Student Name")
    work_sheet_internal.write(0, 2, "Engineering Mathematics IV")
    work_sheet_internal.write(0, 3, "Software Engineering")
    work_sheet_internal.write(0, 4, " Design and Analysis of Algorithms")
    work_sheet_internal.write(0, 5, "Microprocessors and Microcontrollers")
    work_sheet_internal.write(0, 6, "Object Oriented Programming with JAVA")
    work_sheet_internal.write(0, 7, "Data Communications")
    work_sheet_internal.write(0, 8, "Design and Analysis of Algorithms Laboratory")
    work_sheet_internal.write(0, 9, "Microprocessors Laboratory")
    student = 0
    for i in range(1, 218):
        work_sheet_internal.write(i, 0, usn[student])
        work_sheet_internal.write(i, 1, name[student])
        work_sheet_internal.write(i, 2, ia1[student])
        work_sheet_internal.write(i, 3, ia2[student])
        work_sheet_internal.write(i, 4, ia3[student])
        work_sheet_internal.write(i, 5, ia4[student])
        work_sheet_internal.write(i, 6, ia5[student])
        work_sheet_internal.write(i, 7, ia6[student])
        work_sheet_internal.write(i, 8, ia7[student])
        work_sheet_internal.write(i, 9, ia8[student])
        student += 1
    return


def write_to_excel_external(usn, name, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8):
    work_sheet_external.write(0, 0, "Student USN")
    work_sheet_external.write(0, 1, "Student Name")
    work_sheet_external.write(0, 2, "Engineering Mathematics IV")
    work_sheet_external.write(0, 3, "Software Engineering")
    work_sheet_external.write(0, 4, " Design and Analysis of Algorithms")
    work_sheet_external.write(0, 5, "Microprocessors and Microcontrollers")
    work_sheet_external.write(0, 6, "Object Oriented Programming with JAVA")
    work_sheet_external.write(0, 7, "Data Communications")
    work_sheet_external.write(0, 8, "Design and Analysis of Algorithms Laboratory")
    work_sheet_external.write(0, 9, "Microprocessors Laboratory")
    student = 0
    for row in range(1, 218):
        work_sheet_external.write(row, 0, usn[student])
        work_sheet_external.write(row, 1, name[student])
        work_sheet_external.write(row, 2, ea1[student])
        work_sheet_external.write(row, 3, ea2[student])
        work_sheet_external.write(row, 4, ea3[student])
        work_sheet_external.write(row, 5, ea4[student])
        work_sheet_external.write(row, 6, ea5[student])
        work_sheet_external.write(row, 7, ea6[student])
        work_sheet_external.write(row, 8, ea7[student])
        work_sheet_external.write(row, 9, ea8[student])
        student += 1
    return


def write_to_excel_total(usn, name, total1, total2, total3, total4, total5, total6, total7, total8):
    work_sheet_total.write(0, 0, "Student USN")
    work_sheet_total.write(0, 1, "Student Name")
    work_sheet_total.write(0, 2, "Engineering Mathematics IV")
    work_sheet_total.write(0, 3, "Software Engineering")
    work_sheet_total.write(0, 4, " Design and Analysis of Algorithms")
    work_sheet_total.write(0, 5, "Microprocessors and Microcontrollers")
    work_sheet_total.write(0, 6, "Object Oriented Programming with JAVA")
    work_sheet_total.write(0, 7, "Data Communications")
    work_sheet_total.write(0, 8, "Design and Analysis of Algorithms Laboratory")
    work_sheet_total.write(0, 9, "Microprocessors Laboratory")
    student = 0
    for row in range(1, 218):
        work_sheet_total.write(row, 0, usn[student])
        work_sheet_total.write(row, 1, name[student])
        work_sheet_total.write(row, 2, total1[student])
        work_sheet_total.write(row, 3, total2[student])
        work_sheet_total.write(row, 4, total3[student])
        work_sheet_total.write(row, 5, total4[student])
        work_sheet_total.write(row, 6, total5[student])
        work_sheet_total.write(row, 7, total6[student])
        work_sheet_total.write(row, 8, total7[student])
        work_sheet_total.write(row, 9, total8[student])
        student += 1
    return


def write_to_excel_result(usn, name, r1, r2, r3, r4, r5, r6, r7, r8):
    work_sheet_result.write(0, 0, "Student USN")
    work_sheet_result.write(0, 1, "Student Name")
    work_sheet_result.write(0, 2, "Engineering Mathematics IV")
    work_sheet_result.write(0, 3, "Software Engineering")
    work_sheet_result.write(0, 4, "Design and Analysis of Algorithms")
    work_sheet_result.write(0, 5, "Microprocessors and Microcontrollers")
    work_sheet_result.write(0, 6, "Object Oriented Programming with JAVA")
    work_sheet_result.write(0, 7, "Data Communications")
    work_sheet_result.write(0, 8, "Design and Analysis of Algorithms Laboratory")
    work_sheet_result.write(0, 9, "Microprocessors Laboratory")
    student = 0
    for row in range(1, 218):
        work_sheet_result.write(row, 0, usn[student])
        work_sheet_result.write(row, 1, name[student])
        work_sheet_result.write(row, 2, r1[student])
        work_sheet_result.write(row, 3, r2[student])
        work_sheet_result.write(row, 4, r3[student])
        work_sheet_result.write(row, 5, r4[student])
        work_sheet_result.write(row, 6, r5[student])
        work_sheet_result.write(row, 7, r6[student])
        work_sheet_result.write(row, 8, r7[student])
        work_sheet_result.write(row, 9, r8[student])
        student += 1
    return


def close():
    work_book.close()
