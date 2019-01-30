import xlsxwriter
work_book = xlsxwriter.Workbook("Results.xlsx")
work_sheet_internal = work_book.add_worksheet("Internal Marks")
work_sheet_external = work_book.add_worksheet("External Marks")
work_sheet_total = work_book.add_worksheet("Total Marks")


def write_to_excel_internal(usn, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8):
    work_sheet_internal.write(0, 0, "Student USN")
    work_sheet_internal.write(0, 1, "Engineering Mathematics IV")
    work_sheet_internal.write(0, 2, "Software Engineering")
    work_sheet_internal.write(0, 3, " Design and Analysis of Algorithms")
    work_sheet_internal.write(0, 4, "Microprocessors and Microcontrollers")
    work_sheet_internal.write(0, 5, "Object Oriented Programming with JAVA")
    work_sheet_internal.write(0, 6, "Data Communications")
    work_sheet_internal.write(0, 7, "Design and Analysis of Algorithms Laboratory")
    work_sheet_internal.write(0, 8, "Microprocessors Laboratory")
    row = 1
    student = 0
    for i in range(1, 218):
        work_sheet_internal.write(row, 0, usn)
        work_sheet_internal.write(row, 1, ia1[student])
        work_sheet_internal.write(row, 2, ia2[student])
        work_sheet_internal.write(row, 3, ia3[student])
        work_sheet_internal.write(row, 4, ia4[student])
        work_sheet_internal.write(row, 5, ia5[student])
        work_sheet_internal.write(row, 6, ia6[student])
        work_sheet_internal.write(row, 7, ia7[student])
        work_sheet_internal.write(row, 8, ia8[student])
        student += 1
    return


def write_to_excel_external(usn, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8):
    work_sheet_external.write(0, 0, "Student USN")
    work_sheet_external.write(0, 1, "Engineering Mathematics IV")
    work_sheet_external.write(0, 2, "Software Engineering")
    work_sheet_external.write(0, 3, " Design and Analysis of Algorithms")
    work_sheet_external.write(0, 4, "Microprocessors and Microcontrollers")
    work_sheet_external.write(0, 5, "Object Oriented Programming with JAVA")
    work_sheet_external.write(0, 6, "Data Communications")
    work_sheet_external.write(0, 7, "Design and Analysis of Algorithms Laboratory")
    work_sheet_external.write(0, 8, "Microprocessors Laboratory")
    row = 1
    student = 0
    for i in range(1, 218):
        work_sheet_external.write(row, 0, usn)
        work_sheet_external.write(row, 1, ea1[student])
        work_sheet_external.write(row, 2, ea2[student])
        work_sheet_external.write(row, 3, ea3[student])
        work_sheet_external.write(row, 4, ea4[student])
        work_sheet_external.write(row, 5, ea5[student])
        work_sheet_external.write(row, 6, ea6[student])
        work_sheet_external.write(row, 7, ea7[student])
        work_sheet_external.write(row, 8, ea8[student])
        student += 1
    return


def write_to_excel_total(usn, total1, total2, total3, total4, total5, total6, total7, total8):
    work_sheet_total.write(0, 0, "Student USN")
    work_sheet_total.write(0, 1, "Engineering Mathematics IV")
    work_sheet_total.write(0, 2, "Software Engineering")
    work_sheet_total.write(0, 3, " Design and Analysis of Algorithms")
    work_sheet_total.write(0, 4, "Microprocessors and Microcontrollers")
    work_sheet_total.write(0, 5, "Object Oriented Programming with JAVA")
    work_sheet_total.write(0, 6, "Data Communications")
    work_sheet_total.write(0, 7, "Design and Analysis of Algorithms Laboratory")
    work_sheet_total.write(0, 8, "Microprocessors Laboratory")
    row = 1
    student = 0
    for i in range(1, 218):
        work_sheet_total.write(row, 0, usn)
        work_sheet_total.write(row, 1, total1[student])
        work_sheet_total.write(row, 2, total2[student])
        work_sheet_total.write(row, 3, total3[student])
        work_sheet_total.write(row, 4, total4[student])
        work_sheet_total.write(row, 5, total5[student])
        work_sheet_total.write(row, 6, total6[student])
        work_sheet_total.write(row, 7, total7[student])
        work_sheet_total.write(row, 8, total8[student])
        student += 1
    return

