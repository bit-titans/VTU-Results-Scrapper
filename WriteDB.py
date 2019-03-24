import sqlite3
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 02/02/2019
 """
connection = sqlite3.connect("Results.db")
c = connection.cursor()
internal_create = '''
CREATE TABLE internal_marks (
USN text PRIMARY KEY,
Student_Name text,
Engineering_Mathematics_IV text,
Software_Engineering text,
Design_and_Analysis_of_Algorithms text,
Microprocessors_and_Microcontrollers text,
Object_Oriented_Programming_with_JAVA text,
Data_Communications text,
Design_and_Analysis_of_Algorithms_Laboratory text,
Microprocessors_Laboratory text);'''

external_create = '''
CREATE TABLE external_marks(
USN text PRIMARY KEY,
Student_Name text,
Engineering_Mathematics_IV text,
Software_Engineering text,
Design_and_Analysis_of_Algorithms text,
Microprocessors_and_Microcontrollers text,
Object_Oriented_Programming_with_JAVA text,
Data_Communications text,
Design_and_Analysis_of_Algorithms_Laboratory text,
Microprocessors_Laboratory text);'''

total_create = '''
CREATE TABLE total_marks (
USN text PRIMARY KEY,
Student_Name text,
Engineering_Mathematics_IV text,
Software_Engineering text,
Design_and_Analysis_of_Algorithms text,
Microprocessors_and_Microcontrollers text,
Object_Oriented_Programming_with_JAVA text,
Data_Communications text,
Design_and_Analysis_of_Algorithms_Laboratory text,
Microprocessors_Laboratory text);'''

results_create = '''
CREATE TABLE result (
USN text PRIMARY KEY,
Student_Name text,
Engineering_Mathematics_IV text,
Software_Engineering text,
Design_and_Analysis_of_Algorithms text,
Microprocessors_and_Microcontrollers text,
Object_Oriented_Programming_with_JAVA text,
Data_Communications text,
Design_and_Analysis_of_Algorithms_Laboratory text,
Microprocessors_Laboratory text);'''


def create_tables():
    c.execute(internal_create)
    c.execute(external_create)
    c.execute(total_create)
    c.execute(results_create)


def end_connection():
    connection.commit()
    connection.close()


def drop_tables():
    c.execute("DROP TABLE internal_marks;")
    c.execute("DROP TABLE external_marks;")
    c.execute("DROP TABLE total_marks;")
    c.execute("DROP TABLE result;")


def write_internal(usn, name, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8):
    for i in range(0, len(usn)):
        c.execute("INSERT INTO internal_marks values(?,?,?,?,?,?,?,?,?,?);", (usn[i], name[i], ia1[i], ia2[i], ia3[i], ia4[i], ia5[i], ia6[i], ia7[i], ia8[i]))
    return


def write_external(usn, name, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8):
    for i in range(0, len(usn)):
        c.execute("INSERT INTO external_marks values(?,?,?,?,?,?,?,?,?,?);", (usn[i], name[i], ea1[i], ea2[i], ea3[i], ea4[i], ea5[i], ea6[i], ea7[i], ea8[i]))
    return


def write_total(usn, name, total1, total2, total3, total4, total5, total6, total7, total8):
    for i in range(0, len(usn)):
        c.execute("INSERT INTO total_marks values(?,?,?,?,?,?,?,?,?,?);", (usn[i], name[i], total1[i], total2[i], total3[i], total4[i], total5[i], total6[i], total7[i], total8[i]))
    return


def write_results(usn, name, r1, r2, r3, r4, r5, r6, r7, r8):
    for i in range(0, len(usn)):
        c.execute("INSERT INTO result values(?,?,?,?,?,?,?,?,?,?);", (usn[i], name[i], r1[i], r2[i], r3[i], r4[i], r5[i], r6[i], r7[i], r8[i]))
    return

