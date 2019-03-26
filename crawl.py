import requests
from lxml import html
import ocr
import Write
import WriteDB
import xlrd
i = 1
gpa = 0
name = []
names = []
result = []
lsubs = []
loc = ("USN.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
while True:
    try:
        USN = sheet.cell_value(i, 0)
        batch = USN[3:5]
    except:
        break
    gpa = 0
    s = requests.Session()
    headers = {'Referer': 'http://results.vtu.ac.in/vitaviresultcbcs2018/index.php',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
                   'Upgrade-Insecure-Requests': '1',  'Cookie': 'PHPSESSID=u47uot7eg9j6eglqm951e3nfr7'
            , 'Connection': 'keep-alive'}
    image = s.get("http://results.vtu.ac.in/resultsvitavicbcs_19/captcha_new.php", headers=headers)
    with open("snap.png", 'wb') as file:
        file.write(image.content)
    cap = ocr.get_ocr("snap.png")
    #USN = "1BI17CS"+str(format(i, '03d'))
    url = "http://results.vtu.ac.in/resultsvitavicbcs_19/resultpage.php"
    payload = {'lns': USN, 'captchacode': str(cap),
                   'token': 'bEpYczFha3YrYUZMbjhZSDRydkgvZ1kyS3puN09PL09YcWd4bm1lUjBaa0p1eUZRdWFtbzNZamNVeGtnbGtTeHdBNWFydTVyUXJQU0lZYmUxQkRCTHc9PTo6Ydj2a8akPwQGJ9ITlWo7pw',
                   'current_url': 'http://results.vtu.ac.in/resultsvitavicbcs_19/index.php'}
    page = s.post(url, data=payload, headers=headers)
    tree = html.fromstring(page.content)
    print("Sent USN:-"+USN)
   # print("Sent USN:-1BI17CS"+str(format(i, '03d')))
    print("Sent Captcha:"+ocr.get_ocr("snap.png"))
    if "Invalid captcha code !!!" in page.text:
        print("Invalid captcha code !!!")
        continue
    else:
        i += 1
    if "Redirecting to VTU Results Site" in page.text:
        print("Alert:-Token Expired!!:Update new token in Payload")
        exit(2)
    if "University Seat Number is not available or Invalid..!" in page.text:
        print("University Seat Number is not available or Invalid..!")
        exit(-1)
    temp = page.text.find("Student Name")
    name.clear()
    while (page.text[temp + 61] != "<"):
        name.insert(len(name), page.text[temp + 61])
        temp += 1
    names.insert(len(names), ''.join(name))
    if("Semester : 5" in page.text):
        iresult = {
            "USN": "-",
            "Name": "-",
            "15CS51": ["-", "-", "-", "-"],
            "15CS52": ["-", "-", "-", "-"],
            "15CS53": ["-", "-", "-", "-"],
            "15CS54": ["-", "-", "-", "-"],
            "15CSL57": ["-", "-", "-", "-"],
            "15CSL58": ["-", "-", "-", "-"],
            "15ME562": ["-", "-", "-", "-"],
            "15CS564": ["-", "-", "-", "-"],
            "15CS562": ["-", "-", "-", "-"],
            "15PHY561": ["-", "-", "-", "-"],
            "15CS553": ["-", "-", "-", "-"],
        }
        imarks1 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]')[0].text
        imarks2 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[3]')[0].text
        imarks3 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[3]')[0].text
        imarks4 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[3]')[0].text
        imarks5 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[3]')[0].text
        imarks6 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[3]')[0].text
        imarks7 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[3]')[0].text
        imarks8 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[3]')[0].text
        emarks1 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[4]')[0].text
        emarks2 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[4]')[0].text
        emarks3 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[4]')[0].text
        emarks4 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[4]')[0].text
        emarks5 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[4]')[0].text
        emarks6 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[4]')[0].text
        emarks7 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[4]')[0].text
        emarks8 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[4]')[0].text
        tmarks1 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[5]')[0].text
        tmarks2 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[5]')[0].text
        tmarks3 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[5]')[0].text
        tmarks4 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[5]')[0].text
        tmarks5 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[5]')[0].text
        tmarks6 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[5]')[0].text
        tmarks7 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[5]')[0].text
        tmarks8 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[5]')[0].text
        result1 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[6]')[0].text
        result2 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[6]')[0].text
        result3 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[6]')[0].text
        result4 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[6]')[0].text
        result5 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[6]')[0].text
        result6 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[6]')[0].text
        result7 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[6]')[0].text
        result8 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[6]')[0].text
        sub1 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[1]')[0].text
        sub2 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[1]')[0].text
        sub3 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[1]')[0].text
        sub4 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[1]')[0].text
        sub5 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[1]')[0].text
        sub6 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[1]')[0].text
        sub7 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[1]')[0].text
        sub8 = tree.xpath('/html/body/div[2]/form/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[1]')[0].text
        iresult["USN"] = USN
        iresult["Name"] = ''.join(name)
        iresult[sub1] = [imarks1, emarks1, tmarks1, result1]
        iresult[sub2] = [imarks2, emarks2, tmarks2, result2]
        iresult[sub3] = [imarks3, emarks3, tmarks3, result3]
        iresult[sub4] = [imarks4, emarks4, tmarks4, result4]
        iresult[sub5] = [imarks5, emarks5, tmarks5, result5]
        iresult[sub6] = [imarks6, emarks6, tmarks6, result6]
        iresult[sub7] = [imarks7, emarks7, tmarks7, result7]
        iresult[sub8] = [imarks8, emarks8, tmarks8, result8]
        result.insert(len(result), iresult)
    else:
        continue
done = 0
while done == 0:
    print("WRITE:-")
    print("Menu:-\n 1)Write to Spreadsheet\n 2)Write to Database\n 3)Both\n 4)Exit")
    choice = int(input("Enter your choice:-"))
    if choice == 1:
            Write.write_to_excel(result)
            print("Write Finished")
    elif choice == 2:
        try:
            WriteDB.drop_tables()
        except:
            pass
        WriteDB.create_tables()
        WriteDB.write_internal(usn_list, names, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8)
        WriteDB.write_external(usn_list, names, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8)
        WriteDB.write_total(usn_list, names, t1, t2, t3, t4, t5, t6, t7, t8)
        WriteDB.write_results(usn_list, names, r1, r2, r3, r4, r5, r6, r7, r8)
        print("Write Finished")
    elif choice == 3:
        Write.write_to_excel_internal(usn_list, names, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8)
        Write.write_to_excel_external(usn_list, names, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8)
        Write.write_to_excel_total(usn_list, names, t1, t2, t3, t4, t5, t6, t7, t8)
        Write.write_to_excel_result(usn_list, names, r1, r2, r3, r4, r5, r6, r7, r8)
        try:
            WriteDB.drop_tables()
        except:
            pass
        WriteDB.create_tables()
        WriteDB.write_internal(usn_list, names, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8)
        WriteDB.write_external(usn_list, names, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8)
        WriteDB.write_total(usn_list, names, t1, t2, t3, t4, t5, t6, t7, t8)
        WriteDB.write_results(usn_list, names, r1, r2, r3, r4, r5, r6, r7, r8)
        print("Write Finished")
    elif choice == 4:
        done = 1
        Write.close()
        WriteDB.end_connection()
        exit(0)
    else:
        print("Invalid Choice")
