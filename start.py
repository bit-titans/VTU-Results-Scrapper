import requests
from lxml import html
import ocr
import Write
import WriteDB
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 01/02/2019
 """

usn_list = []
ia1 = []
ia2 = []
ia3 = []
ia4 = []
ia5 = []
ia6 = []
ia7 = []
ia8 = []
ea1 = []
ea2 = []
ea3 = []
ea4 = []
ea5 = []
ea6 = []
ea7 = []
ea8 = []
t1 = []
t2 = []
t3 = []
t4 = []
t5 = []
t6 = []
t7 = []
t8 = []
r1 = []
r2 = []
r3 = []
r4 = []
r5 = []
r6 = []
r7 = []
r8 = []
i = 1
name = []
names = []
while i <= 199:
    s = requests.Session()
    headers = {'Referer': 'https://results.vtu.ac.in/vitavicbcsjj19/index.php',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
                   'Upgrade-Insecure-Requests': '1',  'Cookie': 'PHPSESSID=e0e6cf142e7c180354e2d78d0ec2eb63'
            , 'Connection': 'keep-alive'}
    image = s.get("https://results.vtu.ac.in/vitavicbcsjj19/captcha_new.php", headers=headers,verify=False)
    with open("snap.png", 'wb') as file:
        file.write(image.content)
    cap = ocr.get_ocr("snap.png")
    USN = "1BI18CS"+str(format(i, '03d'))
    url = "https://results.vtu.ac.in/vitavicbcsjj19/resultpage.php"
    payload = {'lns': USN, 'captchacode': str(cap),
                   'token': 'YUFMMlhUOVNLUE1RV09Wa1d5ZUh0bU93dUNqSDZmQzlnK3RXVUVnbWt6MjVQWG4xZlZtZ2FFNFdMZGh0MHQvV1pOR0kxMnRtTzc2bk1KczhhcUFLOUE9PTo6qSjR4GNZb1EArB/rUbCE+Q==',
                   'current_url': 'https://results.vtu.ac.in/vitavicbcsjj19/index.php'}
    page = s.post(url, data=payload, headers=headers,verify=False)
    print(page)
    tree = html.fromstring(page.content)
    print("Sent USN:-1BI18CS"+str(format(i, '03d')))
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
        ia1.insert(len(ia1), "-")
        ia2.insert(len(ia2), "-")
        ia3.insert(len(ia3), "-")
        ia4.insert(len(ia4), "-")
        ia5.insert(len(ia5), "-")
        ia6.insert(len(ia6), "-")
        ia7.insert(len(ia7), "-")
        ia8.insert(len(ia8), "-")
        ea1.insert(len(ea1), "-")
        ea2.insert(len(ea2), "-")
        ea3.insert(len(ea3), "-")
        ea4.insert(len(ea4), "-")
        ea5.insert(len(ea5), "-")
        ea6.insert(len(ea6), "-")
        ea7.insert(len(ea7), "-")
        ea8.insert(len(ea8), "-")
        t1.insert(len(t1), "-")
        t2.insert(len(t2), "-")
        t3.insert(len(t3), "-")
        t4.insert(len(t4), "-")
        t5.insert(len(t5), "-")
        t6.insert(len(t6), "-")
        t7.insert(len(t7), "-")
        t8.insert(len(t8), "-")
        r1.insert(len(r1), "-")
        r2.insert(len(r2), "-")
        r3.insert(len(r3), "-")
        r4.insert(len(r4), "-")
        r5.insert(len(r5), "-")
        r6.insert(len(r6), "-")
        r7.insert(len(r7), "-")
        r8.insert(len(r8), "-")
        names.insert(len(names), "USN does'nt Exist")
        continue
    temp = page.text.find("Student Name")
    # print(page.text)
    # print(page.text[temp])
    name.clear()

    while (page.text[temp + 82] != "<"):
        name.insert(len(name), page.text[temp + 82])
        temp += 1
    # print(name)
    # exit(3)
    names.insert(len(names), ''.join(name))
    if("Semester : 2" in page.text):
        imarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]')[0].text
        imarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[3]')[0].text
        imarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[3]')[0].text
        imarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[3]')[0].text
        imarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[3]')[0].text
        imarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[3]')[0].text
        imarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[3]')[0].text
        imarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[3]')[0].text
        emarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[4]')[0].text
        emarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[4]')[0].text
        emarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[4]')[0].text
        emarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[4]')[0].text
        emarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[4]')[0].text
        emarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[4]')[0].text
        emarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[4]')[0].text
        emarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[4]')[0].text
        tmarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[5]')[0].text
        tmarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[5]')[0].text
        tmarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[5]')[0].text
        tmarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[5]')[0].text
        tmarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[5]')[0].text
        tmarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[5]')[0].text
        tmarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[5]')[0].text
        tmarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[5]')[0].text
        result1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[2]/div[6]')[0].text
        result2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[6]')[0].text
        result3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[4]/div[6]')[0].text
        result4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[5]/div[6]')[0].text
        result5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[6]/div[6]')[0].text
        result6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[7]/div[6]')[0].text
        result7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[8]/div[6]')[0].text
        result8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div[9]/div[6]')[0].text
        ia1.insert(len(ia1), imarks1)
        ia2.insert(len(ia2), imarks2)
        ia3.insert(len(ia3), imarks3)
        ia4.insert(len(ia4), imarks4)
        ia5.insert(len(ia5), imarks5)
        ia6.insert(len(ia6), imarks6)
        ia7.insert(len(ia7), imarks7)
        ia8.insert(len(ia8), imarks8)
        ea1.insert(len(ea1), emarks1)
        ea2.insert(len(ea2), emarks2)
        ea3.insert(len(ea3), emarks3)
        ea4.insert(len(ea4), emarks4)
        ea5.insert(len(ea5), emarks5)
        ea6.insert(len(ea6), emarks6)
        ea7.insert(len(ea7), emarks7)
        ea8.insert(len(ea8), emarks8)
        t1.insert(len(t1), tmarks1)
        t2.insert(len(t2), tmarks2)
        t3.insert(len(t3), tmarks3)
        t4.insert(len(t4), tmarks4)
        t5.insert(len(t5), tmarks5)
        t6.insert(len(t6), tmarks6)
        t7.insert(len(t7), tmarks7)
        t8.insert(len(t8), tmarks8)
        r1.insert(len(r1), result1)
        r2.insert(len(r2), result2)
        r3.insert(len(r3), result3)
        r4.insert(len(r4), result4)
        r5.insert(len(r5), result5)
        r6.insert(len(r6), result6)
        r7.insert(len(r7), result7)
        r8.insert(len(r8), result8)
    else:
        ia1.insert(len(ia1), "-")
        ia2.insert(len(ia2), "-")
        ia3.insert(len(ia3), "-")
        ia4.insert(len(ia4), "-")
        ia5.insert(len(ia5), "-")
        ia6.insert(len(ia6), "-")
        ia7.insert(len(ia7), "-")
        ia8.insert(len(ia8), "-")
        ea1.insert(len(ea1), "-")
        ea2.insert(len(ea2), "-")
        ea3.insert(len(ea3), "-")
        ea4.insert(len(ea4), "-")
        ea5.insert(len(ea5), "-")
        ea6.insert(len(ea6), "-")
        ea7.insert(len(ea7), "-")
        ea8.insert(len(ea8), "-")
        t1.insert(len(t1), "-")
        t2.insert(len(t2), "-")
        t3.insert(len(t3), "-")
        t4.insert(len(t4), "-")
        t5.insert(len(t5), "-")
        t6.insert(len(t6), "-")
        t7.insert(len(t7), "-")
        t8.insert(len(t8), "-")
        r1.insert(len(r1), "-")
        r2.insert(len(r2), "-")
        r3.insert(len(r3), "-")
        r4.insert(len(r4), "-")
        r5.insert(len(r5), "-")
        r6.insert(len(r6), "-")
        r7.insert(len(r7), "-")
        r8.insert(len(r8), "-")
for i in range(1, 218):
    usn_list.insert(len(usn_list), "1BI18CS"+str(format(i, '03d')))
done = 0
print(ia1)
print(ea1)
print(t1)
print(usn_list)
print(names)
while done==0 :

    print("WRITE:-")
    print("Menu:-\n 1)Write to Spreadsheet\n 2)Write to Database\n 3)Both\n 4)Exit")
    choice = int(input("Enter your choice:-"))
    if choice == 1:
        Write.write_to_excel(usn_list, names, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8,  ea1, ea2, ea3, ea4, ea5, ea6,
                             ea7, ea8,  t1, t2, t3, t4, t5, t6, t7, t8,  r1, r2, r3, r4, r5, r6, r7, r8
                             )
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
        Write.write_to_excel(usn_list, names, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8, ea1, ea2, ea3, ea4, ea5, ea6,
                             ea7, ea8, t1, t2, t3, t4, t5, t6, t7, t8, r1, r2, r3, r4, r5, r6, r7, r8
                             )
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
