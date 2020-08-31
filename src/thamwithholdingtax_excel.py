import pandas as pd
from exthaminterestdata import ExThamInterest
import re
import json
import withholdingtax as wht


CONFIG_FILE = "./env/config_excel.json"

if __name__ == "__main__":
    print("***** Thamturakit Social Enterprise *****")
    print("***** Generate withholdingtax pdf for excel")
    print("--------------------------------------------------")
    print('*** when exit please type "exit in first row or end row" ')

    with open(CONFIG_FILE, 'r') as fp:
        configdict = json.load(fp)

    try:
        thamdata = ExThamInterest()
        thamdata.openexcel(configdict["filename"], configdict["sheet_name"])
    except:
        print("--!!! Error open excel file --")
        exit()

    while True:

        fr = input("First row:")
        er = input("End row:")

        if fr=="exit" or er=="exit":
            exit()
        elif not fr.isdecimal() or not er.isdecimal():
            print('!!Oop:: input Error please input row')
            continue
    
        try:
            for row in range(int(fr)-2, int(er)+1-2):
                
                # tax date / book number / number is empty -> don't create pdf file
                if not (thamdata.gettaxdate(row) and thamdata.getbooknumber(row) and thamdata.getnumber(row)):
                    print("Data loss : tax data, book number, number in row {}".format(row+2))
                    continue

                whtdict = thamdata.getwhtdict(row)
                whtdict["taxdeduct_name"] = "บริษัท ธรรมธุรกิจ วิสาหกิจเพื่อสังคม จำกัด"
                whtdict["taxdeduct_id"] = "0 5055 56005 09 1"
                whtdict["taxdeduct_address"] = "8/2 หมู่ 4 ต.ดอนฉิมพลี อ.บางน้ำเปรี้ยว จ.ฉะเชิงเทรา 24170"

                fullname = "{} {}".format(thamdata.getfirstname(row), thamdata.getlastname(row))
                fullname = re.sub("[/\:*?<>|.]", "", fullname)
                filename = "{}{}_{}_{}.pdf".format(configdict["output"], thamdata.getbooknumber(row), 
                                                        thamdata.getnumber(row), fullname)
                wht.write_withholdingtax(filename, whtdict)
                print("##sucess create row {} to file : {}".format(row+2, filename))

        except:
            print("----!!! Error create pdf file ----")
            exit()