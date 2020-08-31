import withholdingtax as wht
from ggthaminterestdata import GGThamInterest, GGThamInterestData

CONFIG_FILE = "./env/config_gg.json"

if __name__ == "__main__":

    print("***** Thamturakit Social Enterprise *****")
    print("***** Generate withholdingtax pdf for google sheet")
    print("--------------------------------------------------")
    print("(You can change sheet key and name in config.json)")
    print('*** when exit please type "exit in first row or end row" ')

    try:
        ggsheet = GGThamInterest(CONFIG_FILE)
        ggsheet.opensheet()    
    except:
        print("--!!! Error Connect google sheet --")
        exit()

    while True:

        fr = input("First row: ")
        er = input("End row: ")

        if fr=="exit" or er=="exit":
            exit()
        elif not fr.isdecimal() or not er.isdecimal():
            print('!!Oop:: input Error please input row')
            continue
    
        try:
            for row in range(int(fr), int(er)+1):
                data = ggsheet.getrow(row)
                whtdict = data.getwhtdict()
                whtdict["taxdeduct_name"] = "บริษัท ธรรมธุรกิจ วิสาหกิจเพื่อสังคม จำกัด"
                whtdict["taxdeduct_id"] = "0 5055 56005 09 1"
                whtdict["taxdeduct_address"] = "8/2 หมู่ 4 ต.ดอนฉิมพลี อ.บางน้ำเปรี้ยว จ.ฉะเชิงเทรา 24170"

                filename = "./output/{}_{}_{} {}.pdf".format(data.getbooknumber(),data.getnumber(), 
                                                            data.getfirstname(), data.getlastname())
                wht.write_withholdingtax(filename, whtdict)
                print("##sucess create {}".format(filename))

        except:
            print("----!!! Error create pdf file ----")
            exit()

