#
# Write data to withholding tax form
#
# Write For : Thamturakit Social Enterprise
# Author: Yothin Setthachatanan
# update 13/2/2020
#
#

from oauth2client.service_account import ServiceAccountCredentials
import gspread
import json
from pythainlp.util import bahttext, thai_strftime
from datetime import datetime
import withholdingtax as wht


GG_COL_INDEX = {
    "book_number": 1,
    "number": 2,
    "taxpayer_prefix": 10,
    "taxpayer_firstname": 11,
    "taxpayer_lastname": 12,
    "taxpayer_id": 13,
    "taxpayer_addressno": 14,
    "taxpayer_soi": 15,
    "taxpayer_street": 16,
    "taxpayer_subdist": 17,
    "taxpayer_district": 18,
    "taxpayer_province": 19,
    "taxpayer_postcode": 20,
    "invest_amount": 23,
    "interest_amount": 24,
    "tax_amount": 25,
    "tax_date": 3
}

class GGThamInterest:
    def __init__(self, config_file):
        with open(config_file, 'r') as fp:
            self.configdict = json.load(fp)
            # print(self.configdict)
        
    def opensheet(self):
        # connect google Sheet api
        try:
            credential = ServiceAccountCredentials.from_json_keyfile_name(
                            self.configdict['credential_file'], 
                            self.configdict['scope'])
            client = gspread.authorize(credential)        
        except:
            raise Exception('Cant connect google sheet API')

        # open sheet
        try:    
            self.sheet = client.open_by_key(self.configdict['sheet_key'])
            self.worksheet = self.sheet.worksheet(self.configdict['sheet_name'])
        except:
            raise Exception('Cant open google sheet')
        
    def getrow(self, row):
        try:
            row_list = self.worksheet.row_values(row)
        except:
            raise Exception('Cant load data in {} row'.format(row))
        finally:
            return GGThamInterestData(row_list)

class GGThamInterestData:
    def __init__(self, row_data):
        self.row_data = row_data

    def getwhtdict(self):
        whtdict = {}
        whtdict['book_number'] = self.getbooknumber()
        whtdict['number'] = self.getnumber()

        # tax payer part
        whtdict['taxpayer_name'] = self.getfullname()
        whtdict['taxpayer_id'] = self.getidwithformat()
        whtdict['taxpayer_address'] = self.getfulladdress()

        # section 40 part
        dt = self.gettaxdate_thaiformat()
        whtdict['section40_date'] = dt
        whtdict['section40_amount'] = self.getinterestamount()
        whtdict['section40_tax'] = self.gettaxamount()

        # total part
        whtdict['total_amount'] = self.getinterestamount()
        whtdict['total_tax'] = self.gettaxamount()
        whtdict['total_bahttext'] = "( {} )".format(bahttext(self.gettaxamount()))
        
        # Issue date / month / year part
        whtdict['issue_date'] = "{}".format(dt[0:2])
        whtdict['issue_month'] = "     {}".format(dt[3:7])
        whtdict['issue_year'] =  "{}".format(dt[8:])
        return whtdict
        
    # return id number without format
    def getid(self):
        id = self.row_data[GG_COL_INDEX['taxpayer_id']]
        if id != "" : return id.strip().replace("-","")

    # ID FORMAT : x xxxx xxxxx xx x
    def getidwithformat(self):
        str = self.getid()
        if len(str) == 13:
            return "{0} {1} {2} {3} {4}".format(str[0],str[1:5],str[5:10],str[10:12],str[12])
        else:
            return str

    def getprefix(self):
        return self.row_data[GG_COL_INDEX["taxpayer_prefix"]].strip().replace("\u200b","")
    
    def getfirstname(self):
        return self.row_data[GG_COL_INDEX['taxpayer_firstname']].strip().replace("\u200b","")

    def getlastname(self):
        return self.row_data[GG_COL_INDEX['taxpayer_lastname']].strip().replace("\u200b","")

    # FULLNAME: PREFIX FIRSTNAME LASTNAME
    def getfullname(self):
        prefix = self.getprefix()
        name = "{} {}".format(self.getfirstname(), self.getlastname()) 
        if prefix == "": return name
        else: return "{} {}".format(prefix, name)    
        
    def getfulladdress(self):
        address = self.row_data[GG_COL_INDEX["taxpayer_addressno"]].strip().replace("\u200b","")

        # If empty do not con
        soi = self.row_data[GG_COL_INDEX["taxpayer_soi"]].strip().replace("\u200b","")
        if soi != "":
            address = "{} ซ.{}".format(address, soi)

        street = self.row_data[GG_COL_INDEX["taxpayer_street"]].strip().replace("\u200b","")
        if street !="":
            address = "{} ถ.{}".format(address, street)

        subdist = self.row_data[GG_COL_INDEX["taxpayer_subdist"]].strip().replace("\u200b","")
        dist = self.row_data[GG_COL_INDEX["taxpayer_district"]].strip().replace("\u200b","")
        province = self.row_data[GG_COL_INDEX["taxpayer_province"]].strip().replace("\u200b","")
        postcode = self.row_data[GG_COL_INDEX["taxpayer_postcode"]].strip().replace("\u200b","")
        pre_subdist = "ต."
        pre_dist = "อ."
        pre_province = "จ."

        # PROVINCE IS BANGKOK use แขวง and เขต
        if province.find("กทม") >= 0 or province.find("กรุงเทพ") >= 0:
        #if province == "กทม" or province == "กรุงเทพ":
            pre_subdist = "แขวง"
            pre_dist = "เขต"
            pre_province = ""

        return "{} {}{} {}{} {}{} {}".format(address, pre_subdist, subdist, 
                                    pre_dist, dist, pre_province, province, postcode)

    def convertstrtofloat(self, str):
        return float(str.strip().replace(",",""))

    def getbooknumber(self):
        return self.row_data[GG_COL_INDEX["book_number"]].strip()

    def getnumber(self):
        return self.row_data[GG_COL_INDEX["number"]].strip()

    def gettaxdate(self):
        return datetime.strptime(self.row_data[GG_COL_INDEX["tax_date"]], "%d/%m/%Y")

    def gettaxdate_thaiformat(self):
        return thai_strftime(self.gettaxdate(), "%d %b %Y")

    def gettaxamount(self):
        return self.convertstrtofloat(self.row_data[GG_COL_INDEX['tax_amount']].strip())

    def getinvestamount(self):
        return self.convertstrtofloat(self.row_data[GG_COL_INDEX['inverst_amount']].strip())

    def getinterestamount(self):
        return self.convertstrtofloat(self.row_data[GG_COL_INDEX['interest_amount']].strip())
      
   
if __name__ == "__main__":
    ggsheet = GGThamInterest("./env/config_gg.json")
    ggsheet.opensheet()
    
    data = ggsheet.getrow(18)
    whtdict = data.getwhtdict()

    whtdict["taxdeduct_name"] = "บริษัท ธรรมธุรกิจ วิสาหกิจเพื่อสังคม จำกัด"
    whtdict["taxdeduct_id"] = "0 5055 56005 09 1"
    whtdict["taxdeduct_address"] = "8/2 หมู่ 4 ต.ดอนฉิมพลี อ.บางน้ำเปรี้ยว จ.ฉะเชิงเทรา 24170"

    print(whtdict)

    wht.write_withholdingtax("./output/test1.pdf", whtdict)
    