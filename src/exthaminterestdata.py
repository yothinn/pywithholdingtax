import pandas as pd
from pythainlp.util import bahttext, thai_strftime
from datetime import datetime

EX_COL_INDEX = {
    "book_number": 58,
    "number": 59,
    "taxpayer_prefix": 2,
    "taxpayer_firstname": 3,
    "taxpayer_lastname": 4,
    "taxpayer_id": 1,
    "taxpayer_addressno": 5,
    "taxpayer_soi": 6,
    "taxpayer_street": 7,
    "taxpayer_subdist": 8,
    "taxpayer_district": 9,
    "taxpayer_province": 10,
    "taxpayer_postcode": 11,
    "invest_amount": 15,
    "interest_amount": 49,
    "tax_amount": 50,
    "tax_date": 57
}

class ExThamInterest:
    def __init__(self):
        pass

    def openexcel(self, filename, sheetname):
        self.df = pd.read_excel(filename, sheetname)
        #print(self.df)

    def getwhtdict(self, row):
        whtdict = {}
        whtdict['book_number'] = self.getbooknumber(row)
        whtdict['number'] = self.getnumber(row)

        # tax payer part
        whtdict['taxpayer_name'] = self.getfullname(row)
        whtdict['taxpayer_id'] = self.getidwithformat(row)
        whtdict['taxpayer_address'] = self.getfulladdress(row)

        # section 40 part
        dt = self.gettaxdate_thaiformat(row)
        whtdict['section40_date'] = dt
        whtdict['section40_amount'] = self.getinterestamount(row)
        whtdict['section40_tax'] = self.gettaxamount(row)

        # total part
        whtdict['total_amount'] = self.getinterestamount(row)
        whtdict['total_tax'] = self.gettaxamount(row)
        whtdict['total_bahttext'] = "( {} )".format(bahttext(self.gettaxamount(row)))
        
        # Issue date / month / year part
        if dt:
            whtdict['issue_date'] = "{}".format(dt[0:2])
            whtdict['issue_month'] = "     {}".format(dt[3:7])
            whtdict['issue_year'] =  "{}".format(dt[8:])
        return whtdict

    def getid(self, row):
        id = self.df.iat[row, EX_COL_INDEX['taxpayer_id']]
        if type(id) == str:
            if not pd.isnull(id): 
                return id.strip().replace("-", "").replace(" ", "")
            else:
                return ""
        else:
            return "{}".format(id)


    # ID FORMAT : x xxxx xxxxx xx x
    def getidwithformat(self, row):
        str = self.getid(row)
        if len(str) == 13:
            return "{0} {1} {2} {3} {4}".format(str[0],str[1:5],str[5:10],str[10:12],str[12])
        else:
            return str

    def getprefix(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX["taxpayer_prefix"]]):
            return ""
        else:
            return self.df.iat[row, EX_COL_INDEX["taxpayer_prefix"]]

    def getfullname(self, row):
        prefix = self.getprefix(row)
        name = "{} {}".format(self.getfirstname(row), self.getlastname(row)) 
        if prefix == "": return name
        else: return "{} {}".format(prefix, name) 

    def getfirstname(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX["taxpayer_firstname"]]):
            return ""    
        else:
            return self.df.iat[row, EX_COL_INDEX["taxpayer_firstname"]].strip().replace(" ", "")

    def getlastname(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX["taxpayer_lastname"]]):
            return ""
        else:
            return self.df.iat[row, EX_COL_INDEX["taxpayer_lastname"]].strip().replace(" ", "")

    def getfulladdress(self, row):
        address = self.df.iat[row, EX_COL_INDEX["taxpayer_addressno"]]

        # If empty do not con
        soi = self.df.iat[row, EX_COL_INDEX["taxpayer_soi"]]
        if not pd.isnull(soi):
            address = "{} ซ.{}".format(address, soi)
        
        street = self.df.iat[row, EX_COL_INDEX["taxpayer_street"]]
        if not pd.isnull(street):
            address = "{} ถ.{}".format(address, street)

        subdist = self.df.iat[row, EX_COL_INDEX["taxpayer_subdist"]]
        dist = self.df.iat[row, EX_COL_INDEX["taxpayer_district"]]
        province = self.df.iat[row, EX_COL_INDEX["taxpayer_province"]]
        postcode = self.df.iat[row, EX_COL_INDEX["taxpayer_postcode"]]
        pre_subdist = "ต."
        pre_dist = "อ."
        pre_province = "จ."

        # PROVINCE IS BANGKOK use แขวง and เขต
        if province == "กทม" or province == "กรุงเทพ":
            pre_subdist = "แขวง"
            pre_dist = "เขต"
            pre_province = ""

        return "{} {}{} {}{} {}{} {}".format(address, pre_subdist, subdist, 
                                    pre_dist, dist, pre_province, province, postcode)

    def getbooknumber(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX["book_number"]]):
            return None
        else:
            return int(self.df.iat[row, EX_COL_INDEX["book_number"]])

    def getnumber(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX["number"]]):
            return None
        else:
            return int(self.df.iat[row, EX_COL_INDEX["number"]])

    def gettaxdate(self, row):
        ts = self.df.iat[row, EX_COL_INDEX["tax_date"]]
        if pd.isnull(ts):
            return None
        else:
            return ts.to_pydatetime()

    def gettaxdate_thaiformat(self, row):
        dt = self.gettaxdate(row)
        if dt:
            return thai_strftime(dt, "%d %b %Y")
        else:
            return None

    def convertstrtofloat(self, str):
        return float(str.strip().replace(",",""))

    def gettaxamount(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX['tax_amount']]):
            return 0.0
        else:
            return self.df.iat[row, EX_COL_INDEX['tax_amount']]

    def getinvestamount(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX['inverst_amount']]):
            return 0.0
        else:
            return self.df.iat[row, EX_COL_INDEX['inverst_amount']]

    def getinterestamount(self, row):
        if pd.isnull(self.df.iat[row, EX_COL_INDEX['interest_amount']]):
            return 0.0
        else:
            return self.df.iat[row, EX_COL_INDEX['interest_amount']]