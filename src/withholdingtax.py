#
# Write data to withholding tax form
#
# Write For : Thamturakit Social Enterprise
# Author: Yothin Setthachatanan
# update 12/2/2020
#
#

import pdfrw 

'''
DATA Dictionary structure key
+ book_number
+ number
+ taxdeduct_name
+ taxdeduct_id
+ taxdeduct_address
+ taxpayer_name
+ taxpayer_id
+ taxpayer_address
+ section40_date
+ section40_amount
+ section40_tax
+ total_amount
+ total_tax
+ total_bahttext
+ issue_date
+ issue_month
+ issue_year
'''
# POSTION in PDF structure
PDF_DICT = {
    "book_number": 2,
    "number": 3,
    "taxdeduct_name": 5,
    "taxdeduct_id": 4,
    "taxdeduct_address": 1,
    "taxpayer_name": 8,
    "taxpayer_id": 7,
    "taxpayer_address": 10,
    "section40_date": 28,
    "section40_amount": 29,
    "section40_tax": 30,
    "total_amount": 67,
    "total_tax": 68,
    "total_bahttext": 69,
    "pnd2_chkbox": 14,
    "withholding_chkbox": 73,
    "issue_date": 78,
    "issue_month": 79,
    "issue_year": 80
}

# MASTER WITHHOLDING TAX FORM
WITHHOLDINGTAX_FORM_NAME = "./env/withholdingtax_form.pdf"

# สำหรับสร้าง PDF ภาษีหัก ณ ที่จ่าย ของดอกเบี้ย ผลตอบแทน ภงด 2
# @param 
#  filename : output filename
#  datadict :  data dictionary

def write_withholdingtax(filename, datadict):
    # Read withholdingtax form
    pdf = pdfrw.PdfReader(WITHHOLDINGTAX_FORM_NAME)

    # Display text in fill form
    pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

    # -- write data to pdf ------
    for d in datadict:
        pdf.Root.Pages.Kids[0].Annots[PDF_DICT[d]].update(
            pdfrw.PdfDict(V='{}'.format(datadict[d]))
        )
        # print(pdf.Root.Pages.Kids[0].Annots[PDF_DICT[d]])

    # Click Checkbox : Format '/V': '/Yes'
    # ภงด 2
    pdf.Root.Pages.Kids[0].Annots[PDF_DICT["pnd2_chkbox"]].update(
        pdfrw.PdfDict(V=pdfrw.PdfObject('/Yes'))
    )
    pdf.Root.Pages.Kids[0].Annots[PDF_DICT["pnd2_chkbox"]].update(
        pdfrw.PdfDict(AS=pdfrw.PdfObject('/Yes'))
    )

    # หัก ณ ที่ จ่าย
    pdf.Root.Pages.Kids[0].Annots[PDF_DICT["withholding_chkbox"]].update(
        pdfrw.PdfDict(V=pdfrw.PdfObject('/Yes'))
    )

    pdf.Root.Pages.Kids[0].Annots[PDF_DICT["withholding_chkbox"]].update(
        pdfrw.PdfDict(AS=pdfrw.PdfObject('/Yes'))
    )

    # Save new pdf
    pdfrw.PdfWriter().write(filename, pdf)


if __name__ == "__main__":

    datadict = {
        "book_number": 62,
        "number": 100,
        "taxdeduct_name": "บริษัท ธรรมธุรกิจ วิสาหกิจเพื่อสังคม จำกัด",
        "taxdeduct_id": "0 5055 56005 09 1",
        "taxdeduct_address": "8/2 หมู่ 4 ต.ดอนฉิมพลี อ.บางน้ำเปรี้ยว จ.ฉะเชิงเทรา 24170",
        "taxpayer_name": "นาย ยุทธชัย เปี่ยมพืช",
        "taxpayer_id": "3 1402 00159 92 9",
        "taxpayer_address": "สวนธรรมชาติมีดี เลขที่ 174  ถ.เลียบคลองชัยนาทป่าสัก ต.บ้านครัว อ.บ้านหมอ จ.สระบุรี 18270",
        "section40_date": "6 ก.ย. 2562",
        "section40_amount": 1357.50,
        "section40_tax": 203.62,
        "total_amount": 1357.50,
        "total_tax": 203.62,
        "total_bahttext": "(สองร้อยสามบาทหกสิบสองสตางค์)",
        "issue_date": "06",
        "issue_month": "      ก.ย.",
        "issue_year": "2562"
    }
        
    write_withholdingtax("./output/62_100.pdf", datadict)


 

  
 