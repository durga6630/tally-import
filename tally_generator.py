#!/usr/bin/env python3
"""
Tally Prime Sales Voucher XML Generator
───────────────────────────────────────
1) Generates a sample XLSX template
2) Reads filled XLSX and produces Tally-formatted XML for import

Usage:
  python tally_generator.py template    # creates sample.xlsx
  python tally_generator.py sample.xlsx # reads & generates XML
"""

import sys, os, datetime, json
from xml.dom import minidom

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl

PARTY = "Shringar Fashion"
SALES_LEDGER = "Sales"
LABOUR_LEDGER = "Labour Charges"
STOCK_ITEM = "Raw Materials"
COMPANY = "Aruna Jobwork (Jewellery) - 25-26"
GODOWN = "Main Location"
BATCH = "Primary Batch"
STATE = "West Bengal"

HEADERS = [
    ("Date", "DD.MM.YYYY  e.g. 01.04.2026"),
    ("Narration", "Invoice / ref description"),
    ("Raw Material", "Stock item amount (numeric)"),
    ("Labour Charges", "Labour ledger amount (numeric)"),
    ("Total Amount", "Grand total (numeric)"),
]

TEMPLATE_ROWS = [
    ["01.04.2026", "INV-001", "10000.00", "1500.00", "11500.00"],
    ["02.04.2026", "INV-002", "25000.00", "3200.00", "28200.00"],
    ["03.04.2026", "INV-003", "", "", ""],
]


def create_template(path="sample.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    for c, (h, t) in enumerate(HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center")
    for c, (h, t) in enumerate(HEADERS, 1):
        ws.cell(2, c, t)
        ws.cell(2, c).font = openpyxl.styles.Font(italic=True, color="888888", size=9)
    for r, row in enumerate(TEMPLATE_ROWS, 4):
        for c, val in enumerate(row, 1):
            ws.cell(r, c, val)
    for c in range(1, 6):
        ws.column_dimensions[chr(64 + c)].width = 22
    wb.save(path)
    wb.close()
    print(f"[OK] Template saved: {os.path.abspath(path)}")
    print(f"  Fill the yellow rows with your data and save.")
    print(f"  Then run this program again and choose option 2.")


def fmt_date(d):
    parts = str(d).strip().split(".")
    return f"{parts[2]}{parts[1]}{parts[0]}" if len(parts) == 3 else str(d).replace("-", "")


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


B = "&#4; "


def make_company_item(company):
    return f"""       <STATICVARIABLES>
        <SVCURRENTCOMPANY>{esc(company)}</SVCURRENTCOMPANY>
       </STATICVARIABLES>"""


def make_voucher(i, date, narration, raw_mat, labour, total):
    td = fmt_date(date)
    ref = esc(narration) if narration else f"VCH-{i}"
    vno = i
    r = raw_mat
    l = labour
    t = total

    lines = []
    a = lines.append
    a(f'''     <TALLYMESSAGE xmlns:UDF="TallyUDF">
      <VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">
       <DATE>{td}</DATE>
       <VCHSTATUSDATE>{td}</VCHSTATUSDATE>
       <NARRATION>{ref}</NARRATION>
       <CSTFORMISSUETYPE>{B}Not Applicable</CSTFORMISSUETYPE>
       <CSTFORMRECVTYPE>{B}Not Applicable</CSTFORMRECVTYPE>
       <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
       <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
       <VCHTYPENAME>Sales</VCHTYPENAME>
       <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
       <PARTYNAME>{esc(PARTY)}</PARTYNAME>
       <PARTYLEDGERNAME>{esc(PARTY)}</PARTYLEDGERNAME>
       <VOUCHERNUMBER>{vno}</VOUCHERNUMBER>
       <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
       <NUMBERINGSTYLE>Auto Retain</NUMBERINGSTYLE>
       <EFFECTIVEDATE>{td}</EFFECTIVEDATE>
       <DIFFACTUALQTY>No</DIFFACTUALQTY>
       <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
       <ISDELETED>No</ISDELETED>
       <ISSECURITYONWHENERED>No</ISSECURITYONWHENERED>
       <ASORIGINAL>No</ASORIGINAL>
       <AUDITED>No</AUDITED>
       <ISCOMMONPARTY>No</ISCOMMONPARTY>
       <FORJOBCOSTING>No</FORJOBCOSTING>
       <ISOPTIONAL>No</ISOPTIONAL>
       <USEFOREXCISE>No</USEFOREXCISE>
       <ISFORJOBWORKIN>No</ISFORJOBWORKIN>
       <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>
       <USEFORINTEREST>No</USEFORINTEREST>
       <USEFORGAINLOSS>No</USEFORGAINLOSS>
       <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>
       <USEFORCOMPOUND>No</USEFORCOMPOUND>
       <USEFORSERVICETAX>No</USEFORSERVICETAX>
       <ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>
       <ISSYSTEM>No</ISSYSTEM>
       <ISFETCHEDONLY>No</ISFETCHEDONLY>
       <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>
       <ISCANCELLED>No</ISCANCELLED>
       <ISONHOLD>No</ISONHOLD>
       <ISSUMMARY>No</ISSUMMARY>
       <ISECOMMERCESUPPLY>No</ISECOMMERCESUPPLY>
       <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>
       <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>
       <IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>
       <CMPGSTISOTHTERRITORYASSESSEE>No</CMPGSTISOTHTERRITORYASSESSEE>
       <PARTYGSTISOTHTERRITORYASSESSEE>No</PARTYGSTISOTHTERRITORYASSESSEE>
       <IRNJSONEXPORTED>No</IRNJSONEXPORTED>
       <IRNCANCELLED>No</IRNCANCELLED>
       <IGNOREGSTCONFLICTINMIG>No</IGNOREGSTCONFLICTINMIG>
       <ISOPBALTRANSACTION>No</ISOPBALTRANSACTION>
       <IGNOREGSTFORMATVALIDATION>No</IGNOREGSTFORMATVALIDATION>
       <ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
       <IGNOREGSTOPTIONALUNCERTAIN>No</IGNOREGSTOPTIONALUNCERTAIN>
       <UPDATESUMMARYVALUES>No</UPDATESUMMARYVALUES>
       <ISEWAYBILLAPPLICABLE>No</ISEWAYBILLAPPLICABLE>
       <ISDELETEDRETAINED>No</ISDELETEDRETAINED>
       <ISNULL>No</ISNULL>
       <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>
       <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>
       <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>
       <ISEXER1NOPOVERWRITE>No</ISEXER1NOPOVERWRITE>
       <ISEXF2NOPOVERWRITE>No</ISEXF2NOPOVERWRITE>
       <ISEXER3NOPOVERWRITE>No</ISEXER3NOPOVERWRITE>
       <VATDEALERTYPE>Regular</VATDEALERTYPE>
       <STATENAME>{STATE}</STATENAME>
       <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
       <GSTREGISTRATIONTYPE>{B}Unknown</GSTREGISTRATIONTYPE>
       <VCHGSTCLASS>{B}Not Applicable</VCHGSTCLASS>''')

    if r > 0:
        a(f'''       <ALLINVENTORYENTRIES.LIST>
        <STOCKITEMNAME>{esc(STOCK_ITEM)}</STOCKITEMNAME>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
        <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
        <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
        <AMOUNT>{r:.2f}</AMOUNT>
        <BATCHALLOCATIONS.LIST>
         <GODOWNNAME>{esc(GODOWN)}</GODOWNNAME>
         <BATCHNAME>{esc(BATCH)}</BATCHNAME>
         <INDENTNO>{B}Not Applicable</INDENTNO>
         <ORDERNO>{B}Not Applicable</ORDERNO>
         <TRACKINGNUMBER>{B}Not Applicable</TRACKINGNUMBER>
         <AMOUNT>{r:.2f}</AMOUNT>
        </BATCHALLOCATIONS.LIST>
        <ACCOUNTINGALLOCATIONS.LIST>
         <OLDAUDITENTRYIDS.LIST TYPE="Number">
          <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
         </OLDAUDITENTRYIDS.LIST>
         <LEDGERNAME>{esc(SALES_LEDGER)}</LEDGERNAME>
         <GSTCLASS>{B}Not Applicable</GSTCLASS>
         <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
         <LEDGERFROMITEM>No</LEDGERFROMITEM>
         <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
         <AMOUNT>{r:.2f}</AMOUNT>
         <BANKALLOCATIONS.LIST></BANKALLOCATIONS.LIST>
         <BILLALLOCATIONS.LIST></BILLALLOCATIONS.LIST>
         <INTERESTCOLLECTION.LIST></INTERESTCOLLECTION.LIST>
         <OLDAUDITENTRIES.LIST></OLDAUDITENTRIES.LIST>
         <ACCOUNTAUDITENTRIES.LIST></ACCOUNTAUDITENTRIES.LIST>
         <AUDITENTRIES.LIST></AUDITENTRIES.LIST>
         <INPUTCRALLOCS.LIST></INPUTCRALLOCS.LIST>
         <CENVATDUTYALLOCATIONS.LIST></CENVATDUTYALLOCATIONS.LIST>
         <STPYMTDETAILS.LIST></STPYMTDETAILS.LIST>
         <EXCISEPAYMENTALLOCATIONS.LIST></EXCISEPAYMENTALLOCATIONS.LIST>
         <TAXBILLALLOCATIONS.LIST></TAXBILLALLOCATIONS.LIST>
         <TAXOBJECTALLOCATIONS.LIST></TAXOBJECTALLOCATIONS.LIST>
         <TDSEXPENSEALLOCATIONS.LIST></TDSEXPENSEALLOCATIONS.LIST>
         <VATSTATUTORYDETAILS.LIST></VATSTATUTORYDETAILS.LIST>
         <COSTTRACKALLOCATIONS.LIST></COSTTRACKALLOCATIONS.LIST>
         <REFVOUCHERDETAILS.LIST></REFVOUCHERDETAILS.LIST>
         <INVOICEWISEDETAILS.LIST></INVOICEWISEDETAILS.LIST>
         <VATITCDETAILS.LIST></VATITCDETAILS.LIST>
         <ADVANCETAXDETAILS.LIST></ADVANCETAXDETAILS.LIST>
         <TAXTYPEALLOCATIONS.LIST></TAXTYPEALLOCATIONS.LIST>
        </ACCOUNTINGALLOCATIONS.LIST>
        <DUTYHEADDETAILS.LIST></DUTYHEADDETAILS.LIST>
        <RATEDETAILS.LIST></RATEDETAILS.LIST>
        <SUPPLEMENTARYDUTYHEADDETAILS.LIST></SUPPLEMENTARYDUTYHEADDETAILS.LIST>
        <TAXOBJECTALLOCATIONS.LIST></TAXOBJECTALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST></REFVOUCHERDETAILS.LIST>
        <EXCISEALLOCATIONS.LIST></EXCISEALLOCATIONS.LIST>
        <EXPENSEALLOCATIONS.LIST></EXPENSEALLOCATIONS.LIST>
       </ALLINVENTORYENTRIES.LIST>''')

    a(f'''       <CONTRITRANS.LIST></CONTRITRANS.LIST>
       <EWAYBILLERRORLIST.LIST></EWAYBILLERRORLIST.LIST>
       <IRNERRORLIST.LIST></IRNERRORLIST.LIST>
       <HARYANAVAT.LIST></HARYANAVAT.LIST>
       <SUPPLEMENTARYDUTYHEADDETAILS.LIST></SUPPLEMENTARYDUTYHEADDETAILS.LIST>''')

    a(f'''       <LEDGERENTRIES.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <LEDGERNAME>{esc(PARTY)}</LEDGERNAME>
        <GSTCLASS>{B}Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <AMOUNT>{-t:.2f}</AMOUNT>
        <BANKALLOCATIONS.LIST></BANKALLOCATIONS.LIST>
        <BILLALLOCATIONS.LIST>
         <NAME>{vno}</NAME>
         <BILLTYPE>New Ref</BILLTYPE>
         <TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>
         <AMOUNT>{-t:.2f}</AMOUNT>
        </BILLALLOCATIONS.LIST>
        <INTERESTCOLLECTION.LIST></INTERESTCOLLECTION.LIST>
        <OLDAUDITENTRIES.LIST></OLDAUDITENTRIES.LIST>
        <ACCOUNTAUDITENTRIES.LIST></ACCOUNTAUDITENTRIES.LIST>
        <AUDITENTRIES.LIST></AUDITENTRIES.LIST>
        <INPUTCRALLOCS.LIST></INPUTCRALLOCS.LIST>
        <CENVATDUTYALLOCATIONS.LIST></CENVATDUTYALLOCATIONS.LIST>
        <STPYMTDETAILS.LIST></STPYMTDETAILS.LIST>
        <EXCISEPAYMENTALLOCATIONS.LIST></EXCISEPAYMENTALLOCATIONS.LIST>
        <TAXBILLALLOCATIONS.LIST></TAXBILLALLOCATIONS.LIST>
        <TAXOBJECTALLOCATIONS.LIST></TAXOBJECTALLOCATIONS.LIST>
        <TDSEXPENSEALLOCATIONS.LIST></TDSEXPENSEALLOCATIONS.LIST>
        <VATSTATUTORYDETAILS.LIST></VATSTATUTORYDETAILS.LIST>
        <COSTTRACKALLOCATIONS.LIST></COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST></REFVOUCHERDETAILS.LIST>
        <INVOICEWISEDETAILS.LIST></INVOICEWISEDETAILS.LIST>
        <VATITCDETAILS.LIST></VATITCDETAILS.LIST>
        <ADVANCETAXDETAILS.LIST></ADVANCETAXDETAILS.LIST>
        <TAXTYPEALLOCATIONS.LIST></TAXTYPEALLOCATIONS.LIST>
       </LEDGERENTRIES.LIST>''')

    if l > 0:
        a(f'''       <LEDGERENTRIES.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
         <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <LEDGERNAME>{esc(LABOUR_LEDGER)}</LEDGERNAME>
        <GSTCLASS>{B}Not Applicable</GSTCLASS>
        <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
        <LEDGERFROMITEM>No</LEDGERFROMITEM>
        <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
        <AMOUNT>{l:.2f}</AMOUNT>
        <VATEXPAMOUNT>{l:.2f}</VATEXPAMOUNT>
        <BANKALLOCATIONS.LIST></BANKALLOCATIONS.LIST>
        <BILLALLOCATIONS.LIST></BILLALLOCATIONS.LIST>
        <INTERESTCOLLECTION.LIST></INTERESTCOLLECTION.LIST>
        <OLDAUDITENTRIES.LIST></OLDAUDITENTRIES.LIST>
        <ACCOUNTAUDITENTRIES.LIST></ACCOUNTAUDITENTRIES.LIST>
        <AUDITENTRIES.LIST></AUDITENTRIES.LIST>
        <INPUTCRALLOCS.LIST></INPUTCRALLOCS.LIST>
        <CENVATDUTYALLOCATIONS.LIST></CENVATDUTYALLOCATIONS.LIST>
        <STPYMTDETAILS.LIST></STPYMTDETAILS.LIST>
        <EXCISEPAYMENTALLOCATIONS.LIST></EXCISEPAYMENTALLOCATIONS.LIST>
        <TAXBILLALLOCATIONS.LIST></TAXBILLALLOCATIONS.LIST>
        <TAXOBJECTALLOCATIONS.LIST></TAXOBJECTALLOCATIONS.LIST>
        <TDSEXPENSEALLOCATIONS.LIST></TDSEXPENSEALLOCATIONS.LIST>
        <VATSTATUTORYDETAILS.LIST></VATSTATUTORYDETAILS.LIST>
        <COSTTRACKALLOCATIONS.LIST></COSTTRACKALLOCATIONS.LIST>
        <REFVOUCHERDETAILS.LIST></REFVOUCHERDETAILS.LIST>
        <INVOICEWISEDETAILS.LIST></INVOICEWISEDETAILS.LIST>
        <VATITCDETAILS.LIST></VATITCDETAILS.LIST>
        <ADVANCETAXDETAILS.LIST></ADVANCETAXDETAILS.LIST>
        <TAXTYPEALLOCATIONS.LIST></TAXTYPEALLOCATIONS.LIST>
       </LEDGERENTRIES.LIST>''')

    a(f'''       <GST.LIST></GST.LIST>
       <STKJRNLADDLCOSTDETAILS.LIST></STKJRNLADDLCOSTDETAILS.LIST>
       <PAYROLLMODEOFPAYMENT.LIST></PAYROLLMODEOFPAYMENT.LIST>
       <ATTDRECORDS.LIST></ATTDRECORDS.LIST>
       <GSTEWAYCONSIGNORADDRESS.LIST></GSTEWAYCONSIGNORADDRESS.LIST>
       <VCHGSTCLASS>{B}Not Applicable</VCHGSTCLASS>
       <TEMPGSTRATEDETAILS.LIST></TEMPGSTRATEDETAILS.LIST>
       <TEMPGSTADVADJUSTED.LIST></TEMPGSTADVADJUSTED.LIST>
       <GSTBUYERADDRESS.LIST></GSTBUYERADDRESS.LIST>
      </VOUCHER>
     </TALLYMESSAGE>''')

    return "\n".join(lines)


def generate_xml(path, out_path=None):
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    rows = []
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, 1).value
        n = ws.cell(r, 2).value
        rm = ws.cell(r, 3).value
        lc = ws.cell(r, 4).value
        tt = ws.cell(r, 5).value
        if d and rm is not None and str(rm).strip():
            try:
                raw = round(float(rm), 2)
                lab = round(float(lc), 2) if lc and str(lc).strip() else 0
                total = round(float(tt), 2) if tt and str(tt).strip() else raw + lab
                rows.append({
                    "date": str(d).strip(),
                    "narration": str(n or "").strip(),
                    "raw_mat": raw,
                    "labour": lab,
                    "total": total,
                })
            except (ValueError, TypeError):
                pass

    parts = []
    a = parts.append
    a('<ENVELOPE>')
    a(' <HEADER>')
    a('  <TALLYREQUEST>Import Data</TALLYREQUEST>')
    a(' </HEADER>')
    a(' <BODY>')
    a('  <IMPORTDATA>')
    a('   <REQUESTDESC>')
    a('    <REPORTNAME>All Masters</REPORTNAME>')
    a(make_company_item(COMPANY))
    a('   </REQUESTDESC>')
    a('   <REQUESTDATA>')

    for i, rd in enumerate(rows):
        a(make_voucher(i + 1, rd["date"], rd["narration"], rd["raw_mat"], rd["labour"], rd["total"]))

    a('   </REQUESTDATA>')
    a('  </IMPORTDATA>')
    a(' </BODY>')
    a('</ENVELOPE>')

    xml = "\n".join(parts)

    if out_path is None:
        base = os.path.splitext(os.path.basename(path))[0]
        out_path = f"{base}_tally.xml"

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(xml)

    total_billed = sum(r["total"] for r in rows)
    print(f"\n[OK] Generated: {out_path}")
    print(f"  {len(rows)} voucher(s) | Vch No 1 to {len(rows)}")
    print(f"  Party: {PARTY}")
    print(f"  Company: {COMPANY}")
    print(f"  Total billed: Rs {total_billed:,.2f}")
    print(f"\n  Import in Tally: Gateway of Tally > Import Data > Masters and Vouchers")
    return out_path


def clean_path(s):
    return s.strip().strip('"\'')

def interactive():
    print("""
 Tally Prime Sales Voucher Generator
 ====================================
 1. Create template XLSX file
 2. Generate XML from existing XLSX
 3. Save config file
 4. Exit
""")
    while True:
        try:
            c = input(" Choose (1-4): ").strip()
            if c == "1":
                p = clean_path(input(" Output path [sample.xlsx]: ")) or "sample.xlsx"
                create_template(p)
                print("\n Open the XLSX, fill your data, save it, then run option 2.")
                break
            elif c == "2":
                p = clean_path(input(" Path to your XLSX file: "))
                if not p or not os.path.isfile(p):
                    print(f" File not found: {p}")
                else:
                    generate_xml(p)
                break
            elif c == "3":
                cfg = {
                    "company": COMPANY,
                    "party": PARTY,
                    "sales_ledger": SALES_LEDGER,
                    "labour_ledger": LABOUR_LEDGER,
                    "stock_item": STOCK_ITEM,
                    "godown": GODOWN,
                    "batch": BATCH,
                    "state": STATE,
                }
                with open("tally_config.json", "w") as f:
                    json.dump(cfg, f, indent=2)
                print("[OK] Config saved: tally_config.json")
                break
            elif c == "4":
                break
            else:
                print(" Invalid. Enter 1-4.")
        except Exception as e:
            print(f"\n Error: {e}")
            break
    input("\n Press Enter to exit...")


if __name__ == "__main__":
    try:
        if len(sys.argv) < 2:
            interactive()
            sys.exit(0)

        cmd = sys.argv[1]

        if cmd == "template":
            out = sys.argv[2] if len(sys.argv) > 2 else "sample.xlsx"
            create_template(out)
        elif cmd == "config":
            config = {
                "company": COMPANY,
                "party": PARTY,
                "sales_ledger": SALES_LEDGER,
                "labour_ledger": LABOUR_LEDGER,
                "stock_item": STOCK_ITEM,
                "godown": GODOWN,
                "batch": BATCH,
                "state": STATE,
            }
            with open("tally_config.json", "w") as f:
                json.dump(config, f, indent=2)
            print("[OK] Config saved: tally_config.json")
        elif cmd.endswith(".xlsx"):
            generate_xml(cmd, sys.argv[2] if len(sys.argv) > 2 else None)
        else:
            print(f"Unknown: {cmd}")
            print(__doc__)
    except Exception as e:
        print(f"\n Error: {e}")

    if len(sys.argv) > 1:
        input("\n Press Enter to exit...")
