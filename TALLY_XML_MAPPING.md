# Tally Prime XML Mapping Reference

## Overview

Mapping between Excel sales data and Tally Prime Sales voucher XML for import.
Used by `tally_generator.py` to convert XLSX → XML for Tally Gateway import.

---

## 1. Envelope Structure

```
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>Company Name</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">
            ...
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
```

### Key notes:
- Export uses `IMPORTDATA` > `REQUESTDESC` + `REQUESTDATA` (not `DATA`/`DESC`)
- Company name in `<SVCURRENTCOMPANY>` must match exactly
- Each `<TALLYMESSAGE>` = one voucher

---

## 2. Voucher Header Fields

| XML Tag | Value | Notes |
|---|---|---|
| `VCHTYPE` | `"Sales"` | Attribute on VOUCHER tag |
| `ACTION` | `"Create"` | Attribute on VOUCHER tag |
| `OBJVIEW` | `"Invoice Voucher View"` | Attribute on VOUCHER tag |
| `DATE` | `YYYYMMDD` | Excel DD.MM.YYYY → YYYYMMDD |
| `VCHSTATUSDATE` | Same as DATE | |
| `EFFECTIVEDATE` | Same as DATE | |
| `VOUCHERTYPENAME` | `"Sales"` | Must match Tally master |
| `PARTYNAME` | Party name | |
| `PARTYLEDGERNAME` | Party ledger name | Must match Tally ledger master |
| `NARRATION` | Description / ref | Free text |
| `PERSISTEDVIEW` | `"Invoice Voucher View"` | |
| `VCHENTRYMODE` | `"Item Invoice"` | Required for inventory vouchers |
| `NUMBERINGSTYLE` | `"Auto Retain"` | |
| ~`VOUCHERNUMBER`~ | Omit → Tally auto-numbers | |

### Boolean fields (all set to `No` unless noted):
`DIFFACTUALQTY`, `ISMSTFROMSYNC`, `ISDELETED`, `ISSECURITYONWHENERED`, `ASORIGINAL`, `AUDITED`, `ISCOMMONPARTY`, `FORJOBCOSTING`, `ISOPTIONAL`, `USEFOREXCISE`, `ISFORJOBWORKIN`, `ALLOWCONSUMPTION`, `USEFORINTEREST`, `USEFORGAINLOSS`, `USEFORGODOWNTRANSFER`, `USEFORCOMPOUND`, `USEFORSERVICETAX`, `ISREVERSECHARGEAPPLICABLE`, `ISSYSTEM`, `ISFETCHEDONLY`, `ISGSTOVERRIDDEN`, `ISCANCELLED`, `ISONHOLD`, `ISSUMMARY`, `ISECOMMERCESUPPLY`, `ISBOENOTAPPLICABLE`, `ISGSTSECSEVENAPPLICABLE`, `IGNOREEINVVALIDATION`, `CMPGSTISOTHTERRITORYASSESSEE`, `PARTYGSTISOTHTERRITORYASSESSEE`, `IRNJSONEXPORTED`, `IRNCANCELLED`, `IGNOREGSTCONFLICTINMIG`, `ISOPBALTRANSACTION`, `IGNOREGSTFORMATVALIDATION`, `UPDATESUMMARYVALUES`, `ISEWAYBILLAPPLICABLE`, `ISDELETEDRETAINED`, `ISNULL`, `ISEXCISEVOUCHER`, `EXCISETAXOVERRIDE`, `USEFORTAXUNITTRANSFER`, `ISEXER1NOPOVERWRITE`, `ISEXF2NOPOVERWRITE`, `ISEXER3NOPOVERWRITE`

### Set to `Yes`:
`ISELIGIBLEFORITC`

### State / Location:
`VATDEALERTYPE` = `Regular`  
`STATENAME` = `West Bengal`  
`COUNTRYOFRESIDENCE` = `India`  
`GSTREGISTRATIONTYPE` = `&#4; Unknown`  
`VCHGSTCLASS` = `&#4; Not Applicable`

> `&#4;` is Tally's encoding for "Not Applicable" prefix

---

## 3. Inventory Entries: ALLINVENTORYENTRIES.LIST

One block per stock item per voucher:

```
<ALLINVENTORYENTRIES.LIST>
  <STOCKITEMNAME>Raw Materials</STOCKITEMNAME>
  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
  <ISGSTASSESSABLEVALUEOVERRIDDEN>No</ISGSTASSESSABLEVALUEOVERRIDDEN>
  <STRDISGSTAPPLICABLE>No</STRDISGSTAPPLICABLE>
  <CONTENTNEGISPOS>No</CONTENTNEGISPOS>
  <AMOUNT>10000.00</AMOUNT>

  <BATCHALLOCATIONS.LIST>
    <GODOWNNAME>Main Location</GODOWNNAME>
    <BATCHNAME>Primary Batch</BATCHNAME>
    <INDENTNO>&#4; Not Applicable</INDENTNO>
    <ORDERNO>&#4; Not Applicable</ORDERNO>
    <TRACKINGNUMBER>&#4; Not Applicable</TRACKINGNUMBER>
    <AMOUNT>10000.00</AMOUNT>
  </BATCHALLOCATIONS.LIST>

  <ACCOUNTINGALLOCATIONS.LIST>
    <LEDGERNAME>Sales</LEDGERNAME>
    <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
    <AMOUNT>10000.00</AMOUNT>
    <!-- Plus all empty .LIST containers -->
  </ACCOUNTINGALLOCATIONS.LIST>

  <!-- Plus all empty .LIST containers -->
</ALLINVENTORYENTRIES.LIST>
```

### Key rules:
- `ISDEEMEDPOSITIVE="No"` for inventory (it's a credit/income)
- `AMOUNT` appears 3 times: inventory level, batch level, accounting allocation level — all same value
- `ACCOUNTINGALLOCATIONS.LIST` links the stock item to the Sales ledger
- Empty `.LIST` containers are required for Tally's parser (self-closing: `<TAG.LIST></TAG.LIST>`)

---

## 4. Ledger Entries: LEDGERENTRIES.LIST

### Party ledger (debit):

```
<LEDGERENTRIES.LIST>
  <LEDGERNAME>Shringar Fashion</LEDGERNAME>
  <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
  <AMOUNT>-11500.00</AMOUNT>          <!-- NEGATIVE total -->
  <BILLALLOCATIONS.LIST>
    <NAME>INV-001</NAME>              <!-- bill reference -->
    <BILLTYPE>New Ref</BILLTYPE>
    <AMOUNT>-11500.00</AMOUNT>
  </BILLALLOCATIONS.LIST>
  <!-- Plus all empty .LIST containers -->
</LEDGERENTRIES.LIST>
```

### Labour Charges ledger (credit):

```
<LEDGERENTRIES.LIST>
  <LEDGERNAME>Labour Charges</LEDGERNAME>
  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
  <AMOUNT>1500.00</AMOUNT>            <!-- POSITIVE -->
  <VATEXPAMOUNT>1500.00</VATEXPAMOUNT>
  <!-- Plus all empty .LIST containers -->
</LEDGERENTRIES.LIST>
```

---

## 5. Amount Sign Convention

| Entry | ISDEEMEDPOSITIVE | Amount Sign | Meaning |
|---|---|---|---|
| Party (Shringar Fashion) | `Yes` | **Negative** (-Total) | Debit — party owes this amount |
| Sales ledger (inventory) | `No` | **Positive** (+Raw Mat) | Credit — income from goods |
| Labour Charges | `No` | **Positive** (+Labour) | Credit — income from labour |

**Validation:** Sum of all AMOUNT values must = 0
```
Party(-Total) + Sales(+RawMat) + Labour(+Labour) = 0
-Total + RawMat + Labour = 0
Total = RawMat + Labour ✓
```

---

## 6. Excel to XML Field Mapping

| Excel Column | XML Location | Transformation |
|---|---|---|
| Date (DD.MM.YYYY) | `DATE`, `VCHSTATUSDATE`, `EFFECTIVEDATE` | Split by `.`, rejoin as YYYYMMDD |
| Narration | `NARRATION`, `REFERENCE`, bill `NAME` | Direct copy |
| Raw Material (amount) | `ALLINVENTORYENTRIES.LIST` > `AMOUNT` (×3) + `ACCOUNTINGALLOCATIONS` > `AMOUNT` | Direct copy (positive) |
| Raw Material | `LEDGERENTRIES.LIST` (Sales) | Handled inside inventory's ACCOUNTINGALLOCATIONS |
| Labour Charges | `LEDGERENTRIES.LIST` (Labour Charges) + `VATEXPAMOUNT` | Direct copy (positive) |
| Total Amount | `LEDGERENTRIES.LIST` (Party) + `BILLALLOCATIONS` | **Negative** of total |

---

## 7. Empty .LIST Containers Required

Tally's XML parser expects these empty list tags. Without them, import may fail:

```
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
```

---

## 8. Import in Tally

**Gateway of Tally > Import Data > Masters and Vouchers**

Tally will show a summary before importing. Check:
- Number of vouchers matches expected
- Company name is correct
- Party ledger exists in Tally masters
- All stock items and ledgers exist in Tally masters

---

## 9. Common Errors

| Symptom | Likely Cause |
|---|---|
| "Permission denied" | XLSX file open in Excel, or path is a folder |
| Tally shows 0 vouchers imported | Company name doesn't match, or ledger names don't exist |
| Party not found | `PARTYLEDGERNAME` doesn't match a ledger in Tally |
| Stock item not found | `STOCKITEMNAME` doesn't match a stock item in Tally |
| Amount mismatch | Sum of entries ≠ 0 (check sign convention) |
| Import fails silently | Missing `.LIST` containers, or encoding issue |
