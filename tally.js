#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const [, , inputCsv, ...rest] = process.argv;
const outFlag = rest.indexOf('-o');
const outputXml = outFlag !== -1 ? rest[outFlag + 1] : 'output.xml';

if (!inputCsv) {
  console.error(`
  Tally Prime XML Generator
  ─────────────────────────
  Usage:
    node tally.js sales.csv -o output.xml

  CSV format (header row required):

  date,ref,party,gstin,item,hsn,qty,rate,taxPct,godown,voucherNo
  20240401,INV-001,Acme Corp,27AABCU1234D1Z2,Widget A,8471,10,100,12,Main,INV-001
  20240401,INV-001,Acme Corp,27AABCU1234D1Z2,Widget B,8472,5,200,12,Main,INV-001
`);
  process.exit(1);
}

const csv = fs.readFileSync(inputCsv, 'utf-8').trim();
const lines = csv.split('\n');
const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
const rows = lines.slice(1).filter(l => l.trim());

const idx = (name) => {
  const i = headers.indexOf(name.toLowerCase());
  if (i === -1) throw new Error(`Column "${name}" not found in CSV`);
  return i;
};

const groupBy = (rows, refIdx) => {
  const map = new Map();
  for (const r of rows) {
    const key = r[refIdx].trim();
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(r);
  }
  return map;
};

const escXml = (s) => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

try {
  const refIdx = idx('ref');
  const grouped = groupBy(rows.map(r => r.split(',').map(c => c.trim())), refIdx);

  const vouchers = [];
  for (const [ref, items] of grouped) {
    const row = items[0];
    const date = row[idx('date')];
    const party = row[idx('party')];
    const gstin = row[idx('gstin')];
    const voucherNo = row[idx('vouchernumber')] || ref;

    let totalTax = 0;
    let totalAmount = 0;
    const invEntries = items.map(item => {
      const qty = parseFloat(item[idx('qty')]) || 0;
      const rate = parseFloat(item[idx('rate')]) || 0;
      const taxPct = parseFloat(item[idx('taxpct')]) || 0;
      const amt = qty * rate;
      const taxAmt = amt * taxPct / 100;
      totalTax += taxAmt;
      totalAmount += amt;

      return `        <INVENTORYENTRY>
          <STOCKITEMNAME>${escXml(item[idx('item')])}</STOCKITEMNAME>
          <RATE>${rate.toFixed(2)}</RATE>
          <ACTUALQTY>${qty}</ACTUALQTY>
          <AMOUNT>${amt.toFixed(2)}</AMOUNT>
          <BATCHALLOCATIONS>
            <BATCHALLOCATION>
              <GODOWNNAME>${escXml(item[idx('godown')] || 'Main Location')}</GODOWNNAME>
            </BATCHALLOCATION>
          </BATCHALLOCATIONS>
        </INVENTORYENTRY>`;
    });

    const grandTotal = totalAmount + totalTax;
    const salesLedger = `Sales @ ${items[0][idx('taxpct')]}%`;

    vouchers.push(`    <TALLYMESSAGE>
      <VOUCHER VCHTYPE="Sales" ACTION="Create">
        <DATE>${date}</DATE>
        <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
        <PARTYLEDGERNAME>${escXml(party)}</PARTYLEDGERNAME>
        <REFERENCE>${escXml(ref)}</REFERENCE>
        <VOUCHERNUMBER>${escXml(voucherNo)}</VOUCHERNUMBER>
        <PERSISTEDVIEW>Invoice</PERSISTEDVIEW>
${gstin ? `        <PARTYGSTIN>${escXml(gstin)}</PARTYGSTIN>` : ''}
        <LEDGERENTRIES>
          <LEDGERENTRY>
            <LEDGERNAME>${escXml(party)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <AMOUNT>${grandTotal.toFixed(2)}</AMOUNT>
          </LEDGERENTRY>
          <LEDGERENTRY>
            <LEDGERNAME>${escXml(salesLedger)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <AMOUNT>${totalAmount.toFixed(2)}</AMOUNT>
          </LEDGERENTRY>
          ${totalTax > 0 ? `<LEDGERENTRY>
            <LEDGERNAME>Output GST @ ${escXml(items[0][idx('taxpct')])}%</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <AMOUNT>${totalTax.toFixed(2)}</AMOUNT>
          </LEDGERENTRY>` : ''}
        </LEDGERENTRIES>
        <INVENTORYENTRIES>
${invEntries.join('\n')}
        </INVENTORYENTRIES>
      </VOUCHER>
    </TALLYMESSAGE>`);
  }

  const xml = `<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
    <TYPE>Data</TYPE>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>Your Company Name</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
${vouchers.join('\n')}
    </DATA>
  </BODY>
</ENVELOPE>`;

  fs.writeFileSync(outputXml, xml, 'utf-8');
  console.log(`✓ Generated ${outputXml} — ${vouchers.length} voucher(s) from ${rows.length} line item(s)`);
  console.log(`  Import in Tally: Gateway of Tally > Import Data > Masters and Vouchers`);
} catch (e) {
  console.error(`✗ ${e.message}`);
  process.exit(1);
}
