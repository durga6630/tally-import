# tally-generator

Tally Prime Sales Voucher XML Generator

Generate Tally-formatted XML from Excel data — auto-matches your existing Tally export format.

## Usage

### 1. Generate a template

```bash
python tally_generator.py template
```

Creates `sample.xlsx` with headers: Date, Narration, Raw Material, Labour Charges, Total Amount.

### 2. Fill your data in the XLSX

### 3. Generate the XML

```bash
python tally_generator.py your_data.xlsx
```

Output: `your_data_tally.xml`

### Config

```bash
python tally_generator.py config
```

Edits `tally_config.json` to set your company name, party, ledger names, godown, etc.

### Full example

```bash
python tally_generator.py template sales_data.xlsx
# ... fill sales_data.xlsx ...
python tally_generator.py sales_data.xlsx
```

Import in Tally: Gateway of Tally > Import Data > Masters and Vouchers
