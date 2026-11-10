# Excel COUNT Formulas: Practice Workbook

A hands-on collection of Excel formulas for **counting**: one sheet per problem, each solved
in several ways (`COUNTIF`, `COUNTIFS`, `SUMPRODUCT`, array formulas, `UNIQUE`, `FILTER`, `MATCH`, `LEN`, `SEARCH` and more).

Every sheet shows the data, the formulas side by side, and the formula text (via `FORMULATEXT`) so you can
compare approaches and pick the one that fits.

## What is inside

```
excel-count-formulas/
├── workbook/
│   └── Multiples_Examples_of_Count.xlsx   # the practice file (27 sheets)
├── docs/
│   ├── guide.md                           # which formula to pick when (English + Urdu)
│   ├── formulas.md                        # every formula, grouped by sheet
│   └── formulas.csv                       # same data, filterable in Excel
├── scripts/
│   └── extract_formulas.py                # regenerates docs/ from the workbook
├── requirements.txt
└── README.md
```

## Sheets

<!-- SHEETS:START -->
| # | Sheet | Topic | Main functions | Formulas |
| --- | --- | --- | --- | --- |
| 0 | [Sheet1](docs/formulas.md#sheet1-count-birthdays-by-year) | Count birthdays by year | COUNTIFS, DATE, SUMPRODUCT, YEAR | 8 |
| 1 | [01-P](docs/formulas.md#01-p-count-cells-between-dates) | Count cells between dates | SUMPRODUCT, COUNTIFS | 4 |
| 2 | [02-P](docs/formulas.md#02-p-count-birthdays-by-year) | Count birthdays by year | SUMPRODUCT, YEAR, COUNTIFS, BYROW, UNIQUE, SORT, LET | 10 |
| 3 | [03-P](docs/formulas.md#03-p-count-cells-between-two-numbers-and-sum) | Count cells between two numbers and sum | SUMPRODUCT, COUNTIFS | 21 |
| 4 | [04-P](docs/formulas.md#04-p-count-cells-equal-to) | Count cells equal to | SUMPRODUCT, COUNTIF | 14 |
| 5 | [05-P](docs/formulas.md#05-p-count-cells-equal-to-one-of-many-things) | Count cells equal to one of many things | SUMPRODUCT, COUNTIF, ISNUMBER, MATCH | 5 |
| 6 | [06-P](docs/formulas.md#06-p-count-cells-equal-to-this-or-that) | Count cells equal to this or that | SUMPRODUCT, COUNTIF, SUM | 10 |
| 7 | [07-P](docs/formulas.md#07-p-count-cells-greater-than) | Count cells greater than | COUNTIF, COUNTIFS | 6 |
| 8 | [08-P](docs/formulas.md#08-p-count-cells-less-than) | Count cells Less than | COUNTIF, COUNTIFS | 6 |
| 9 | [09-P](docs/formulas.md#09-p-count-cells-not-equal-to-many-things) | Count cells not equal to many things | COUNTIF, COUNTA, SUMPRODUCT, MATCH, ISNA | 4 |
| 10 | [10-P](docs/formulas.md#10-p-count-cells-over-n-characters) | Count cells over n characters | SUMPRODUCT, LEN | 30 |
| 11 | [11-P](docs/formulas.md#11-p-count-cells-that-are-blank) | Count cells that are blank | COUNTBLANK, COUNTA, COUNTIF, SUMPRODUCT | 10 |
| 12 | [12-P](docs/formulas.md#12-p-count-cells-that-are-not-blank) | Count cells that are not blank | COUNTBLANK, COUNTA, COUNTIFS, COUNTIF, SUMPRODUCT | 16 |
| 13 | [13-P](docs/formulas.md#13-p-count-cells-that-contain-either-x-or-y) | Count cells that contain either x or y | SUMPRODUCT, COUNTIF, ISNUMBER, FIND | 13 |
| 14 | [14-P](docs/formulas.md#14-p-count-cells-that-contain-negative-numbers) | Count cells that contain negative numbers | SUMPRODUCT, COUNTIF | 3 |
| 15 | [15-P](docs/formulas.md#15-p-count-cells-that-contain-odd-numbers) | Count cells that contain odd numbers | SUMPRODUCT, ISODD, MOD, ISEVEN | 6 |
| 16 | [16-P](docs/formulas.md#16-p-count-cells-that-contain-positive-numbers) | Count cells that contain positive numbers | SUMPRODUCT, COUNTIF | 4 |
| 17 | [17-P](docs/formulas.md#17-p-count-cells-that-contain-specific-text) | Count cells that contain specific text | SUMPRODUCT, COUNTIF, SEARCH, FIND, ISNUMBER | 14 |
| 18 | [18-P](docs/formulas.md#18-p-count-cells-that-do-not-contain-specific-text) | Count cells that do not contain specific text | SUMPRODUCT, COUNTIF, NOT, SEARCH, FIND, ISNUMBER | 11 |
| 19 | [19-P](docs/formulas.md#19-p-count-cells-that-contain-errors) | Count cells that contain errors | ERROR.TYPE, FIND, IFERROR, IFNA | 7 |
| 20 | [20-P](docs/formulas.md#20-p-count-cells-that-contain-errors) | Count cells that contain errors | ERROR.TYPE, FIND, IFNA, ISERROR | 5 |
| 21 | [21-P](docs/formulas.md#21-p-count-cells-that-do-not-contain-specific-text) | Count cells that do not contain specific text | ISNUMBER, SEARCH, MMULT, TRANSPOSE, REDUCE | 0 |
| 22 | [22-P](docs/formulas.md#22-p-count-dates-in-given-year) | Count dates in given year | SUMPRODUCT | 27 |
| 23 | [23-P](docs/formulas.md#23-p-count-if-row-meets-internal-criteria) | Count if row meets internal criteria | SUMPRODUCT | 6 |
| 24 | [24-P](docs/formulas.md#24-p-count-if-row-meets-multiple-criteria) | Count if row meets multiple criteria | MONTH, SUM, SUMPRODUCT | 2 |
| 25 | [25-P](docs/formulas.md#25-p-count-if-two-criteria-match) | Count if two criteria match | COUNTIFS, SUMPRODUCT | 6 |
| 26 | [26-P](docs/formulas.md#26-p-count-items-in-list) | Count items in list | COUNTIF | 11 |
<!-- SHEETS:END -->

Full formula text and results for each sheet are in [docs/formulas.md](docs/formulas.md).
Not sure which approach to use? [docs/guide.md](docs/guide.md) explains, sheet by sheet, which formula is better when (English and Urdu).

Urdu readers: see [Urdu summary](#urdu-summary) below.

## How to use

1. Download `workbook/Multiples_Examples_of_Count.xlsx` and open it in Excel.
2. Open a sheet, read the title, then compare the formulas in the **Functions** column.
3. Change the sample data and watch each formula's result change.

## Requirements

- **Excel 365 or Excel 2021** for the sheets that use dynamic array functions
  (`UNIQUE`, `FILTER`, spilled ranges with `#`). Older versions show `#NAME?` on those cells.
- `FORMULATEXT` needs Excel 2013 or later.
- Google Sheets and LibreOffice work for most `COUNTIF`/`SUMPRODUCT` examples, but some array formulas behave differently.

## Regenerate the docs

After you add or change sheets in the workbook:

```bash
pip install -r requirements.txt
python scripts/extract_formulas.py
```

This rewrites `docs/formulas.md`, `docs/formulas.csv` and the sheet table above.

## Adding a new example

Follow the existing layout so the script can read it:

- Put the **title** in the first rows of the sheet (for example `Count cells greater than`).
- List the functions under a cell labelled `Related Functions`.
- Name each formula approach clearly and keep the sample data next to it.
- Name the sheet with the next number, for example `27-P`.
- Add a short note for the new sheet to `docs/guide.md` (this file is written by hand, the script does not touch it).

## Notes

- The catalog shows results saved by Excel the last time the workbook was saved. Open and save
  the file in Excel before running the script if you changed data.
- Sheet `21-P` has the sample data and an "Exclude" list but no formulas yet (work in progress).

## Urdu summary

<div dir="rtl" lang="ur">

یہ ریپو ایکسل میں گننے کے فارمولوں کی پریکٹس ورک بک ہے۔ ہر شیٹ میں ایک مسئلہ ہے اور اسے کئی طریقوں سے حل کیا گیا ہے۔ ہر شیٹ میں ڈیٹا اور فارمولے ساتھ ساتھ موجود ہیں، تاکہ آپ طریقوں کا موازنہ کر کے اپنے کام کا طریقہ چن سکیں۔

**استعمال کا طریقہ**

1. ورک بک والے فولڈر سے ایکسل فائل ڈاؤن لوڈ کر کے ایکسل میں کھولیں۔
2. کوئی شیٹ کھولیں، عنوان پڑھیں، اور فارمولوں کا موازنہ کریں۔
3. نمونہ ڈیٹا بدلیں اور دیکھیں کہ ہر فارمولے کا نتیجہ کیسے بدلتا ہے۔

**مزید رہنمائی**

- کون سا فارمولا کب بہتر ہے، اردو اور انگریزی دونوں میں: [docs/guide.md](docs/guide.md)
- تمام فارمولے اور ان کے نتائج: [docs/formulas.md](docs/formulas.md)

**ایکسل ورژن کی شرطیں**

- UNIQUE FILTER BYROW اور ہیش والے spill ریفرنس کے لیے Excel 365 یا Excel 2021 چاہیے۔ پرانے ورژن میں ان سیلز میں نام کی error آتی ہے۔
- FORMULATEXT کے لیے Excel 2013 یا نیا ورژن چاہیے۔
- Google Sheets اور LibreOffice میں زیادہ تر مثالیں چلتی ہیں، مگر کچھ array فارمولے مختلف برتاؤ کرتے ہیں۔

**نوٹ**

- ورک بک میں شیٹ شامل کرنے یا بدلنے کے بعد docs دوبارہ بنانے کا طریقہ اوپر انگریزی حصے میں Regenerate the docs کے تحت لکھا ہے۔
- ایک شیٹ میں ڈیٹا اور Exclude فہرست ہے مگر فارمولے ابھی نہیں لکھے گئے۔ تفصیل انگریزی Notes میں ہے۔

</div>
