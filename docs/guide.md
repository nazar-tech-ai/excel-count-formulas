# Which formula when?

A short note for each sheet on which approach to pick and why. Every note has an English and an Urdu version. This file is written by hand (unlike `formulas.md`, which is generated), so edit it freely.

Full formulas and results: [formulas.md](formulas.md).

**Excel versions:** `UNIQUE`, `FILTER`, `BYROW`, `LET`, `REDUCE` and `#` spill references need Excel 365/2021. `SUM(--(...))`-style formulas need Ctrl+Shift+Enter in Excel 2019 and older; `SUMPRODUCT` does not.

<div align="right" dir="rtl" lang="ur">

‫ہر شیٹ کے لیے مختصر رہنمائی کہ کون سا طریقہ کب بہتر ہے۔ یہ فائل ہاتھ سے لکھی گئی ہے، جبکہ فارمولوں کی فہرست والی فائل خود بنتی ہے۔ ایکسل ورژن کے بارے میں: UNIQUE FILTER BYROW LET اور REDUCE کے لیے Excel 365 یا Excel 2021 چاہیے۔ پرانے ورژن میں SUM والے array فارمولے Ctrl+Shift+Enter مانگتے ہیں، جبکہ SUMPRODUCT نہیں مانگتا۔‬

</div>

## Sheet1: Count birthdays by year

Use `SUMPRODUCT(--(YEAR(range)=year))` for the shortest, most readable formula. On large tables `COUNTIFS` with `DATE(year,1,1)` and `DATE(year,12,31)` is usually faster, because it compares dates directly instead of extracting the year from every row.

<div align="right" dir="rtl" lang="ur">

‫مختصر اور آسان فارمولا چاہیے تو SUMPRODUCT کے ساتھ YEAR استعمال کریں۔ بڑے ڈیٹا میں COUNTIFS کے ساتھ DATE والی حدیں عموماً زیادہ تیز رہتی ہیں، کیونکہ یہ ہر قطار سے سال نکالنے کے بجائے تاریخوں کا سیدھا موازنہ کرتی ہیں۔‬

</div>

[Formulas for this sheet](formulas.md#sheet1-count-birthdays-by-year)

## 01-P: Count cells between dates

`COUNTIFS` (and `SUMIFS` for amounts) is the first choice for "between two dates": easy to read and quick. Use `SUMPRODUCT` when the condition needs OR logic or a calculation on the column that COUNTIFS cannot do. Both limits are inclusive here (`>=` and `<=`).

<div align="right" dir="rtl" lang="ur">

‫دو تاریخوں کے درمیان گننے کے لیے پہلی پسند COUNTIFS ہے، اور رقم جمع کرنی ہو تو SUMIFS۔ دونوں پڑھنے میں آسان اور تیز ہیں۔ اگر شرط میں OR والی منطق ہو، یا کالم پر ایسا حساب کرنا ہو جو COUNTIFS نہیں کر سکتا، تو SUMPRODUCT لیں۔ یہاں شروع اور آخر کی دونوں تاریخیں گنتی میں شامل ہیں۔‬

</div>

[Formulas for this sheet](formulas.md#01-p-count-cells-between-dates)

## 02-P: Count birthdays by year

Same two approaches as Sheet1. The new part is `UNIQUE(YEAR(...))`, which builds the list of years automatically, so new birthdays show up without retyping the years (Excel 365/2021).

<div align="right" dir="rtl" lang="ur">

‫طریقے وہی ہیں جو پہلی شیٹ میں ہیں۔ نئی چیز UNIQUE ہے جو سالوں کی فہرست خود بناتی ہے، اس لیے نئی تاریخیں شامل کرنے پر سال دوبارہ ٹائپ نہیں کرنے پڑتے۔ اس کے لیے Excel 365 یا Excel 2021 چاہیے۔‬

</div>

[Formulas for this sheet](formulas.md#02-p-count-birthdays-by-year)

## 03-P: Count cells between two numbers and sum

Start with `COUNTIFS` and `SUMIFS`. `SUMPRODUCT` does the same job and needs no Ctrl+Shift+Enter. `SUM(FILTER(...))` is handy when you also want to see the matching rows, but it returns `#CALC!` when nothing matches. The plain `SUM((...)*(...))` versions need Ctrl+Shift+Enter in Excel 2019 and older.

<div align="right" dir="rtl" lang="ur">

‫پہلے COUNTIFS اور SUMIFS آزمائیں۔ SUMPRODUCT وہی کام کرتا ہے اور Ctrl+Shift+Enter نہیں مانگتا۔ SUM کے ساتھ FILTER تب کام کا ہے جب میچ ہونے والی قطاریں بھی دیکھنی ہوں، مگر کوئی میچ نہ ہو تو اس میں error آ جاتی ہے۔ سادہ SUM والے array فارمولے پرانے Excel میں Ctrl+Shift+Enter مانگتے ہیں۔‬

</div>

[Formulas for this sheet](formulas.md#03-p-count-cells-between-two-numbers-and-sum)

## 04-P: Count cells equal to

`COUNTIF(range, value)` is the simplest choice. `SUMPRODUCT(--(range=value))` gives the same result and is safer when the value can contain `*`, `?` or `~`, because COUNTIF treats those as wildcards. Both ignore upper/lower case. `UNIQUE` lists the distinct values for you.

<div align="right" dir="rtl" lang="ur">

‫سب سے آسان COUNTIF ہے۔ اگر ویلیو میں ستارہ، سوالیہ نشان یا ٹلڈا جیسے خاص نشان ہو سکتے ہوں تو SUMPRODUCT بہتر ہے، کیونکہ COUNTIF انہیں وائلڈ کارڈ سمجھتا ہے۔ دونوں بڑے اور چھوٹے حروف کا فرق نہیں کرتے۔ UNIQUE الگ الگ ویلیوز کی فہرست خود بنا دیتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#04-p-count-cells-equal-to)

## 05-P: Count cells equal to one of many things

`SUMPRODUCT(COUNTIF(range, list))` is the shortest. If the list can contain the same item twice, use the `MATCH` version (`ISNUMBER(MATCH(range, list, 0))`): it counts each cell once, while COUNTIF counts the cell again for every repeat. Prefer `SUMPRODUCT` over `SUM` if you do not want Ctrl+Shift+Enter in older Excel.

<div align="right" dir="rtl" lang="ur">

‫سب سے مختصر طریقہ SUMPRODUCT کے اندر COUNTIF ہے۔ اگر فہرست میں ایک ہی آئٹم دو بار آ سکتا ہو تو MATCH والا طریقہ لیں، کیونکہ وہ ہر سیل کو ایک ہی بار گنتا ہے، جبکہ COUNTIF فہرست کی ہر تکرار پر وہی سیل دوبارہ گنتا ہے۔ پرانے Excel میں Ctrl+Shift+Enter سے بچنا ہو تو SUM کی جگہ SUMPRODUCT استعمال کریں۔‬

</div>

[Formulas for this sheet](formulas.md#05-p-count-cells-equal-to-one-of-many-things)

## 06-P: Count cells equal to this or that

For two or three fixed values, `SUM(COUNTIF(range, {"Red","Blue"}))` is the quickest to type. If the values live in cells, use the list-based approach from 05-P. The table below it counts each distinct color with `UNIQUE` plus `SUMPRODUCT` or `SUM`.

<div align="right" dir="rtl" lang="ur">

‫دو تین مقررہ ویلیوز ہوں تو انہیں براہ راست فارمولے میں لکھ کر COUNTIF پر SUM لگانا سب سے تیز ہے۔ اگر ویلیوز سیلز میں ہوں تو پانچویں شیٹ والا فہرست کا طریقہ لیں۔ نیچے والا ٹیبل UNIQUE کے ساتھ ہر رنگ الگ گنتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#06-p-count-cells-equal-to-this-or-that)

## 07-P: Count cells greater than

Use `COUNTIF(range, ">"&cell)` when the threshold sits in a cell and may change, and `">90"` typed directly when it is fixed. You can also type the whole criterion (for example `>90`) in a cell and pass that cell. Switch to `COUNTIFS` when you need a second condition or both an upper and a lower limit.

<div align="right" dir="rtl" lang="ur">

‫حد کسی سیل میں ہو اور بدل سکتی ہو تو COUNTIF میں شرط کو سیل کے ساتھ جوڑ کر دیں، اور حد مستقل ہو تو سیدھا نمبر لکھ دیں۔ پوری شرط کسی سیل میں لکھ کر وہ سیل بھی دے سکتے ہیں۔ دوسری شرط ہو، یا اوپر اور نیچے دونوں حدیں ہوں، تو COUNTIFS پر جائیں۔‬

</div>

[Formulas for this sheet](formulas.md#07-p-count-cells-greater-than)

## 08-P: Count cells Less than

Same idea as 07-P with `<` and `<=`: `COUNTIF(range, "<"&cell)` for a changing threshold, a typed `"<75"` for a fixed one, and `COUNTIFS` for extra conditions or a range with two limits.

<div align="right" dir="rtl" lang="ur">

‫خیال وہی ہے جو ساتویں شیٹ میں ہے، بس `>` کی جگہ `<`۔ بدلتی حد کے لیے سیل کے ساتھ جوڑ کر COUNTIF، مستقل حد کے لیے سیدھا نمبر، اور اضافی شرطوں کے لیے COUNTIFS۔‬

</div>

[Formulas for this sheet](formulas.md#08-p-count-cells-less-than)

## 09-P: Count cells not equal to many things

These formulas count the cells that are found in the list (`ISNUMBER(MATCH(...))`). For "not equal to many things", subtract that count from the total (`COUNTA`), or test `ISNA(MATCH(...))` directly. `SUMPRODUCT` avoids the Ctrl+Shift+Enter that `SUM(--(...))` needs in older Excel.

<div align="right" dir="rtl" lang="ur">

‫یہ فارمولے ان سیلز کو گنتے ہیں جو فہرست میں ملتے ہیں۔ کئی چیزوں کے برابر نہ ہونے والے سیل گننے کے لیے یہ تعداد کل تعداد سے منفی کر دیں، جو COUNTA سے ملتی ہے، یا سیدھا ISNA کے ساتھ MATCH استعمال کریں۔ SUMPRODUCT اس Ctrl+Shift+Enter سے بچاتا ہے جو پرانے Excel میں SUM کو چاہیے۔‬

</div>

[Formulas for this sheet](formulas.md#09-p-count-cells-not-equal-to-many-things)

## 10-P: Count cells over n characters

A helper column with `LEN` plus `COUNTIF(range, ">"&n)` is the easiest to check by eye. If you do not want a helper column, `SUMPRODUCT(--(LEN(range)>n))` does it in one cell. The sheet's `SUM(N(...))` version also works but needs Ctrl+Shift+Enter in older Excel.

<div align="right" dir="rtl" lang="ur">

‫LEN کا ہیلپر کالم اور اس پر COUNTIF آنکھ سے جانچنے میں سب سے آسان ہے۔ ہیلپر کالم نہ چاہیے ہو تو SUMPRODUCT کے اندر LEN ایک ہی سیل میں کام کر دیتا ہے۔ شیٹ کا SUM اور N والا طریقہ بھی چلتا ہے مگر پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#10-p-count-cells-over-n-characters)

## 11-P: Count cells that are blank

`COUNTBLANK` is the simplest. `COUNTIF(range, "")` and `SUMPRODUCT(--(range=""))` give the same answer, and all of them also count cells that hold a formula returning `""`. For blanks under another condition use `COUNTIFS(range, "", other, "A")`. A cell that holds only a space is not blank for any of them.

<div align="right" dir="rtl" lang="ur">

‫سب سے آسان COUNTBLANK ہے۔ COUNTIF اور SUMPRODUCT بھی وہی جواب دیتے ہیں، اور یہ سب ان سیلز کو بھی گنتے ہیں جن میں خالی متن لوٹانے والا فارمولا ہو۔ کسی اور شرط کے ساتھ خالی سیل گننے ہوں تو COUNTIFS لیں۔ جس سیل میں صرف ایک اسپیس ہو وہ ان میں سے کسی کے لیے بھی خالی نہیں۔‬

</div>

[Formulas for this sheet](formulas.md#11-p-count-cells-that-are-blank)

## 12-P: Count cells that are not blank

`COUNTA` is the simplest, but it (like `COUNTIF(range, "<>")`) also counts cells holding a formula that returns `""`. `SUMPRODUCT(--(range<>""))` treats those as blank, so use it when the data comes from formulas. `COUNT` counts numbers only, so it returns 0 for text (see P8). With a second condition use `COUNTIFS(range, "<>", other, "A")`.

<div align="right" dir="rtl" lang="ur">

‫سب سے آسان COUNTA ہے، مگر یہ ان سیلز کو بھی گنتا ہے جن میں خالی متن لوٹانے والا فارمولا ہو، اور COUNTIF کی `<>` والی شرط بھی یہی کرتی ہے۔ SUMPRODUCT انہیں خالی سمجھتا ہے، اس لیے ڈیٹا فارمولوں سے آ رہا ہو تو یہی لیں۔ COUNT صرف نمبر گنتا ہے، اس لیے ٹیکسٹ پر صفر دیتا ہے۔ دوسری شرط ہو تو COUNTIFS استعمال کریں۔‬

</div>

[Formulas for this sheet](formulas.md#12-p-count-cells-that-are-not-blank)

## 13-P: Count cells that contain either x or y

Adding two separate counts (one per word) double counts a row that contains both words. This sheet avoids that: it tests each row for "either" and sums the TRUE rows. `FIND` is case-sensitive; use `SEARCH`, or the `COUNTIF` wildcard form (`"*blue*"`), if "Blue" and "blue" should both match.

<div align="right" dir="rtl" lang="ur">

‫دو الگ گنتیاں، ہر لفظ کی ایک، جمع کریں تو وہ قطار دو بار گنی جاتی ہے جس میں دونوں الفاظ ہوں۔ یہ شیٹ اس سے بچتی ہے: ہر قطار کو "کوئی ایک لفظ موجود ہو" کے لیے جانچتی ہے اور درست قطاریں جمع کرتی ہے۔ FIND میں بڑے اور چھوٹے حروف کا فرق ہوتا ہے؛ دونوں چاہیے ہوں تو SEARCH یا COUNTIF کا وائلڈ کارڈ لیں۔‬

</div>

[Formulas for this sheet](formulas.md#13-p-count-cells-that-contain-either-x-or-y)

## 14-P: Count cells that contain negative numbers

`COUNTIF(range, "<0")` is all you need. `SUM(--(range<0))` gives the same count but must be array-entered in older Excel. Zero is neither negative nor positive.

<div align="right" dir="rtl" lang="ur">

‫COUNTIF کافی ہے۔ SUM والا array طریقہ وہی تعداد دیتا ہے مگر پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔ صفر نہ منفی ہے نہ مثبت۔‬

</div>

[Formulas for this sheet](formulas.md#14-p-count-cells-that-contain-negative-numbers)

## 15-P: Count cells that contain odd numbers

`ISODD` and `MOD(x,2)=1` agree on whole numbers. With decimals they differ: `ISODD` drops the decimal part first, `MOD` does not. Both return `#VALUE!` on text. The `+` in `ISODD(+range)` forces it to evaluate every cell of the range.

<div align="right" dir="rtl" lang="ur">

‫پورے نمبروں پر ISODD اور MOD ایک جیسا نتیجہ دیتے ہیں۔ اعشاریہ والے نمبروں پر فرق آتا ہے: ISODD پہلے اعشاریہ حصہ ہٹا دیتا ہے، MOD نہیں ہٹاتا۔ ٹیکسٹ پر دونوں error دیتے ہیں۔ ISODD کے اندر جمع کا نشان اسے رینج کے ہر سیل پر چلنے پر مجبور کرتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#15-p-count-cells-that-contain-odd-numbers)

## 16-P: Count cells that contain positive numbers

`COUNTIF(range, ">0")` for positive numbers, `">=0"` if zero should count too. Watch out: `SUM(--(range>=0))` also counts empty cells (an empty cell compares as 0), while `COUNTIF` does not.

<div align="right" dir="rtl" lang="ur">

‫مثبت نمبروں کے لیے COUNTIF میں `>0`، اور صفر بھی شامل کرنا ہو تو `>=0`۔ خبردار: SUM والا array طریقہ خالی سیلز کو بھی گن لیتا ہے، کیونکہ خالی سیل کا موازنہ صفر کی طرح ہوتا ہے، جبکہ COUNTIF نہیں گنتا۔‬

</div>

[Formulas for this sheet](formulas.md#16-p-count-cells-that-contain-positive-numbers)

## 17-P: Count cells that contain specific text

`COUNTIF(range, "*text*")` is the shortest and ignores case. `ISNUMBER(SEARCH(...))` gives the same result and can be used inside bigger formulas. `FIND` is case-sensitive (the sheet wraps the range in `LOWER`, which only works if the search text is lowercase). If the text itself contains `*` or `?`, use `FIND`, because `SEARCH` and `COUNTIF` treat them as wildcards.

<div align="right" dir="rtl" lang="ur">

‫COUNTIF کے ساتھ ستارے والا پیٹرن سب سے مختصر ہے اور حروف کی بڑائی چھوٹائی کا فرق نہیں کرتا۔ SEARCH کے ساتھ ISNUMBER وہی نتیجہ دیتا ہے اور بڑے فارمولوں کے اندر استعمال ہو سکتا ہے۔ FIND میں بڑے اور چھوٹے حروف کا فرق ہوتا ہے۔ اگر لفظ میں خود ستارہ یا سوالیہ نشان ہو تو FIND لیں، کیونکہ SEARCH اور COUNTIF انہیں وائلڈ کارڈ سمجھتے ہیں۔‬

</div>

[Formulas for this sheet](formulas.md#17-p-count-cells-that-contain-specific-text)

## 18-P: Count cells that do not contain specific text

Mirror image of 17-P: `COUNTIF(range, "<>*a*")` is the shortest, `NOT(ISNUMBER(SEARCH(...)))` the most flexible. Empty cells count as "does not contain" in every version; add `range, "<>"` as a second condition in `COUNTIFS` if you want to skip them.

<div align="right" dir="rtl" lang="ur">

‫یہ سترہویں شیٹ کا الٹ ہے: COUNTIF سب سے مختصر، اور NOT کے ساتھ ISNUMBER اور SEARCH سب سے لچکدار ہے۔ ہر طریقے میں خالی سیل بھی "متن موجود نہیں" والوں میں گنے جاتے ہیں؛ انہیں چھوڑنا ہو تو COUNTIFS میں دوسری شرط لگا دیں۔‬

</div>

[Formulas for this sheet](formulas.md#18-p-count-cells-that-do-not-contain-specific-text)

## 19-P: Count cells that contain errors

`ISERROR` counts every error type; `ISERR` skips `#N/A`. To count one specific error use `ERROR.TYPE` (3 is `#VALUE!`) wrapped in `IFNA` or `IFERROR`, because `ERROR.TYPE` itself returns `#N/A` for normal cells. For plain counting, `SUMPRODUCT(--ISERROR(range))` is the shortest.

<div align="right" dir="rtl" lang="ur">

‫ISERROR ہر قسم کی error گنتا ہے؛ ISERR صرف N/A والی error کو چھوڑ دیتا ہے۔ کوئی خاص error گننی ہو تو ERROR.TYPE کو IFNA یا IFERROR میں لپیٹ کر استعمال کریں، کیونکہ عام سیلز پر ERROR.TYPE خود N/A دیتا ہے۔ سادہ گنتی کے لیے SUMPRODUCT کے ساتھ ISERROR سب سے مختصر ہے۔‬

</div>

[Formulas for this sheet](formulas.md#19-p-count-cells-that-contain-errors)

## 20-P: Count cells that contain errors

Same functions as 19-P, but the formulas here count the cells that are **not** errors (result 9), even though the sheet title says "contain errors". Pick between the `SUMPRODUCT` and `SUM(--...)` forms the same way as in 19-P.

<div align="right" dir="rtl" lang="ur">

‫فنکشنز وہی ہیں جو انیسویں شیٹ میں ہیں، مگر یہاں کے فارمولے ان سیلز کو گنتے ہیں جن میں error نہیں ہے، حالانکہ شیٹ کا عنوان errors والے سیلز کا ہے۔ SUMPRODUCT اور SUM میں سے انتخاب انیسویں شیٹ کی طرح کریں۔‬

</div>

[Formulas for this sheet](formulas.md#20-p-count-cells-that-contain-errors)

## 21-P: Count cells that do not contain specific text

No formulas yet (work in progress). The related-functions list (`ISNUMBER`, `SEARCH`, `MMULT`, `TRANSPOSE`, `REDUCE`) suggests the plan: count cells that do not contain any word from the "Exclude" list.

<div align="right" dir="rtl" lang="ur">

‫ابھی فارمولے نہیں لکھے گئے، کام جاری ہے۔ متعلقہ فنکشنز سے اندازہ ہوتا ہے کہ منصوبہ Exclude فہرست کے کسی بھی لفظ پر مشتمل نہ ہونے والے سیلز گننے کا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#21-p-count-cells-that-do-not-contain-specific-text)

## 22-P: Count dates in given year

Easiest to read: a helper column `YEAR(date)` plus `COUNTIF` on it. In one cell, `SUMPRODUCT(--(YEAR(range)=year))` works. `COUNTIFS` with date limits is usually fastest on big data, but note that L8 uses `"<"&DATE(year,12,31)`, which skips 31 December. Use `"<="&DATE(year,12,31)` (or `"<"&DATE(year+1,1,1)`); the sample dates have no 31 December, so today's results are not affected. `ANCHORARRAY` is the `#` spill reference and needs Excel 365/2021.

<div align="right" dir="rtl" lang="ur">

‫پڑھنے میں سب سے آسان طریقہ YEAR کا ہیلپر کالم اور اس پر COUNTIF ہے۔ ایک ہی سیل میں SUMPRODUCT کے ساتھ YEAR چلتا ہے۔ بڑے ڈیٹا میں تاریخ کی حدوں والا COUNTIFS عموماً سب سے تیز ہے، مگر دھیان رہے کہ سیل L8 میں آخری حد کے لیے چھوٹے کا نشان (`<`) استعمال ہوا ہے، جس سے 31 دسمبر رہ جاتا ہے۔ اسے `<=` کر دیں۔ نمونہ ڈیٹا میں 31 دسمبر کی کوئی تاریخ نہیں، اس لیے موجودہ نتائج پر اثر نہیں پڑتا۔ ANCHORARRAY دراصل ہیش والا spill ریفرنس ہے اور Excel 365 یا Excel 2021 مانگتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#22-p-count-dates-in-given-year)

## 23-P: Count if row meets internal criteria

`COUNTIFS` cannot compare one column with another, so use `SUMPRODUCT(--(C>D))`. `SUM(--(C>D))` gives the same result but needs Ctrl+Shift+Enter in older Excel.

<div align="right" dir="rtl" lang="ur">

‫COUNTIFS ایک کالم کا دوسرے کالم سے موازنہ نہیں کر سکتا، اس لیے SUMPRODUCT استعمال کریں۔ SUM والا طریقہ وہی نتیجہ دیتا ہے مگر پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#23-p-count-if-row-meets-internal-criteria)

## 24-P: Count if row meets multiple criteria

`COUNTIFS` cannot wrap a column in `MONTH()`, so multiply the conditions inside `SUMPRODUCT`. If the month always falls in one known year, you can go back to `COUNTIFS` with date limits. The `SUM` version needs Ctrl+Shift+Enter in older Excel.

<div align="right" dir="rtl" lang="ur">

‫COUNTIFS کسی کالم پر MONTH نہیں لگا سکتا، اس لیے SUMPRODUCT کے اندر شرطوں کو ضرب دیں۔ اگر مہینہ ہمیشہ ایک معلوم سال کا ہو تو دوبارہ تاریخ کی حدوں والے COUNTIFS پر آ سکتے ہیں۔ SUM والا طریقہ پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#24-p-count-if-row-meets-multiple-criteria)

## 25-P: Count if two criteria match

`COUNTIFS` is the best choice for AND conditions: short and fast. `SUMPRODUCT` is the fallback when you need OR logic (add the conditions with `+`) or functions applied to the columns.

<div align="right" dir="rtl" lang="ur">

‫AND والی شرطوں کے لیے COUNTIFS سب سے بہتر ہے: مختصر اور تیز۔ OR والی منطق چاہیے ہو یا کالموں پر فنکشن لگانے ہوں تو SUMPRODUCT متبادل ہے۔‬

</div>

[Formulas for this sheet](formulas.md#25-p-count-if-two-criteria-match)

## 26-P: Count items in list

`COUNTIF($B$7:$B$17, B7)` filled down shows how many times each item appears. A result above 1 means the item is repeated.

<div align="right" dir="rtl" lang="ur">

‫COUNTIF کو نیچے تک کاپی کریں تو ہر آئٹم کی تکرار کی تعداد نظر آتی ہے۔ نتیجہ ایک سے زیادہ ہو تو آئٹم دہرایا گیا ہے۔‬

</div>

[Formulas for this sheet](formulas.md#26-p-count-items-in-list)
