# Which formula when? / کون سا فارمولا کب

A short note for each sheet on which approach to pick and why. Every note has an English and an Urdu version. This file is written by hand (unlike `formulas.md`, which is generated), so edit it freely.

ہر شیٹ کے لیے مختصر رہنمائی کہ کون سا طریقہ کب بہتر ہے۔ یہ فائل ہاتھ سے لکھی گئی ہے (`formulas.md` کے برعکس جو خود بنتی ہے)، اس لیے اسے آزادی سے بدل سکتے ہیں۔

Full formulas and results: [formulas.md](formulas.md).

**Excel versions / ایکسل ورژن:** `UNIQUE`, `FILTER`, `BYROW`, `LET`, `REDUCE` and `#` spill references need Excel 365/2021. `SUM(--(...))`-style formulas need Ctrl+Shift+Enter in Excel 2019 and older; `SUMPRODUCT` does not.

<div dir="rtl" lang="ur">

`UNIQUE`، `FILTER`، `BYROW`، `LET`، `REDUCE` اور `#` والے spill ریفرنس کے لیے Excel 365/2021 چاہیے۔ `SUM(--(...))` جیسے فارمولے Excel 2019 اور پرانے ورژن میں Ctrl+Shift+Enter مانگتے ہیں؛ `SUMPRODUCT` نہیں مانگتا۔

</div>

## Sheet1: Count birthdays by year

Use `SUMPRODUCT(--(YEAR(range)=year))` for the shortest, most readable formula. On large tables `COUNTIFS` with `DATE(year,1,1)` and `DATE(year,12,31)` is usually faster, because it compares dates directly instead of extracting the year from every row.

<div dir="rtl" lang="ur">

مختصر اور آسان فارمولا چاہیے تو `SUMPRODUCT(--(YEAR(range)=year))` استعمال کریں۔ بڑے ڈیٹا میں `COUNTIFS` کے ساتھ `DATE(year,1,1)` اور `DATE(year,12,31)` عموماً زیادہ تیز رہتا ہے، کیونکہ یہ ہر قطار سے سال نکالنے کے بجائے تاریخوں کا سیدھا موازنہ کرتا ہے۔

</div>

[Formulas for this sheet](formulas.md#sheet1-count-birthdays-by-year)

## 01-P: Count cells between dates

`COUNTIFS` (and `SUMIFS` for amounts) is the first choice for "between two dates": easy to read and quick. Use `SUMPRODUCT` when the condition needs OR logic or a calculation on the column that COUNTIFS cannot do. Both limits are inclusive here (`>=` and `<=`).

<div dir="rtl" lang="ur">

دو تاریخوں کے درمیان گننے کے لیے پہلی پسند `COUNTIFS` ہے (رقم جمع کرنی ہو تو `SUMIFS`): پڑھنے میں آسان اور تیز۔ اگر شرط میں OR ہو یا کالم پر ایسا حساب کرنا ہو جو COUNTIFS نہیں کر سکتا تو `SUMPRODUCT` لیں۔ یہاں دونوں حدیں شامل ہیں (`>=` اور `<=`)۔

</div>

[Formulas for this sheet](formulas.md#01-p-count-cells-between-dates)

## 02-P: Count birthdays by year

Same two approaches as Sheet1. The new part is `UNIQUE(YEAR(...))`, which builds the list of years automatically, so new birthdays show up without retyping the years (Excel 365/2021).

<div dir="rtl" lang="ur">

وہی دو طریقے جو Sheet1 میں ہیں۔ نئی چیز `UNIQUE(YEAR(...))` ہے جو سالوں کی فہرست خود بناتی ہے، اس لیے نئی تاریخیں شامل کرنے پر سال دوبارہ ٹائپ نہیں کرنے پڑتے (Excel 365/2021)۔

</div>

[Formulas for this sheet](formulas.md#02-p-count-birthdays-by-year)

## 03-P: Count cells between two numbers and sum

Start with `COUNTIFS` and `SUMIFS`. `SUMPRODUCT` does the same job and needs no Ctrl+Shift+Enter. `SUM(FILTER(...))` is handy when you also want to see the matching rows, but it returns `#CALC!` when nothing matches. The plain `SUM((...)*(...))` versions need Ctrl+Shift+Enter in Excel 2019 and older.

<div dir="rtl" lang="ur">

پہلے `COUNTIFS` اور `SUMIFS` آزمائیں۔ `SUMPRODUCT` وہی کام کرتا ہے اور Ctrl+Shift+Enter نہیں مانگتا۔ `SUM(FILTER(...))` تب کام کا ہے جب میچ ہونے والی قطاریں بھی دیکھنی ہوں، مگر کوئی میچ نہ ہو تو `#CALC!` آتا ہے۔ سادہ `SUM((...)*(...))` والے فارمولے Excel 2019 اور پرانے ورژن میں Ctrl+Shift+Enter مانگتے ہیں۔

</div>

[Formulas for this sheet](formulas.md#03-p-count-cells-between-two-numbers-and-sum)

## 04-P: Count cells equal to

`COUNTIF(range, value)` is the simplest choice. `SUMPRODUCT(--(range=value))` gives the same result and is safer when the value can contain `*`, `?` or `~`, because COUNTIF treats those as wildcards. Both ignore upper/lower case. `UNIQUE` lists the distinct values for you.

<div dir="rtl" lang="ur">

سب سے آسان `COUNTIF(range, value)` ہے۔ اگر ویلیو میں `*`، `?` یا `~` ہو سکتا ہو تو `SUMPRODUCT(--(range=value))` بہتر ہے، کیونکہ COUNTIF انہیں وائلڈ کارڈ سمجھتا ہے۔ دونوں بڑے/چھوٹے حروف کا فرق نہیں کرتے۔ `UNIQUE` الگ الگ ویلیوز کی فہرست خود بنا دیتا ہے۔

</div>

[Formulas for this sheet](formulas.md#04-p-count-cells-equal-to)

## 05-P: Count cells equal to one of many things

`SUMPRODUCT(COUNTIF(range, list))` is the shortest. If the list can contain the same item twice, use the `MATCH` version (`ISNUMBER(MATCH(range, list, 0))`): it counts each cell once, while COUNTIF counts the cell again for every repeat. Prefer `SUMPRODUCT` over `SUM` if you do not want Ctrl+Shift+Enter in older Excel.

<div dir="rtl" lang="ur">

سب سے مختصر `SUMPRODUCT(COUNTIF(range, list))` ہے۔ اگر فہرست میں ایک ہی آئٹم دو بار آ سکتا ہو تو `MATCH` والا طریقہ (`ISNUMBER(MATCH(range, list, 0))`) لیں: وہ ہر سیل کو ایک ہی بار گنتا ہے، جبکہ COUNTIF فہرست کی ہر تکرار پر وہی سیل دوبارہ گنتا ہے۔ پرانے Excel میں Ctrl+Shift+Enter سے بچنا ہو تو `SUM` کی جگہ `SUMPRODUCT` استعمال کریں۔

</div>

[Formulas for this sheet](formulas.md#05-p-count-cells-equal-to-one-of-many-things)

## 06-P: Count cells equal to this or that

For two or three fixed values, `SUM(COUNTIF(range, {"Red","Blue"}))` is the quickest to type. If the values live in cells, use the list-based approach from 05-P. The table below it counts each distinct color with `UNIQUE` plus `SUMPRODUCT` or `SUM`.

<div dir="rtl" lang="ur">

دو تین مقررہ ویلیوز ہوں تو `SUM(COUNTIF(range, {"Red","Blue"}))` ٹائپ کرنا سب سے تیز ہے۔ اگر ویلیوز سیلز میں ہوں تو 05-P والا فہرست کا طریقہ لیں۔ نیچے والا ٹیبل `UNIQUE` اور `SUMPRODUCT`/`SUM` کے ساتھ ہر رنگ الگ گنتا ہے۔

</div>

[Formulas for this sheet](formulas.md#06-p-count-cells-equal-to-this-or-that)

## 07-P: Count cells greater than

Use `COUNTIF(range, ">"&cell)` when the threshold sits in a cell and may change, and `">90"` typed directly when it is fixed. You can also type the whole criterion (for example `>90`) in a cell and pass that cell. Switch to `COUNTIFS` when you need a second condition or both an upper and a lower limit.

<div dir="rtl" lang="ur">

حد کسی سیل میں ہو اور بدل سکتی ہو تو `COUNTIF(range, ">"&cell)` لیں، اور مستقل ہو تو سیدھا `">90"` لکھ دیں۔ پوری شرط (مثلاً `>90`) کسی سیل میں لکھ کر وہ سیل بھی دے سکتے ہیں۔ دوسری شرط ہو، یا اوپر اور نیچے دونوں حدیں ہوں، تو `COUNTIFS` پر جائیں۔

</div>

[Formulas for this sheet](formulas.md#07-p-count-cells-greater-than)

## 08-P: Count cells Less than

Same idea as 07-P with `<` and `<=`: `COUNTIF(range, "<"&cell)` for a changing threshold, a typed `"<75"` for a fixed one, and `COUNTIFS` for extra conditions or a range with two limits.

<div dir="rtl" lang="ur">

خیال وہی ہے جو 07-P میں ہے، بس `<` اور `<=` کے ساتھ: بدلتی حد کے لیے `COUNTIF(range, "<"&cell)`، مستقل حد کے لیے سیدھا `"<75"`، اور اضافی شرطوں یا دو حدوں کے لیے `COUNTIFS`۔

</div>

[Formulas for this sheet](formulas.md#08-p-count-cells-less-than)

## 09-P: Count cells not equal to many things

These formulas count the cells that are found in the list (`ISNUMBER(MATCH(...))`). For "not equal to many things", subtract that count from the total (`COUNTA`), or test `ISNA(MATCH(...))` directly. `SUMPRODUCT` avoids the Ctrl+Shift+Enter that `SUM(--(...))` needs in older Excel.

<div dir="rtl" lang="ur">

یہ فارمولے ان سیلز کو گنتے ہیں جو فہرست میں ملتے ہیں (`ISNUMBER(MATCH(...))`)۔ "کئی چیزوں کے برابر نہیں" گننے کے لیے یہ تعداد کل تعداد (`COUNTA`) سے منفی کر دیں، یا سیدھا `ISNA(MATCH(...))` استعمال کریں۔ `SUMPRODUCT` اس Ctrl+Shift+Enter سے بچاتا ہے جو پرانے Excel میں `SUM(--(...))` کو چاہیے۔

</div>

[Formulas for this sheet](formulas.md#09-p-count-cells-not-equal-to-many-things)

## 10-P: Count cells over n characters

A helper column with `LEN` plus `COUNTIF(range, ">"&n)` is the easiest to check by eye. If you do not want a helper column, `SUMPRODUCT(--(LEN(range)>n))` does it in one cell. The sheet's `SUM(N(...))` version also works but needs Ctrl+Shift+Enter in older Excel.

<div dir="rtl" lang="ur">

`LEN` کا ہیلپر کالم اور اس پر `COUNTIF(range, ">"&n)` آنکھ سے جانچنے میں سب سے آسان ہے۔ ہیلپر کالم نہ چاہیے ہو تو `SUMPRODUCT(--(LEN(range)>n))` ایک ہی سیل میں کام کر دیتا ہے۔ شیٹ کا `SUM(N(...))` والا طریقہ بھی چلتا ہے مگر پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔

</div>

[Formulas for this sheet](formulas.md#10-p-count-cells-over-n-characters)

## 11-P: Count cells that are blank

`COUNTBLANK` is the simplest. `COUNTIF(range, "")` and `SUMPRODUCT(--(range=""))` give the same answer, and all of them also count cells that hold a formula returning `""`. For blanks under another condition use `COUNTIFS(range, "", other, "A")`. A cell that holds only a space is not blank for any of them.

<div dir="rtl" lang="ur">

سب سے آسان `COUNTBLANK` ہے۔ `COUNTIF(range, "")` اور `SUMPRODUCT(--(range=""))` وہی جواب دیتے ہیں، اور یہ سب ان سیلز کو بھی گنتے ہیں جن میں `""` لوٹانے والا فارمولا ہو۔ کسی اور شرط کے ساتھ خالی سیل گننے ہوں تو `COUNTIFS(range, "", other, "A")` لیں۔ جس سیل میں صرف ایک اسپیس ہو وہ ان میں سے کسی کے لیے بھی خالی نہیں۔

</div>

[Formulas for this sheet](formulas.md#11-p-count-cells-that-are-blank)

## 12-P: Count cells that are not blank

`COUNTA` is the simplest, but it (like `COUNTIF(range, "<>")`) also counts cells holding a formula that returns `""`. `SUMPRODUCT(--(range<>""))` treats those as blank, so use it when the data comes from formulas. `COUNT` counts numbers only, so it returns 0 for text (see P8). With a second condition use `COUNTIFS(range, "<>", other, "A")`.

<div dir="rtl" lang="ur">

سب سے آسان `COUNTA` ہے، مگر یہ (اور `COUNTIF(range, "<>")` بھی) ان سیلز کو بھی گنتا ہے جن میں `""` لوٹانے والا فارمولا ہو۔ `SUMPRODUCT(--(range<>""))` انہیں خالی سمجھتا ہے، اس لیے ڈیٹا فارمولوں سے آ رہا ہو تو یہی لیں۔ `COUNT` صرف نمبر گنتا ہے، اس لیے ٹیکسٹ پر 0 دیتا ہے (P8 دیکھیں)۔ دوسری شرط ہو تو `COUNTIFS(range, "<>", other, "A")` استعمال کریں۔

</div>

[Formulas for this sheet](formulas.md#12-p-count-cells-that-are-not-blank)

## 13-P: Count cells that contain either x or y

Adding two separate counts (one per word) double counts a row that contains both words. This sheet avoids that: it tests each row for "either" and sums the TRUE rows. `FIND` is case-sensitive; use `SEARCH`, or the `COUNTIF` wildcard form (`"*blue*"`), if "Blue" and "blue" should both match.

<div dir="rtl" lang="ur">

دو الگ گنتیاں (ہر لفظ کی ایک) جمع کریں تو وہ قطار دو بار گنی جاتی ہے جس میں دونوں الفاظ ہوں۔ یہ شیٹ اس سے بچتی ہے: ہر قطار کو "یا" کے لیے جانچتی ہے اور TRUE قطاریں جمع کرتی ہے۔ `FIND` بڑے/چھوٹے حروف میں فرق کرتا ہے؛ "Blue" اور "blue" دونوں چاہیے ہوں تو `SEARCH` یا `COUNTIF` کا وائلڈ کارڈ (`"*blue*"`) لیں۔

</div>

[Formulas for this sheet](formulas.md#13-p-count-cells-that-contain-either-x-or-y)

## 14-P: Count cells that contain negative numbers

`COUNTIF(range, "<0")` is all you need. `SUM(--(range<0))` gives the same count but must be array-entered in older Excel. Zero is neither negative nor positive.

<div dir="rtl" lang="ur">

`COUNTIF(range, "<0")` کافی ہے۔ `SUM(--(range<0))` وہی تعداد دیتا ہے مگر پرانے Excel میں اسے array کے طور پر (Ctrl+Shift+Enter سے) ڈالنا پڑتا ہے۔ صفر نہ منفی ہے نہ مثبت۔

</div>

[Formulas for this sheet](formulas.md#14-p-count-cells-that-contain-negative-numbers)

## 15-P: Count cells that contain odd numbers

`ISODD` and `MOD(x,2)=1` agree on whole numbers. With decimals they differ: `ISODD` drops the decimal part first, `MOD` does not. Both return `#VALUE!` on text. The `+` in `ISODD(+range)` forces it to evaluate every cell of the range.

<div dir="rtl" lang="ur">

پورے نمبروں پر `ISODD` اور `MOD(x,2)=1` ایک جیسا نتیجہ دیتے ہیں۔ اعشاریہ والے نمبروں پر فرق آتا ہے: `ISODD` پہلے اعشاریہ حصہ ہٹا دیتا ہے، `MOD` نہیں ہٹاتا۔ ٹیکسٹ پر دونوں `#VALUE!` دیتے ہیں۔ `ISODD(+range)` میں `+` اسے رینج کے ہر سیل پر چلنے پر مجبور کرتا ہے۔

</div>

[Formulas for this sheet](formulas.md#15-p-count-cells-that-contain-odd-numbers)

## 16-P: Count cells that contain positive numbers

`COUNTIF(range, ">0")` for positive numbers, `">=0"` if zero should count too. Watch out: `SUM(--(range>=0))` also counts empty cells (an empty cell compares as 0), while `COUNTIF` does not.

<div dir="rtl" lang="ur">

مثبت نمبروں کے لیے `COUNTIF(range, ">0")`، اور صفر بھی شامل کرنا ہو تو `">=0"`۔ خبردار: `SUM(--(range>=0))` خالی سیلز کو بھی گن لیتا ہے (خالی سیل کا موازنہ 0 کی طرح ہوتا ہے)، جبکہ `COUNTIF` نہیں گنتا۔

</div>

[Formulas for this sheet](formulas.md#16-p-count-cells-that-contain-positive-numbers)

## 17-P: Count cells that contain specific text

`COUNTIF(range, "*text*")` is the shortest and ignores case. `ISNUMBER(SEARCH(...))` gives the same result and can be used inside bigger formulas. `FIND` is case-sensitive (the sheet wraps the range in `LOWER`, which only works if the search text is lowercase). If the text itself contains `*` or `?`, use `FIND`, because `SEARCH` and `COUNTIF` treat them as wildcards.

<div dir="rtl" lang="ur">

`COUNTIF(range, "*text*")` سب سے مختصر ہے اور حروف کی بڑائی چھوٹائی کا فرق نہیں کرتا۔ `ISNUMBER(SEARCH(...))` وہی نتیجہ دیتا ہے اور بڑے فارمولوں کے اندر استعمال ہو سکتا ہے۔ `FIND` میں بڑے/چھوٹے حروف کا فرق ہوتا ہے (شیٹ رینج پر `LOWER` لگاتی ہے، جو تب ہی کام کرتا ہے جب تلاش کا لفظ چھوٹے حروف میں ہو)۔ اگر لفظ میں خود `*` یا `?` ہو تو `FIND` لیں، کیونکہ `SEARCH` اور `COUNTIF` انہیں وائلڈ کارڈ سمجھتے ہیں۔

</div>

[Formulas for this sheet](formulas.md#17-p-count-cells-that-contain-specific-text)

## 18-P: Count cells that do not contain specific text

Mirror image of 17-P: `COUNTIF(range, "<>*a*")` is the shortest, `NOT(ISNUMBER(SEARCH(...)))` the most flexible. Empty cells count as "does not contain" in every version; add `range, "<>"` as a second condition in `COUNTIFS` if you want to skip them.

<div dir="rtl" lang="ur">

یہ 17-P کا الٹ ہے: `COUNTIF(range, "<>*a*")` سب سے مختصر، اور `NOT(ISNUMBER(SEARCH(...)))` سب سے لچکدار ہے۔ ہر طریقے میں خالی سیل بھی "نہیں رکھتا" میں گنے جاتے ہیں؛ انہیں چھوڑنا ہو تو `COUNTIFS` میں دوسری شرط `range, "<>"` لگا دیں۔

</div>

[Formulas for this sheet](formulas.md#18-p-count-cells-that-do-not-contain-specific-text)

## 19-P: Count cells that contain errors

`ISERROR` counts every error type; `ISERR` skips `#N/A`. To count one specific error use `ERROR.TYPE` (3 is `#VALUE!`) wrapped in `IFNA` or `IFERROR`, because `ERROR.TYPE` itself returns `#N/A` for normal cells. For plain counting, `SUMPRODUCT(--ISERROR(range))` is the shortest.

<div dir="rtl" lang="ur">

`ISERROR` ہر قسم کی error گنتا ہے؛ `ISERR` `#N/A` کو چھوڑ دیتا ہے۔ کوئی خاص error گننی ہو تو `ERROR.TYPE` کو `IFNA` یا `IFERROR` میں لپیٹ کر استعمال کریں (3 کا مطلب `#VALUE!`)، کیونکہ عام سیلز پر `ERROR.TYPE` خود `#N/A` دیتا ہے۔ سادہ گنتی کے لیے `SUMPRODUCT(--ISERROR(range))` سب سے مختصر ہے۔

</div>

[Formulas for this sheet](formulas.md#19-p-count-cells-that-contain-errors)

## 20-P: Count cells that contain errors

Same functions as 19-P, but the formulas here count the cells that are **not** errors (result 9), even though the sheet title says "contain errors". Pick between the `SUMPRODUCT` and `SUM(--...)` forms the same way as in 19-P.

<div dir="rtl" lang="ur">

فنکشنز وہی ہیں جو 19-P میں ہیں، مگر یہاں کے فارمولے ان سیلز کو گنتے ہیں جن میں error **نہیں** ہے (نتیجہ 9)، حالانکہ شیٹ کا عنوان "contain errors" ہے۔ `SUMPRODUCT` اور `SUM(--...)` میں سے انتخاب 19-P کی طرح کریں۔

</div>

[Formulas for this sheet](formulas.md#20-p-count-cells-that-contain-errors)

## 21-P: Count cells that do not contain specific text

No formulas yet (work in progress). The related-functions list (`ISNUMBER`, `SEARCH`, `MMULT`, `TRANSPOSE`, `REDUCE`) suggests the plan: count cells that do not contain any word from the "Exclude" list.

<div dir="rtl" lang="ur">

ابھی فارمولے نہیں لکھے گئے (کام جاری ہے)۔ متعلقہ فنکشنز کی فہرست (`ISNUMBER`، `SEARCH`، `MMULT`، `TRANSPOSE`، `REDUCE`) سے اندازہ ہوتا ہے کہ منصوبہ "Exclude" فہرست کے کسی بھی لفظ پر مشتمل نہ ہونے والے سیلز گننے کا ہے۔

</div>

[Formulas for this sheet](formulas.md#21-p-count-cells-that-do-not-contain-specific-text)

## 22-P: Count dates in given year

Easiest to read: a helper column `YEAR(date)` plus `COUNTIF` on it. In one cell, `SUMPRODUCT(--(YEAR(range)=year))` works. `COUNTIFS` with date limits is usually fastest on big data, but note that L8 uses `"<"&DATE(year,12,31)`, which skips 31 December. Use `"<="&DATE(year,12,31)` (or `"<"&DATE(year+1,1,1)`); the sample dates have no 31 December, so today's results are not affected. `ANCHORARRAY` is the `#` spill reference and needs Excel 365/2021.

<div dir="rtl" lang="ur">

پڑھنے میں سب سے آسان: `YEAR(date)` کا ہیلپر کالم اور اس پر `COUNTIF`۔ ایک ہی سیل میں `SUMPRODUCT(--(YEAR(range)=year))` چلتا ہے۔ بڑے ڈیٹا میں تاریخ کی حدوں والا `COUNTIFS` عموماً سب سے تیز ہے، مگر دھیان رہے کہ L8 میں `"<"&DATE(year,12,31)` ہے جو 31 دسمبر کو چھوڑ دیتا ہے۔ `"<="&DATE(year,12,31)` (یا `"<"&DATE(year+1,1,1)`) لگائیں؛ نمونہ ڈیٹا میں 31 دسمبر کی کوئی تاریخ نہیں، اس لیے موجودہ نتائج پر اثر نہیں پڑتا۔ `ANCHORARRAY` دراصل `#` والا spill ریفرنس ہے اور Excel 365/2021 مانگتا ہے۔

</div>

[Formulas for this sheet](formulas.md#22-p-count-dates-in-given-year)

## 23-P: Count if row meets internal criteria

`COUNTIFS` cannot compare one column with another, so use `SUMPRODUCT(--(C>D))`. `SUM(--(C>D))` gives the same result but needs Ctrl+Shift+Enter in older Excel.

<div dir="rtl" lang="ur">

`COUNTIFS` ایک کالم کا دوسرے کالم سے موازنہ نہیں کر سکتا، اس لیے `SUMPRODUCT(--(C>D))` استعمال کریں۔ `SUM(--(C>D))` وہی نتیجہ دیتا ہے مگر پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔

</div>

[Formulas for this sheet](formulas.md#23-p-count-if-row-meets-internal-criteria)

## 24-P: Count if row meets multiple criteria

`COUNTIFS` cannot wrap a column in `MONTH()`, so multiply the conditions inside `SUMPRODUCT`. If the month always falls in one known year, you can go back to `COUNTIFS` with date limits. The `SUM` version needs Ctrl+Shift+Enter in older Excel.

<div dir="rtl" lang="ur">

`COUNTIFS` کسی کالم پر `MONTH()` نہیں لگا سکتا، اس لیے `SUMPRODUCT` کے اندر شرطوں کو ضرب دیں۔ اگر مہینہ ہمیشہ ایک معلوم سال کا ہو تو دوبارہ تاریخ کی حدوں والے `COUNTIFS` پر آ سکتے ہیں۔ `SUM` والا طریقہ پرانے Excel میں Ctrl+Shift+Enter مانگتا ہے۔

</div>

[Formulas for this sheet](formulas.md#24-p-count-if-row-meets-multiple-criteria)

## 25-P: Count if two criteria match

`COUNTIFS` is the best choice for AND conditions: short and fast. `SUMPRODUCT` is the fallback when you need OR logic (add the conditions with `+`) or functions applied to the columns.

<div dir="rtl" lang="ur">

AND شرطوں کے لیے `COUNTIFS` سب سے بہتر ہے: مختصر اور تیز۔ OR منطق درکار ہو (شرطیں `+` سے جمع کریں) یا کالموں پر فنکشن لگانے ہوں تو `SUMPRODUCT` متبادل ہے۔

</div>

[Formulas for this sheet](formulas.md#25-p-count-if-two-criteria-match)

## 26-P: Count items in list

`COUNTIF($B$7:$B$17, B7)` filled down shows how many times each item appears. A result above 1 means the item is repeated.

<div dir="rtl" lang="ur">

`COUNTIF($B$7:$B$17, B7)` کو نیچے تک کاپی کریں تو ہر آئٹم کی تکرار کی تعداد نظر آتی ہے۔ نتیجہ 1 سے زیادہ ہو تو آئٹم دہرایا گیا ہے۔

</div>

[Formulas for this sheet](formulas.md#26-p-count-items-in-list)
