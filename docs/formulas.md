# Formula catalog

Generated from `Multiples_Examples_of_Count.xlsx` by `scripts/extract_formulas.py`. Do not edit by hand; re-run the script instead.

Formulas that are filled down a column (same pattern, different row) are shown once with a cell count.

## Functions used across the workbook

| Function | Sheets using it |
| --- | --- |
| `SUM` | 20 |
| `SUMPRODUCT` | 16 |
| `COUNTIF` | 15 |
| `COUNTIFS` | 8 |
| `FIND` | 5 |
| `ISNUMBER` | 5 |
| `N` | 4 |
| `DATE` | 3 |
| `LEN` | 3 |
| `NOT` | 3 |
| `UNIQUE` | 3 |
| `YEAR` | 3 |
| `ERROR.TYPE` | 2 |
| `IFNA` | 2 |
| `ISERROR` | 2 |
| `LOWER` | 2 |
| `MATCH` | 2 |
| `NA` | 2 |
| `SEARCH` | 2 |
| `SUMIFS` | 2 |
| `ANCHORARRAY` | 1 |
| `COUNT` | 1 |
| `COUNTA` | 1 |
| `COUNTBLANK` | 1 |
| `FILTER` | 1 |
| `IF` | 1 |
| `IFERROR` | 1 |
| `ISERR` | 1 |
| `ISEVEN` | 1 |
| `ISODD` | 1 |
| `MOD` | 1 |
| `MONTH` | 1 |

## Sheet1: Count birthdays by year

**Functions used:** `COUNTIFS`, `DATE`, `SUMPRODUCT`, `YEAR`  
**Excel tables:** `data`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| F5 | array | `=SUMPRODUCT(--(YEAR(data[Birthday])=E5))` | `3` | 4 |
| O5 | normal | `=COUNTIFS(data[Birthday],">="&DATE(E5,1,1),data[Birthday],"<="&DATE(E5,12,31))` | `3` | 4 |

## 01-P: Count cells between dates

**Related functions:** `SUMPRODUCT`, `COUNTIFS`  
**Functions used:** `COUNTIFS`, `SUMIFS`, `SUMPRODUCT`  
**Excel tables:** `Data1`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| H8 | normal | `=COUNTIFS(Data1[Birthday],">="&H6,Data1[Birthday],"<="&H7)` | `6` | 1 |
| H9 | normal | `=SUMIFS(Data1[Amount],Data1[Birthday],">="&H6,Data1[Birthday],"<="&H7)` | `3898` | 1 |
| H11 | array | `=SUMPRODUCT((Data1[Birthday]>=H6)*(Data1[Birthday]<=H7))` | `6` | 1 |
| H12 | array | `=SUMPRODUCT((Data1[Birthday]>=H6)*(Data1[Birthday]<=H7)*Data1[Amount])` | `3898` | 1 |

## 02-P: Count birthdays by year

**Related functions:** `SUMPRODUCT`, `YEAR`, `COUNTIFS`, `BYROW`, `UNIQUE`, `SORT`, `LET`  
**Functions used:** `COUNTIFS`, `DATE`, `SUMPRODUCT`, `UNIQUE`, `YEAR`  
**Excel tables:** `Data2`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| F8:F11 | array | `=UNIQUE(YEAR(Data2[Birthday]))` | `1999` | 2 |
| G8 | array | `=SUMPRODUCT(--(YEAR(Data2[Birthday])=F8))` | `3` | 4 |
| G17 | normal | `=COUNTIFS(Data2[Birthday],">="&DATE(F17,1,1),Data2[Birthday],"<="&DATE(F17,12,31))` | `3` | 4 |

## 03-P: Count cells between two numbers and sum

**Related functions:** `SUMPRODUCT`, `COUNTIFS`  
**Functions used:** `COUNTIFS`, `FILTER`, `SUM`, `SUMIFS`, `SUMPRODUCT`  
**Excel tables:** `Data3`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| H8 | array | `=SUMPRODUCT((Data3[Score]>=F8)*(Data3[Score]<=G8))` | `2` | 3 |
| I8 | array | `=SUMPRODUCT((Score>=F8)*(Score<=G8)*Score)` | `147` | 3 |
| J8 | array | `=SUM(FILTER(Data3[Score],((Data3[Score]>=F8)*(Data3[Score]<=G8))))` | `147` | 3 |
| H20 | normal | `=COUNTIFS(Data3[Score],">="&F20,Data3[Score],"<="&G20)` | `2` | 3 |
| I20 | array | `=SUM((Score>=F20)*(Score<=G20))` | `2` | 3 |
| J20 | normal | `=SUMIFS(Data3[Score],Data3[Score],">="&F20,Data3[Score],"<="&G20)` | `147` | 3 |
| K20 | array | `=SUM((Score>=F20)*(Score<=G20)*Data3[Score])` | `147` | 3 |

## 04-P: Count cells equal to

**Related functions:** `SUMPRODUCT`, `COUNTIF`  
**Functions used:** `COUNTIF`, `SUM`, `SUMPRODUCT`, `UNIQUE`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| H8:H11 | array | `=UNIQUE(D8:D19)` | `Red` | 2 |
| I8 | normal | `=COUNTIF(Color,H8)` | `5` | 4 |
| I18 | array | `=SUMPRODUCT(--(Color=H18))` | `5` | 4 |
| J18 | array | `=SUM(--(Color=H18))` | `5` | 4 |

## 05-P: Count cells equal to one of many things

**Related functions:** `SUMPRODUCT`, `COUNTIF`, `ISNUMBER`, `MATCH`  
**Functions used:** `COUNTIF`, `ISNUMBER`, `MATCH`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| D8:D19 | array | `=MATCH(B8:B19,Thinges,0)` | `1` | 1 |
| H10 | array | `=SUM(COUNTIF(B8:B19,Thinges))` | `7` | 1 |
| H11 | array | `=SUMPRODUCT(COUNTIF(B8:B19,Thinges))` | `7` | 1 |
| H12 | array | `=SUM(--(ISNUMBER(MATCH(B8:B19,Thinges,0))))` | `7` | 1 |
| H13 | array | `=SUMPRODUCT(--(ISNUMBER(MATCH(B8:B19,Thinges,0))))` | `7` | 1 |

## 06-P: Count cells equal to this or that

**Related functions:** `SUMPRODUCT`, `COUNTIF`, `SUM`  
**Functions used:** `COUNTIF`, `SUM`, `SUMPRODUCT`, `UNIQUE`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| J8 | array | `=SUM(COUNTIF(Color,{"Red","Blue"}))` | `8` | 1 |
| H18:H21 | array | `=UNIQUE(D8:D19)` | `Red` | 1 |
| I18 | array | `=SUMPRODUCT(--(Color=H18))` | `5` | 4 |
| J18 | array | `=SUM(--(Color=H18))` | `5` | 4 |

## 07-P: Count cells greater than

**Related functions:** `COUNTIF`, `COUNTIFS`  
**Functions used:** `COUNTIF`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| G7 | normal | `=COUNTIF(C7:C18,F7)` | `3` | 2 |
| G15 | normal | `=COUNTIF(C7:C18,">"&F15)` | `3` | 1 |
| H15 | normal | `=COUNTIF(C7:C18,">90")` | `3` | 1 |
| G16 | normal | `=COUNTIF(C7:C18,">="&F16)` | `4` | 1 |
| H16 | normal | `=COUNTIF(C7:C18,">=90")` | `4` | 1 |

## 08-P: Count cells Less than

**Related functions:** `COUNTIF`, `COUNTIFS`  
**Functions used:** `COUNTIF`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| G7 | normal | `=COUNTIF(C7:C18,F7)` | `3` | 2 |
| G15 | normal | `=COUNTIF(C7:C18,"<"&F15)` | `3` | 1 |
| H15 | normal | `=COUNTIF(C7:C18,"<75")` | `3` | 1 |
| G16 | normal | `=COUNTIF(C7:C18,"<="&F16)` | `4` | 1 |
| H16 | normal | `=COUNTIF(C7:C18,"<=75")` | `4` | 1 |

## 09-P: Count cells not equal to many things

**Related functions:** `COUNTIF`, `COUNTA`, `SUMPRODUCT`, `MATCH`, `ISNA`  
**Functions used:** `ISNUMBER`, `MATCH`, `N`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| D7:D18 | array | `=N(ISNUMBER(MATCH(B7:B18,F15:F17,0)))` | `1` | 1 |
| G7 | array | `=SUM(--(ISNUMBER(MATCH(B7:B18,F7:F9,0))))` | `9` | 1 |
| H7 | array | `=SUM(N(ISNUMBER(MATCH(B7:B18,F15:F17,0))))` | `9` | 1 |
| G15 | array | `=SUMPRODUCT(--(ISNUMBER(MATCH(B7:B18,F15:F17,0))))` | `9` | 1 |

## 10-P: Count cells over n characters

**Related functions:** `SUMPRODUCT`, `LEN`  
**Functions used:** `COUNTIF`, `LEN`, `N`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7 | normal | `=LEN(B7)` | `29` | 27 |
| D7:D33 | array | `=N(LEN(B7:B33)>G6)` | `0` | 1 |
| G7 | array | `=SUM(N(LEN(B7:B33)>G6))` | `9` | 1 |
| G9 | normal | `=COUNTIF(C7:C33,">"&G6)` | `9` | 1 |

## 11-P: Count cells that are blank

**Related functions:** `COUNTBLANK`, `COUNTA`, `COUNTIF`, `SUMPRODUCT`  
**Functions used:** `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `LEN`, `N`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| G15 | normal | `=COUNTBLANK(C7:C21)` | `6` | 1 |
| H15 | normal | `=COUNTIF(C7:C21,"")` | `6` | 1 |
| I15 | array | `=SUM(--(C7:C21=""))` | `6` | 1 |
| J15 | array | `=SUMPRODUCT(--(C7:C21=""))` | `6` | 1 |
| K15 | array | `=SUMPRODUCT(N(C7:C21=""))` | `6` | 1 |
| L15 | array | `=SUM(N(LEN(C7:C21)=0))` | `6` | 1 |
| G16 | normal | `=COUNTIFS(C7:C21,"",B7:B21,"A")` | `2` | 1 |
| H16 | array | `=SUM(--(C7:C21="")*(B7:B21="A"))` | `2` | 1 |
| I16 | array | `=SUMPRODUCT(N(C7:C21="")*(B7:B21="A"))` | `2` | 1 |
| J16 | array | `=SUM(--(LEN(C7:C21)=0)*(B7:B21="A"))` | `2` | 1 |

## 12-P: Count cells that are not blank

**Related functions:** `COUNTBLANK`, `COUNTA`, `COUNTIFS`, `COUNTIF`, `SUMPRODUCT`  
**Functions used:** `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `LEN`, `N`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| P6 | normal | `=COUNTIF(O6:O9,"<>")` | `2` | 1 |
| Q6:Q11 | array | `=LEN(O5:O10)` | `0` | 1 |
| N7 | normal | `=COUNTA(O4:O9)` | `2` | 1 |
| P8 | normal | `=COUNT(O5:O10)` | `0` | 1 |
| P9 | array | `=SUM(--(O5:O10<>0))` | `2` | 1 |
| G15 | normal | `=COUNTA(C7:C21)` | `9` | 1 |
| H15 | normal | `=COUNTIF(C7:C21,"<>")` | `9` | 1 |
| I15 | array | `=SUM(--(C7:C21<>""))` | `9` | 1 |
| J15 | array | `=SUMPRODUCT(--(C7:C21<>""))` | `9` | 1 |
| K15 | array | `=SUMPRODUCT(N(C7:C21<>""))` | `9` | 1 |
| L15 | array | `=SUM(N(LEN(C7:C21)<>0))` | `9` | 1 |
| G16 | normal | `=COUNTIFS(C7:C21,"<>",B7:B21,"A")` | `4` | 1 |
| H16 | array | `=SUM(--(C7:C21<>"")*(B7:B21="A"))` | `4` | 1 |
| I16 | array | `=SUMPRODUCT(N(C7:C21<>"")*(B7:B21="A"))` | `4` | 1 |
| J16 | array | `=SUM(--(LEN(C7:C21)<>0)*(B7:B21="A"))` | `4` | 1 |
| K16 | array | `=SUM(--(LEN(C7:C21)>0)*(B7:B21="A"))` | `4` | 1 |

## 13-P: Count cells that contain either x or y

**Related functions:** `SUMPRODUCT`, `COUNTIF`, `ISNUMBER`, `FIND`  
**Functions used:** `COUNTIF`, `FIND`, `ISNUMBER`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7 | array | `=SUM(--(SUM(COUNTIF(B7,{"*blue*","*green*"}))>0))` | `1` | 11 |
| G8 | normal | `=SUM(C7:C17)` | `7` | 1 |
| H8 | array | `=SUM(--(ISNUMBER(FIND("green",$B$7:$B$17,1))+ISNUMBER(FIND("blue",$B$7:$B$17,1))>0))` | `7` | 1 |

## 14-P: Count cells that contain negative numbers

**Related functions:** `SUMPRODUCT`, `COUNTIF`  
**Functions used:** `COUNTIF`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| G8 | normal | `=COUNTIF(B7:B17,"<"&0)` | `3` | 1 |
| H8 | array | `=SUM(--(B7:B17<0))` | `3` | 1 |
| I8 | normal | `=COUNTIF(B7:B17,"<0")` | `3` | 1 |

## 15-P: Count cells that contain odd numbers

**Related functions:** `SUMPRODUCT`, `ISODD`, `MOD`, `ISEVEN`  
**Functions used:** `ISEVEN`, `ISODD`, `MOD`, `NOT`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7:C17 | array | `=--(ISODD(+B7:B17))` | `1` | 1 |
| D7:D17 | array | `=MOD(B7:B17,2)=1` | `True` | 1 |
| E7:E17 | array | `=NOT(--(ISEVEN(+B7:B17)))` | `True` | 1 |
| G8 | array | `=SUM(--(ISODD(+B7:B17)))` | `5` | 1 |
| H8 | array | `=SUM(--(MOD(B7:B17,2)=1))` | `5` | 1 |
| I8 | array | `=SUM(--(NOT(--(ISEVEN(+B7:B17)))))` | `5` | 1 |

## 16-P: Count cells that contain positive numbers

**Related functions:** `SUMPRODUCT`, `COUNTIF`  
**Functions used:** `COUNTIF`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| F8 | normal | `=COUNTIF(B7:B17,">0")` | `6` | 1 |
| G8 | array | `=SUM(--(B7:B17>0))` | `6` | 1 |
| F9 | normal | `=COUNTIF(B7:B17,">=0")` | `8` | 1 |
| G9 | array | `=SUM(--(B7:B17>=0))` | `8` | 1 |

## 17-P: Count cells that contain specific text

**Related functions:** `SUMPRODUCT`, `COUNTIF`, `SEARCH`, `FIND`, `ISNUMBER`  
**Functions used:** `COUNTIF`, `FIND`, `ISNUMBER`, `LOWER`, `SEARCH`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7:C17 | array | `=FIND(C6,LOWER(B7:B17),1)` | `1` | 1 |
| D7:D17 | array | `=SEARCH(C6,B7:B17)` | `1` | 1 |
| G9 | normal | `=COUNTIF(B7:B17,"*a*")` | `6` | 1 |
| H9 | array | `=SUM(--ISNUMBER(SEARCH("a",B7:B17)))` | `6` | 1 |
| I9 | array | `=SUM(--(ISNUMBER(FIND("a",LOWER(B7:B17)))))` | `6` | 1 |
| G10 | normal | `=COUNTIF(B7:B17,"*2*")` | `4` | 1 |
| H10 | array | `=SUM(--ISNUMBER(SEARCH("2",B7:B17)))` | `4` | 1 |
| I10 | array | `=SUM(--(ISNUMBER(FIND("2",LOWER(B7:B17)))))` | `4` | 1 |
| G11 | normal | `=COUNTIF(B7:B17,"*-s*")` | `3` | 1 |
| H11 | array | `=SUM(--ISNUMBER(SEARCH("-s",B7:B17)))` | `3` | 1 |
| I11 | array | `=SUM(--(ISNUMBER(FIND("-s",LOWER(B7:B17)))))` | `3` | 1 |
| G12 | normal | `=COUNTIF(B7:B17,"*x*")` | `1` | 1 |
| H12 | array | `=SUM(--ISNUMBER(SEARCH("x",B7:B17)))` | `1` | 1 |
| I12 | array | `=SUM(--(ISNUMBER(FIND("x",LOWER(B7:B17)))))` | `1` | 1 |

## 18-P: Count cells that do not contain specific text

**Related functions:** `SUMPRODUCT`, `COUNTIF`, `NOT`, `SEARCH`, `FIND`, `ISNUMBER`  
**Functions used:** `COUNTIF`, `FIND`, `ISNUMBER`, `LOWER`, `NOT`, `SEARCH`, `SUM`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7:C17 | array | `=NOT(ISNUMBER(FIND(C6,LOWER(B7:B17),1)))` | `False` | 1 |
| D7:D17 | array | `=NOT(ISNUMBER(SEARCH(C6,B7:B17)))` | `False` | 1 |
| G9 | normal | `=COUNTIF(B7:B17,"<>*a*")` | `5` | 1 |
| H9 | array | `=SUM(--NOT(ISNUMBER(SEARCH("a",B7:B17))))` | `5` | 1 |
| I9 | array | `=SUM(--NOT(ISNUMBER(FIND("a",LOWER(B7:B17)))))` | `5` | 1 |
| G10 | normal | `=COUNTIF(B7:B17,"<>*0*")` | `1` | 1 |
| H10 | array | `=SUM(--NOT(ISNUMBER(SEARCH("0",B7:B17))))` | `1` | 1 |
| I10 | array | `=SUM(--NOT(ISNUMBER(FIND("0",LOWER(B7:B17)))))` | `1` | 1 |
| G11 | normal | `=COUNTIF(B7:B17,"<>*-r*")` | `7` | 1 |
| H11 | array | `=SUM(--NOT(ISNUMBER(SEARCH("-r",B7:B17))))` | `7` | 1 |
| I11 | array | `=SUM(--NOT(ISNUMBER(FIND("-r",LOWER(B7:B17)))))` | `7` | 1 |

## 19-P: Count cells that contain errors

**Functions used:** `ERROR.TYPE`, `FIND`, `IFERROR`, `IFNA`, `ISERR`, `ISERROR`, `NA`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| E6 | array | `=SUMPRODUCT(--ISERROR(DAT))` | `2` | 1 |
| F6 | array | `=SUM(--(ISERROR(DAT)))` | `2` | 1 |
| G6 | array | `=SUM(--(IFNA(ERROR.TYPE(DAT),0)>0))` | `2` | 1 |
| P6 | array | `=SUMPRODUCT(--ISERR(DAT))` | `1` | 1 |
| Q6 | array | `=SUMPRODUCT(--(IFERROR(ERROR.TYPE(DAT),0)=3))` | `1` | 1 |
| B8 | normal | `=FIND("x","the")` | `#VALUE!` | 1 |
| B12 | normal | `=NA()` | `#N/A` | 1 |

## 20-P: Count cells that contain errors

**Functions used:** `ERROR.TYPE`, `FIND`, `IFNA`, `ISERROR`, `NA`, `NOT`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| E6 | array | `=SUMPRODUCT(--NOT(ISERROR(DAT)))` | `9` | 1 |
| F6 | array | `=SUM(--NOT(ISERROR(DAT)))` | `9` | 1 |
| G6 | array | `=SUM(--NOT(IFNA(ERROR.TYPE(DAT),0)>0))` | `9` | 1 |
| B8 | normal | `=FIND("x","the")` | `#VALUE!` | 1 |
| B12 | normal | `=NA()` | `#N/A` | 1 |

## 21-P: Count cells that do not contain specific text

**Related functions:** `ISNUMBER`, `SEARCH`, `MMULT`, `TRANSPOSE`, `REDUCE`  

_No formulas on this sheet (explanation or layout only)._

## 22-P: Count dates in given year

**Related functions:** `SUMPRODUCT`  
**Functions used:** `ANCHORARRAY`, `COUNTIF`, `COUNTIFS`, `DATE`, `IF`, `SUM`, `SUMPRODUCT`, `YEAR`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C8:C17 | array | `=YEAR(B8:B17)` | `2023` | 1 |
| D8 | normal | `=DATE(F8,1,1)` | `2023-01-01` | 10 |
| G8 | array | `=SUM(--(YEAR($B$8:$B$17)=F8))` | `2` | 3 |
| H8 | normal | `=COUNTIF(ANCHORARRAY($C$8),F8)` | `2` | 3 |
| I8:I10 | array | `=COUNTIF(ANCHORARRAY(C8),F8:F10)` | `2` | 1 |
| J8 | array | `=SUMPRODUCT(--(YEAR($B$8:$B$17)=F8))` | `2` | 3 |
| K8 | array | `=SUM(IF(YEAR($B$8:$B$17)=F8,1,0))` | `2` | 3 |
| L8 | normal | `=COUNTIFS($B$8:$B$17,">="&DATE(F8,1,1),$B$8:$B$17,"<"&DATE(F8,12,31))` | `2` | 3 |

## 23-P: Count if row meets internal criteria

**Related functions:** `SUMPRODUCT`  
**Functions used:** `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| E8:E18 | array | `=C8:C18>D8:D18` | `False` | 1 |
| F8:F18 | array | `=C8:C18<D8:D18` | `True` | 1 |
| I8 | array | `=SUM(--(C8:C18>D8:D18))` | `4` | 1 |
| J8 | array | `=SUMPRODUCT(--(C8:C18>D8:D18))` | `4` | 1 |
| I9 | array | `=SUM(--(C8:C18<D8:D18))` | `6` | 1 |
| J9 | array | `=SUMPRODUCT(--(C8:C18<D8:D18))` | `6` | 1 |

## 24-P: Count if row meets multiple criteria

**Functions used:** `MONTH`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| H5 | array | `=SUM((E5:E15="TX")*(C5:C15>100)*(MONTH(F5:F15)=3))` | `2` | 1 |
| H6 | array | `=SUMPRODUCT((E5:E15="TX")*(C5:C15>100)*(MONTH(F5:F15)=3))` | `2` | 1 |

## 25-P: Count if two criteria match

**Related functions:** `COUNTIFS`, `SUMPRODUCT`  
**Functions used:** `COUNTIFS`, `SUM`, `SUMPRODUCT`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| H7 | normal | `=COUNTIFS(B7:B17,"Blue",C7:C17,">15")` | `4` | 1 |
| I7 | array | `=SUM(--(B7:B17="Blue")*(C7:C17>15))` | `4` | 1 |
| J7 | array | `=SUMPRODUCT(--($B$7:$B$17="Blue"),--($C$7:$C$17>15))` | `4` | 1 |
| H8 | normal | `=COUNTIFS(B7:B17,"Red",C7:C17,">15")` | `1` | 1 |
| I8 | array | `=SUM(--(B7:B17="Red")*(C7:C17>15))` | `1` | 1 |
| J8 | array | `=SUMPRODUCT(--($B$7:$B$17="Red"),--($C$7:$C$17>15))` | `1` | 1 |

## 26-P: Count items in list

**Related functions:** `COUNTIF`  
**Functions used:** `COUNTIF`  

| Cell | Type | Formula | Result | Cells |
| --- | --- | --- | --- | --- |
| C7 | normal | `=COUNTIF($B$7:$B$17,B7)` | `3` | 11 |
