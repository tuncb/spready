Below is a compact, implementable **basic Excel formula language spec**. It is not the full Excel language; it deliberately omits dynamic arrays, structured table references, external workbook links, cube functions, lambda helpers, array constants, and locale-specific syntax. I’m assuming **English-style A1 notation** and comma-separated function arguments.

Microsoft’s own docs group Excel operators into arithmetic, comparison, text-concatenation, and reference operators, and state that formulas begin with `=` and follow a defined operator precedence; the rules below mirror that core behavior. ([Microsoft Support][1])

## 1. Formula shape

```ebnf
formula ::= "=" expression
```

Examples:

```excel
=1+2
=A1+B1
=SUM(A1:A10)
=IF(A1>0,"positive","non-positive")
```

A formula evaluates to one of:

```text
number
text
boolean
error
reference / range       // usually consumed by functions
```

For a basic evaluator, treat dates/times as numbers with formatting handled outside the formula language.

---

## 2. Literals

### Numbers

```ebnf
number ::= digits ["." digits] [("E"|"e") ["+"|"-"] digits]
```

Examples:

```excel
=123
=12.5
=1.2E-3
```

Excel has a finite calculation model; Microsoft lists 15 digits of number precision and a formula-content length limit of 8,192 characters in current Excel specs. ([Microsoft Support][2])

### Text

```ebnf
text ::= '"' { character | '""' } '"'
```

Examples:

```excel
="hello"
="quoted ""text"""
```

Inside text, a literal double quote is represented by two double quotes: `""`.

### Booleans

```excel
TRUE
FALSE
```

Usually these appear as constants or function results:

```excel
=A1>10
=IF(TRUE,1,0)
```

### Error values

A basic subset should support these common Excel error values:

```text
#DIV/0!
#N/A
#NAME?
#NULL!
#NUM!
#REF!
#VALUE!
```

Microsoft lists these as Excel error values. ([Microsoft Support][3])

---

## 3. Cell references and ranges

### A1 cell reference

```ebnf
cell_ref ::= [sheet_ref "!"] ["$"] column ["$"] row
column   ::= "A".."XFD"
row      ::= 1..1048576
```

Examples:

```excel
A1
$A$1
A$1
$A1
Sheet1!A1
'My Sheet'!A1
```

Current Excel worksheets have 1,048,576 rows and 16,384 columns, so the last column is `XFD`. ([Microsoft Support][2])

### Relative, absolute, and mixed references

```text
A1      relative column, relative row
$A$1    absolute column, absolute row
A$1     relative column, absolute row
$A1     absolute column, relative row
```

Relative references adjust when copied; `$` locks the column, row, or both. ([Microsoft Support][4])

### Ranges

```ebnf
range_ref ::= cell_ref ":" cell_ref
```

Examples:

```excel
A1:A10
A1:C3
Sheet1!A1:Sheet1!C10
```

---

## 4. Names / variables

Excel does not have variables in the same way as most programming languages. The closest concepts are:

```text
Defined names      workbook- or worksheet-scoped aliases
LET names          formula-local names
```

### Defined names

A defined name can refer to a cell, range, constant, or formula. Microsoft’s name rules include: first character must be a letter, underscore `_`, or backslash `\`; remaining characters may include letters, numbers, periods, and underscores; spaces are not allowed; names cannot be cell references such as `Z$100` or `R1C1`; names may be up to 255 characters; and names are not case-sensitive. ([Microsoft Support][5])

Basic name grammar:

```ebnf
name ::= name_start { name_char }

name_start ::= letter | "_" | "\"
name_char  ::= letter | digit | "_" | "."
```

Invalid examples:

```text
1Rate       // starts with digit
Sales Tax   // contains space
A1          // conflicts with cell reference
R1C1        // conflicts with R1C1 reference syntax
```

Valid examples:

```text
Sales_Tax
First.Quarter
_rate
\internalName
```

### LET local variables

Basic syntax:

```excel
=LET(name1, value1, calculation)
=LET(x, A1+B1, x*2)
```

Microsoft describes `LET` as assigning names to calculation results inside a formula; the first name must start with a letter and must not conflict with range syntax. ([Microsoft Support][6])

For a basic subset, use the same name rules as defined names, but it is wise to require LET variables to start with a letter or underscore and to reject anything that looks like a cell reference.

---

## 5. Operators

### Arithmetic operators

```text
+     addition
-     subtraction / negation
*     multiplication
/     division
%     percent
^     exponentiation
```

Examples:

```excel
=1+2
=10-3
=2*5
=10/2
=50%
=2^3
=-A1
```

### Comparison operators

```text
=     equal
<>    not equal
<     less than
<=    less than or equal
>     greater than
>=    greater than or equal
```

Comparison operators return `TRUE` or `FALSE`. ([Microsoft Support][1])

Examples:

```excel
=A1=B1
=A1<>B1
=A1>=10
```

### Text concatenation

```text
&     concatenate text
```

Examples:

```excel
="Hello" & " " & "world"
=A1 & "-" & B1
```

Microsoft identifies `&` as Excel’s text concatenation operator. ([Microsoft Support][1])

### Reference operators

```text
:       range
,       union / argument separator depending on context
space   intersection
```

Examples:

```excel
=A1:A10
=SUM(A1:A10,C1:C10)
=SUM(B7:D7 C6:C8)
```

Excel’s reference operators are colon for range, comma for union, and a single space for intersection. ([Microsoft Support][1])

For a basic language, you can support `:` and function-argument comma first, then add reference union/intersection later.

---

## 6. Operator precedence

Highest to lowest:

| Precedence | Operators                       | Meaning                  |
| ---------: | ------------------------------- | ------------------------ |
|          1 | `:`, space, `,`                 | Reference operators      |
|          2 | unary `-`                       | Negation                 |
|          3 | `%`                             | Percent                  |
|          4 | `^`                             | Exponentiation           |
|          5 | `*`, `/`                        | Multiplication, division |
|          6 | `+`, `-`                        | Addition, subtraction    |
|          7 | `&`                             | Text concatenation       |
|          8 | `=`, `<>`, `<`, `<=`, `>`, `>=` | Comparison               |

Parentheses override precedence, and Excel evaluates operators of the same precedence from left to right. ([Microsoft Support][1])

Examples:

```excel
=5+2*3        // 11
=(5+2)*3      // 21
=2^3*4        // 32
="A"&1+2      // "A3" in a permissive coercion model
```

---

## 7. Core grammar

This is a practical parser grammar for a basic Excel-like formula language.

```ebnf
formula        ::= "=" expression

expression     ::= comparison

comparison     ::= concatenation [ comp_op concatenation ]
comp_op        ::= "=" | "<>" | "<" | "<=" | ">" | ">="

concatenation  ::= additive { "&" additive }

additive       ::= multiplicative { ("+" | "-") multiplicative }

multiplicative ::= power { ("*" | "/") power }

power          ::= percent { "^" percent }

percent        ::= unary { "%" }

unary          ::= ("+" | "-") unary
                 | primary

primary        ::= number
                 | text
                 | boolean
                 | error
                 | reference
                 | name
                 | function_call
                 | "(" expression ")"

function_call  ::= function_name "(" [ argument_list ] ")"

argument_list  ::= expression { "," expression }

reference      ::= cell_ref [ ":" cell_ref ]

cell_ref       ::= [ sheet_ref "!" ] ["$"] column ["$"] row

sheet_ref      ::= simple_sheet_name | quoted_sheet_name

simple_sheet_name ::= name

quoted_sheet_name ::= "'" { character | "''" } "'"
```

Notes:

```text
1. Whitespace is generally insignificant, except when used as the reference intersection operator.
2. Function names and defined names should be treated case-insensitively.
3. Comma is ambiguous: inside function calls it separates arguments; between references it can be a union operator.
4. This grammar treats chained comparisons as invalid or implementation-defined. Excel-style formulas usually use AND/OR instead.
```

---

## 8. Basic function set

Microsoft categorizes Excel worksheet functions by purpose, including math/trig, logical, text, lookup/reference, date/time, statistical, and other groups. The set below is a small practical subset, not the entire Excel function catalog. ([Microsoft Support][7])

### Math and aggregation

| Function  | Signature                        | Meaning                        |
| --------- | -------------------------------- | ------------------------------ |
| `SUM`     | `SUM(value1, [value2], ...)`     | Sum numbers and numeric ranges |
| `PRODUCT` | `PRODUCT(value1, [value2], ...)` | Multiply values                |
| `MIN`     | `MIN(value1, [value2], ...)`     | Minimum                        |
| `MAX`     | `MAX(value1, [value2], ...)`     | Maximum                        |
| `AVERAGE` | `AVERAGE(value1, [value2], ...)` | Arithmetic mean                |
| `COUNT`   | `COUNT(value1, [value2], ...)`   | Count numeric values           |
| `COUNTA`  | `COUNTA(value1, [value2], ...)`  | Count non-empty values         |
| `ABS`     | `ABS(number)`                    | Absolute value                 |
| `ROUND`   | `ROUND(number, num_digits)`      | Round to digits                |
| `INT`     | `INT(number)`                    | Round down to integer          |
| `MOD`     | `MOD(number, divisor)`           | Remainder                      |
| `POWER`   | `POWER(number, power)`           | Exponentiation                 |
| `SQRT`    | `SQRT(number)`                   | Square root                    |

Example:

```excel
=SUM(A1:A10)
=ROUND(AVERAGE(B1:B10),2)
```

### Logical

Microsoft’s logical-function reference includes `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `NOT`, `OR`, `TRUE`, and related functions. ([Microsoft Support][8])

| Function  | Signature                                 | Meaning                        |
| --------- | ----------------------------------------- | ------------------------------ |
| `TRUE`    | `TRUE()`                                  | Boolean true                   |
| `FALSE`   | `FALSE()`                                 | Boolean false                  |
| `AND`     | `AND(logical1, [logical2], ...)`          | True if all arguments are true |
| `OR`      | `OR(logical1, [logical2], ...)`           | True if any argument is true   |
| `NOT`     | `NOT(logical)`                            | Negate boolean                 |
| `IF`      | `IF(test, value_if_true, value_if_false)` | Conditional                    |
| `IFERROR` | `IFERROR(value, value_if_error)`          | Error fallback                 |

Examples:

```excel
=IF(A1>0,"positive","not positive")
=AND(A1>0,B1>0)
=IFERROR(A1/B1,0)
```

### Text

Microsoft’s text-function reference includes common text functions such as `LEFT`, `LEN`, `LOWER`, `MID`, `RIGHT`, `TEXTJOIN`, `TRIM`, `UPPER`, and `VALUE`. ([Microsoft Support][9])

| Function   | Signature                                        | Meaning                |
| ---------- | ------------------------------------------------ | ---------------------- |
| `LEN`      | `LEN(text)`                                      | Text length            |
| `LEFT`     | `LEFT(text, [num_chars])`                        | Leftmost characters    |
| `RIGHT`    | `RIGHT(text, [num_chars])`                       | Rightmost characters   |
| `MID`      | `MID(text, start_num, num_chars)`                | Substring              |
| `TRIM`     | `TRIM(text)`                                     | Remove extra spaces    |
| `LOWER`    | `LOWER(text)`                                    | Lowercase              |
| `UPPER`    | `UPPER(text)`                                    | Uppercase              |
| `CONCAT`   | `CONCAT(value1, [value2], ...)`                  | Concatenate            |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, value1, ...)` | Join with delimiter    |
| `VALUE`    | `VALUE(text)`                                    | Convert text to number |

Examples:

```excel
=LEFT(A1,3)
=UPPER(TRIM(A1))
=TEXTJOIN(", ",TRUE,A1:A5)
```

### Lookup and reference

Microsoft’s lookup/reference documentation includes functions such as `CHOOSE`, `COLUMN`, `COLUMNS`, and newer lookup functions such as `XLOOKUP`; Microsoft describes `XLOOKUP` as an improved version of `VLOOKUP` that can return exact matches by default. ([Microsoft Support][10])

| Function  | Signature                                                           | Meaning                               |
| --------- | ------------------------------------------------------------------- | ------------------------------------- |
| `INDEX`   | `INDEX(array, row_num, [column_num])`                               | Return item at row/column             |
| `MATCH`   | `MATCH(lookup_value, lookup_array, [match_type])`                   | Return position                       |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])` | Lookup and return corresponding value |
| `VLOOKUP` | `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])` | Vertical lookup                       |
| `CHOOSE`  | `CHOOSE(index_num, value1, [value2], ...)`                          | Pick nth value                        |
| `ROW`     | `ROW([reference])`                                                  | Row number                            |
| `COLUMN`  | `COLUMN([reference])`                                               | Column number                         |

Examples:

```excel
=INDEX(B1:B10,3)
=MATCH("abc",A1:A10,0)
=XLOOKUP(E1,A1:A10,B1:B10,"not found")
```

### Date/time, optional basic set

| Function | Signature                | Meaning               |
| -------- | ------------------------ | --------------------- |
| `TODAY`  | `TODAY()`                | Current date          |
| `NOW`    | `NOW()`                  | Current date/time     |
| `DATE`   | `DATE(year, month, day)` | Construct date serial |
| `YEAR`   | `YEAR(date)`             | Year                  |
| `MONTH`  | `MONTH(date)`            | Month                 |
| `DAY`    | `DAY(date)`              | Day                   |

For implementation, store dates as numbers and leave display formatting to the host spreadsheet.

---

## 9. Basic evaluation rules

A practical evaluator can use these rules:

```text
1. Evaluate references to values or ranges.
2. Evaluate arithmetic over numbers.
3. Evaluate text concatenation by converting operands to text.
4. Evaluate comparisons to TRUE/FALSE.
5. Propagate errors unless a function such as IFERROR handles them.
6. Evaluate ranges lazily where possible, because functions such as SUM need the whole range.
7. Reject unknown names/functions with #NAME?.
8. Reject invalid references with #REF!.
9. Return #DIV/0! for division by zero.
10. Return #VALUE! for incompatible value types.
```

Suggested type coercion for a simple subset:

```text
Arithmetic:
  number op number -> number
  boolean may coerce to 1/0
  numeric text may optionally coerce to number
  otherwise #VALUE!

Concatenation:
  any scalar -> text
  text & text -> text

Comparison:
  numbers compare numerically
  text compares lexicographically
  booleans compare as booleans
```

Excel’s real coercion behavior is more nuanced, so this is a good “basic compatibility” model rather than a full clone.

---

## 10. Minimal implementation checklist

Support these first:

```text
Syntax:
  = prefix
  numbers, strings, booleans, errors
  A1 references
  ranges with :
  parentheses
  function calls

Operators:
  + - * / ^ %
  &
  = <> < <= > >=

Names:
  defined names
  LET local names, optional

Functions:
  SUM, MIN, MAX, AVERAGE, COUNT, COUNTA
  IF, AND, OR, NOT, IFERROR
  LEN, LEFT, RIGHT, MID, TRIM, UPPER, LOWER, VALUE
  INDEX, MATCH, XLOOKUP or VLOOKUP
```

A small example using most of the subset:

```excel
=IFERROR(
  IF(AVERAGE(B2:B10)>=70,
     "Pass: " & ROUND(AVERAGE(B2:B10),1),
     "Fail"),
  "No valid score"
)
```

Equivalent compact grammar target:

```ebnf
formula        ::= "=" expression
expression     ::= comparison
comparison     ::= concatenation [ comp_op concatenation ]
concatenation  ::= additive { "&" additive }
additive       ::= multiplicative { ("+" | "-") multiplicative }
multiplicative ::= power { ("*" | "/") power }
power          ::= percent { "^" percent }
percent        ::= unary { "%" }
unary          ::= ("+" | "-") unary | primary
primary        ::= literal | reference | name | function_call | "(" expression ")"
function_call  ::= name "(" [ expression { "," expression } ] ")"
reference      ::= cell_ref [ ":" cell_ref ]
cell_ref       ::= [ sheet_ref "!" ] ["$"] column ["$"] row
```

[1]: https://support.microsoft.com/en-us/office/calculation-operators-and-precedence-36de9366-46fe-43a3-bfa8-cf6d8068eacc "Calculation operators and precedence - Microsoft Support"
[2]: https://support.microsoft.com/en-gb/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?utm_source=chatgpt.com "Excel specifications and limits"
[3]: https://support.microsoft.com/en-us/office/hide-error-values-and-error-indicators-in-cells-d171b96e-8fb4-4863-a1ba-b64557474439?utm_source=chatgpt.com "Hide error values and error indicators in cells"
[4]: https://support.microsoft.com/en-gb/office/switch-between-relative-and-absolute-references-981f5871-7864-42cc-b3f0-41ffa10cc6fc "Switch between relative and absolute references - Microsoft Support"
[5]: https://support.microsoft.com/en-au/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1 "Names in formulas - Microsoft Support"
[6]: https://support.microsoft.com/en-au/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999?utm_source=chatgpt.com "LET function"
[7]: https://support.microsoft.com/en-au/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb "Excel functions (by category) - Microsoft Support"
[8]: https://support.microsoft.com/en-us/office/logical-functions-reference-e093c192-278b-43f6-8c3a-b6ce299931f5 "Logical functions (reference) - Microsoft Support"
[9]: https://support.microsoft.com/en-us/office/text-functions-reference-cccd86ad-547d-4ea9-a065-7bb697c2a56e?utm_source=chatgpt.com "Text functions (reference)"
[10]: https://support.microsoft.com/en-us/office/lookup-and-reference-functions-reference-8aa21a3a-b56a-4055-8257-3ec89df2b23e "Lookup and reference functions (reference) - Microsoft Support"
