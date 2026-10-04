===============================================================
PR. 1 FUNDAMENTAL BOOSTER
Excel / Google Sheets Practice Project
===============================================================

HOW TO USE THIS README
----------------------
1. Create a new Google Sheet or Excel file.
2. Create 4 sheets with these exact names:
   - Project Instructions
   - Students Grade
   - Sales Data
   - Employee Data
3. Follow the steps below one by one.
4. Yellow cells = your inputs (you can change them).
5. All other formulas should be typed exactly as shown.


===============================================================
SHEET 1: Project Instructions
===============================================================
Just copy the topics and main tasks from the original project.
No formulas needed on this sheet.


===============================================================
SHEET 2: Students Grade
===============================================================

STEP 1: Enter Headers in Row 3
--------------------------------
A3 = Student ID
B3 = Full Name
C3 = Math
D3 = Science
E3 = English
F3 = Average
G3 = Grade
H3 = Status (Math & Science >80)
I3 = First Name
J3 = Name (UPPER)
K3 = Name (LOWER)


STEP 2: Enter Sample Data (Rows 4 to 15)
----------------------------------------
S001 | Rahul Sharma      | 85 | 78 | 92
S002 | Priya Patel       | 92 | 88 | 85
S003 | Amit Kumar        | 45 | 52 | 60
S004 | Sneha Reddy       | 78 | 81 | 75
S005 | Vikram Singh      | 95 | 90 | 88
S006 | Ananya Gupta      | 62 | 55 | 70
S007 | Rohan Mehta       | 88 | 85 | 90
S008 | Kavita Joshi      | 55 | 48 | 52
S009 | Arjun Nair        | 73 | 79 | 68
S010 | Meera Iyer        | 91 | 94 | 89
S011 | Siddharth Rao     | 40 | 35 | 45
S012 | Divya Krishnan    | 82 | 87 | 80


STEP 3: Write Formulas (start in Row 4 and drag down)
-----------------------------------------------------

Average (F4):
=AVERAGE(C4:E4)

Grade - Nested IF (G4):
=IF(F4>=90,"A",IF(F4>=80,"B",IF(F4>=70,"C",IF(F4>=60,"D","F"))))

Status - IF + AND (H4):
=IF(AND(C4>80,D4>80),"Excellent","Needs Improvement")

First Name - TEXT (I4):
=LEFT(B4,FIND(" ",B4)-1)

UPPER (J4):
=UPPER(B4)

LOWER (K4):
=LOWER(B4)


STEP 4: Summary Formulas (write below the table)
------------------------------------------------
Students with Math > 50:
=COUNTIF(C4:C15,">50")

Students with Average > 60:
=COUNTIF(F4:F15,">60")

Average Math score (only Math > 60):
=AVERAGEIF(C4:C15,">60")

Students with Math > 80 AND Science > 80:
=COUNTIFS(C4:C15,">80",D4:D15,">80")


STEP 5: FILTER Function (Excel 365 / Google Sheets)
---------------------------------------------------
=FILTER(A4:E15,F4:F15>80,"No students")


STEP 6: VLOOKUP Demo
--------------------
Type any Student ID in a cell (example B40 = S005)

Student Name:
=VLOOKUP(B40,A4:B15,2,FALSE)

Math Score:
=VLOOKUP(B40,A4:C15,3,FALSE)


===============================================================
SHEET 3: Sales Data
===============================================================

STEP 1: Product Master Table (Rows 4-12)
----------------------------------------
A4 = Product Code | B4 = Product Name | C4 = Unit Price | D4 = Category

P101 | Laptop Pro 15                | 85000 | Electronics
P102 | Wireless Mouse               | 1200  | Accessories
P103 | USB-C Hub                    | 2500  | Accessories
P104 | Monitor 27"                  | 18500 | Electronics
P105 | Keyboard Mechanical          | 4500  | Accessories
P106 | Webcam HD                    | 3200  | Electronics
P107 | Laptop Stand                 | 1800  | Accessories
P108 | Noise Cancelling Headphones  | 9500  | Electronics


STEP 2: Sales Transactions Headers (Row 16)
-------------------------------------------
A16 = Sale ID
B16 = Date
C16 = Product Code
D16 = Product Name (VLOOKUP)
E16 = Unit Price (VLOOKUP)
F16 = Qty
G16 = Region
H16 = Salesperson
I16 = Total Sales
J16 = Discount Eligible?


STEP 3: Enter some sales data (example starting Row 17)
-------------------------------------------------------
SL001 | 15-Jan-2025 | P101 | (formula) | (formula) | 2 | North | Ravi
SL002 | 18-Jan-2025 | P102 | (formula) | (formula) | 5 | South | Priya
SL003 | 03-Feb-2025 | P104 | (formula) | (formula) | 1 | West  | Amit
... (add 8-12 rows yourself)


STEP 4: Formulas for Sales (Row 17 example)
-------------------------------------------
Product Name (D17):
=VLOOKUP(C17,$A$5:$D$12,2,FALSE)

Unit Price (E17):
=VLOOKUP(C17,$A$5:$D$12,3,FALSE)

Total Sales (I17):
=E17*F17

Discount Eligible? (J17):
=IF(OR(E17>10000,F17>=5),"Yes","No")


STEP 5: Nested IF for Discount %
--------------------------------
=IF(price>50000,0.15,IF(price>10000,0.1,IF(price>5000,0.05,0)))


STEP 6: Aggregation Formulas
----------------------------
Total Sales - North Region:
=SUMIF(G17:G28,"North",I17:I28)

Total Sales - Product P101:
=SUMIF(C17:C28,"P101",I17:I28)

North + P101 (SUMIFS):
=SUMIFS(I17:I28,G17:G28,"North",C17:C28,"P101")

Count of Sales - South:
=COUNTIF(G17:G28,"South")


STEP 7: Other Lookup Functions
------------------------------
Position of a product (MATCH):
=MATCH("P105",A5:A12,0)

INDEX + MATCH (first sale of a salesperson):
=INDEX(A17:A28,MATCH("Ravi",H17:H28,0))

INDIRECT example:
=INDIRECT("'Sales Data'!C5")

OFFSET dynamic sum:
=SUM(OFFSET(I17,0,0,5,1))


===============================================================
SHEET 4: Employee Data
===============================================================

STEP 1: Headers (Row 3)
-----------------------
A3 = Emp ID
B3 = Full Name
C3 = Department
D3 = Date of Joining
E3 = Date of Birth
F3 = Salary
G3 = Years of Service
H3 = Age
I3 = First Name
J3 = Status


STEP 2: Enter Sample Employee Data
----------------------------------
E001 | Rajesh Khanna  | Sales     | 15-Mar-2018 | 20-Jun-1985 | 65000
E002 | Sunita Agarwal | HR        | 01-Jul-2019 | 05-Nov-1990 | 55000
E003 | Mohit Verma    | IT        | 10-Jan-2017 | 14-Feb-1988 | 85000
E004 | Pooja Desai    | Finance   | 22-Sep-2020 | 30-Aug-1992 | 70000
E005 | Karan Malhotra | Sales     | 05-May-2016 | 12-Dec-1983 | 72000
E006 | Neha Kapoor    | Marketing | 18-Feb-2021 | 25-Apr-1995 | 48000
E007 | Aakash Jain    | IT        | 30-Nov-2019 | 08-Sep-1991 | 78000
E008 | Ritu Sharma    | HR        | 12-Apr-2022 | 17-Jan-1994 | 52000
E009 | Vishal Reddy   | Finance   | 08-Aug-2015 | 03-Jul-1980 | 95000
E010 | Anjali Nair    | Marketing | 25-Jun-2020 | 19-Oct-1993 | 58000


STEP 3: Formulas (Row 4 and drag down)
--------------------------------------
Years of Service (G4):
=DATEDIF(D4,TODAY(),"Y")

Age (H4):
=DATEDIF(E4,TODAY(),"Y")

First Name (I4):
=LEFT(B4,FIND(" ",B4)-1)

Status (J4):
=IF(G4>=5,"Senior",IF(G4>=2,"Mid-Level","Junior"))


STEP 4: XLOOKUP / INDEX+MATCH Demo
----------------------------------
Type any Emp ID in a cell (example B17 = E003)

Name:
=XLOOKUP(B17,A4:A13,B4:B13,"Not Found")
OR
=INDEX(B4:B13,MATCH(B17,A4:A13,0))

Salary:
=XLOOKUP(B17,A4:A13,F4:F13,"Not Found")
OR
=INDEX(F4:F13,MATCH(B17,A4:A13,0))


STEP 5: Math Functions
----------------------
ROUND to nearest 1000:
=ROUND(F4,-3)

CEILING to next 5000:
=CEILING(F4,5000)

FLOOR to previous 5000:
=FLOOR(F4,5000)


STEP 6: Absolute vs Relative Reference
--------------------------------------
1. Put Tax Rate 10% in cell B46
2. Make it absolute: $B$46
3. Tax formula:
   =B48*$B$46
   (When you drag this formula down, $B$46 stays fixed)


===============================================================
RECOMMENDED PRACTICE ORDER
===============================================================
1. Students Grade   → IF, Nested IF, AND, COUNTIF, TEXT, VLOOKUP, FILTER
2. Sales Data       → VLOOKUP, Nested IF, SUMIFS, MATCH, OFFSET, INDIRECT
3. Employee Data    → XLOOKUP / INDEX-MATCH, DATEDIF, ROUND, Absolute refs


===============================================================
COLOR TIPS (Optional but useful)
===============================================================
- Blue text   = Input values (you can change)
- Black text  = Formulas (do not overwrite)
- Yellow cell = Important input cell


===============================================================
NOTES
===============================================================
- FILTER and XLOOKUP work in Google Sheets and Excel 365.
- In older Excel use INDEX + MATCH instead of XLOOKUP.
- DATEDIF works in both Excel and Google Sheets.
- Always use absolute references ($A$5:$D$12) when looking up
  from a fixed table so the range does not shift when copied.


Good luck!
Complete each formula yourself — that is the real practice.
===============================================================
