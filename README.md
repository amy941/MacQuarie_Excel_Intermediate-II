# WEEK 1: 
# 🔗Link: [Week 1_folder](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/tree/main/Week%201)
### - Data Validation
- ```Data``` tab--> Data Validation (Settings/Input Message/ Error Alert)
   - *Settings*: define the validation criteria. The default allows ```any value```, meaning **no validation** is taking place. Drop-down shows: Any value, Whole number, Decimal, List, Date, Time, ...
     ```Ignore blank```: Excel won't consider a blank cell to be invalid.
     
   - *Input Message*:
 
   - *Error Alert*: Stop 🚫, Warning ⚠️, Information ℹ️

- If data is **copy-pasted**, or **imported**, it actually **doesn't enforce** data validation rules. **Only works for data that's been entered manually.**
- **Text length** refers to any characters, or combination of text and numeric characters.

### - Create Drop-down Lists
- ```Data``` --> ```Data Validation``` --> Settings --> **Allow: List** | Source: better type in alphabetically
- Converting lookup list into a named range and table so we don't need to update the validation criteria as the look-up list changes.
- **Drop-down list**, items should be seperated by **comma** or **comma and Space**

### - Using Formulas in Data Validation
- **Duplicate code:** ```Data``` --> ```Data Validation``` --> Settings --> **Allow: Custom** | **Formula: =countifs(Product_Code,A4) <= 1**
- **Allow** in **Data Validation** use a formula: **Custom**, **List**
    
### - Working w Data Validation
- ```Data Validation``` drop-down: Circle Invalid Data ⭕
- ```Find & Select``` tab --> Go to Special... --> Data Validation: All or Same
- **Copy data Validation** from one sheet to another: **Paste Special**

### - Advanced Conditional Formatting
- ```Conditional Formatting``` --> New Rule...--> "Use a formula to determine which cells to format" --> **Format values where this formula is true:** = H4 < J4 (w/o $ signs) --> Preview: Format (Font:Bold, Fill:Color)
- ```Conditional Formatting``` --> New Rule...--> "Use a formula to determine which cells to format" --> **Format values where this formula is true:** = **$E4** = $O$4 **(⚠️ Row to go Relative while Column remain Abs)** --> Preview: Format (Fill:Color)
  
💥 **- Week 1_Practice Challenge:** [challenge](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%201/C3-W1-Practice-Challenge.xlsx)

💥💥 **- Week 1_Assessment:** [assessment_Week 1](https://github.com/amy941/MacQuarie_Excel_Intermediate-II/blob/main/Week%201/C3-W1-Assessment.xlsx)

---

# WEEK 2
# 🔗Link: [Week 2_folder]()
### - Logical Functions I: IF

**=IF(logical test, [value_if_true], [value_if_false])**
- First argument is **a logical test**, compares 2 values using a **logical operator**
  ![logical operator](https://github.com/user-attachments/assets/4e3ef65a-d3e2-4e5f-abfe-975a3472416a)
  
- Second argument in brackets is the **"value_if_true"**, could be a value we just type into the cell /or a calculated value.
  * if the logical test equates to True, then whatever we've got between two commas will occur.
  * if the logical test equates to False, then it's going to do the third and last argument **"value_if_false**

- If working w text, put double quotes **" "** /or quotation marks **' '** /or **""** (leave Blank) 
- When comparing text, the equals is **not case sensitive**
- =IF(F4="Y",D4*5%,0)

### - Logical Functions II: AND, OR

**=AND(logical1, [logical2], ...)
  =OR(logical1, [logical2], ...)
  Up to 255 logical tests❗
  Only returns TRUE/FALSE**

- **=AND(logical1, [logical2], ...)**
  * =AND(B4>0,C4<>"Y")
  * evaluate multiple logical tests
  * If x & y & z & ... are **ALL** True, then it returns True

- **=OR(logical1, [logical2], ...)**
  * =OR(I4>=16, J4)
  * If **any** of these are True: x,y,z,..., then returns True


### - Combining Logical Functions I: IF, AND, OR

**=IF(AND(logical1, logical2, ...), [value_if_true], [value_if_false])
  =IF(OR(logical1, logical2, ...), [value_if_true], [value_if_false])**

- =IF(AND(B4>0,C4<>"Y"),B4*10%,0)
- =IF(OR(I4>=16,J4),250,0)


### - Combining Logical Functions II: Nested IFs
![nested IFs](https://github.com/user-attachments/assets/80651868-47ff-4e95-83cb-385a064d9bbb)

**=IF(Balance= 0, "A", IF(Balance > 0, "B", "C"))**

### - Handling Errors: IFERROR, IFNA
- =IFERROR(AVERAGE('Invoice Data'!$O$4:$O$654),"")
- =IFNA(VLOOKUP('Invoice Data'!$A4,BPay!$B$4:$D$10,3,0),0)

💥 **- Week 2_Practice Challenge:** [challenge]()

💥💥 **- Week 2_Assessment:** [assessment_Week 2]() 

---

# WEEK 3
# 🔗Link: [Week 3_folder]()
### - Introduction to lookups: CHOOSE
### - Approximate Matches: Range VLOOKUP
### - Exact Matches: Exact Match VLOOKUP
### - Finding a Position: MATCH
### - Dynamic Lookups: INDEX, MATCH
  
💥 **- Week 3_Practice Challenge:** 

💥💥 **- Week 3_Assessment:** [assessment_Week 3]()

---

# WEEK 4
# 🔗Link: [Week 4_folder]()
### - Error Checking
### - Formula Calc Options
### - Tracing Precedents & Dependents
### - Evaluate Formula, Watch Window
### - Protecting Workbooks & Worksheets


  
💥 **- Week 4_Practice Challenge:** [challenge]()

💥💥 **- Week 4_Assessment:** [assessment_Week 4]()

---

# WEEK 5
# 🔗Link: [Week 5_folder]()
### - Modelling Functions: SUMPRODUCT
### - Data Tables
### - Goal Seek
### - Scenario Manager
### - Solver

💥 **- Week 5_Practice Challenge:** [challenge]()

💥💥 **- Week 5_Assessment:** [assessment_Week 5]()

---

# WEEK 6
# 🔗Link: [Week 6_folder]()
### - Record a Macro
### - Run a Macro
### - Edit a Marco
### - Work w Marcos
### - Relative Reference Macros
  
💥 **- Week 6_Practice Challenge:** [challenge]()

💥💥 **- Week 6_Assessment:** [assessment_Week 6]()

---

# CERTIFICATE
