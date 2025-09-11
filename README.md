# Payroll-Automation

⚠️⚠️ Please note that this is synthetic data and does not contain any real information about my clients or their company.

First, I duplicated the dataset. As data professionals, we often do this to preserve the original copy for reference in case anything goes wrong. Issues rarely occur, but keeping an untouched version is considered best practice.


<img width="1366" height="547" alt="Image" src="https://github.com/user-attachments/assets/eed4382d-0e53-4b7a-852a-b0028fd68223" />

This is the data we got. It’s ugly and messy, so we’re going to fix it.

Next, I converted the dataset into a table and applied a purple theme to match the company’s colors. I reformatted columns like DateOfBirth, HireDate, TerminationDate, and LastPromotion into proper date types and added placeholders where needed.

Still within General Data, I changed columns such as JobLevel, EmploymentType, PayFrequency, BenefitPlan, and IsActive into dropdown lists for easier data entry and consistency.

<img width="1366" height="484" alt="Image" src="https://github.com/user-attachments/assets/d2dc8a91-a63f-484b-9df9-1f10cc528da7" />

As you can see, our data is already starting to look much better.
The dataset is already quite clean except for a few blank cells, so we can skip extensive cleaning and move straight to calculations.

# Employee Master Data
As requested by the client, I created three new sheets. The first is Employee Master Data, which contains mostly static personal and employment details that rarely change. The columns include:
EmployeeID, FirstName, LastName, Email, PhoneNumber, DateOfBirth, Gender, MaritalStatus, SSN, JobTitle, JobLevel, Department, EmploymentType, HireDate, TerminationDate, CompanyHireDate, LastPromotionDate, StreetAddress, City, State, ZipCode, IBAN, RoutingNumber, and BankAccountNumber.

For the EmployeeID column, I used the formula 
```excel
='General Data'!A2
```
to pull the ID directly from the **General Data** sheet, ensuring consistency across both sheets.

<img width="1361" height="423" alt="Image" src="https://github.com/user-attachments/assets/bac48513-6a6e-482d-9ede-3cfaf96234db" />

For the FirstName column, I used the formula

```excel
=INDEX('General Data'!B:B, MATCH(A2, 'General Data'!A:A, 0))
```

This finds the row where the EmployeeID in A2 appears in column A of the General Data sheet and returns the FirstName from column B of that row.

I used this INDEX MATCH method for multiple columns including EmployeeID, FirstName, LastName, Email, PhoneNumber, DateOfBirth, Gender, MaritalStatus, SSN, JobTitle, JobLevel, Department, EmploymentType, HireDate, StreetAddress, City, State, ZipCode, IBAN, RoutingNumber, and BankAccountNumber.

MATCH finds the correct row, and INDEX returns the value from the specified column in that row.

You might be thinking, why not use VLOOKUP? VLOOKUP can only look to the right of the lookup column. Even though EmployeeID is in the first column, inserting a new column like MiddleName would break the formulas. INDEX MATCH avoids this because it specifies which column to return, making it more flexible and reliable.

For the TerminationDate and LastPromotionDate columns, I did things a bit differently. Using the regular formula caused an error because some cells were blank, which makes sense since active employees don’t have a termination date. To fix this, I used

```excel
=IFERROR(INDEX('General Data'!O:O, MATCH(A2, 'General Data'!A:A, 0)),"")
```

This way, any empty values in the column return "" instead of an error.

<img width="1366" height="413" alt="Image" src="https://github.com/user-attachments/assets/0a353c8c-d23e-452e-a77e-d3d990d4e84f" />

As you can see, our HR data is looking nice.

<img width="1318" height="428" alt="Image" src="https://github.com/user-attachments/assets/4e5c3fbb-8d59-409c-97a9-11887a370dde" />

All of this data is referenced from the General Data sheet, so General Data remains the ultimate source. Any changes made there will automatically be reflected here, as requested by the client.


# Payroll Data
Next, I created a PayrollData sheet that only includes employees who meet the condition of being active, meaning their IsActive column is set to ‘Yes’.

## JavaScript in Apps Script

To pull data from General Data into a separate Payroll sheet, I wrote a script using JavaScript in Apps Script. The goal was to only bring in active employees, keeping the payroll data clean and up to date.

The script starts by accessing the spreadsheet and defining the source and target sheets:

```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();
var source = ss.getSheetByName('General Data');
var target = ss.getSheetByName('Sheet5');
```

It then grabs all the data from the source sheet and identifies the IsActive column. This column tells whether an employee is active. If it’s missing, the script stops because filtering would be impossible:

```javascript
var data = source.getDataRange().getValues();
var headers = data[0];
var isActiveCol = headers.indexOf("IsActive"); 
if (isActiveCol === -1) throw new Error("IsActive column not found.");
```

Next, it filters only the rows where IsActive = "Yes", ensuring that only active employees are pulled:

```javascript
var active = data.slice(1).filter(r => r[isActiveCol] === "Yes");
```

Before writing the new data, the script clears any old rows in the target sheet while keeping the headers intact, so old and new data don’t stack up:

```javascript
if (target.getLastRow() > 1) {
  target.getRange(2, 1, target.getLastRow() - 1, headers.length).clearContent();
}
```

Then, it writes the filtered active employees into the target sheet starting from row 2:

```javascript
if (active.length) {
  target.getRange(2, 1, active.length, headers.length).setValues(active);
}
```

Finally, the script logs how many active employees were added, so I can confirm that the update ran successfully:

```javascript
Logger.log("PayrollData updated with " + active.length + " active employees.");
```

This approach ensures the Payroll sheet always reflects the active employees from General Data automatically, without any manual filtering or copying.


<img width="1198" height="643" alt="Image" src="https://github.com/user-attachments/assets/6dc1358f-4011-49d6-8a2b-cfaaff48332d" />


After that, we hit **Run**, and the script started executing. As you can see in the execution log, it ran successfully without errors.

<img width="1284" height="653" alt="Image" src="https://github.com/user-attachments/assets/4c51955c-6fbe-4535-8995-4eab196acd2f" />

Please note that I also had to grant the required Google permissions for the script to run. Without that, the execution would not have been successful.

As you can see below, the PayrollData has now been successfully pulled out from the General Data sheet, and I formatted it neatly into a table so it is structured, easy to read, and ready for analysis or reporting.

<img width="1254" height="420" alt="Image" src="https://github.com/user-attachments/assets/6a43ec41-3716-4e0d-aa16-9528cec4dec6" />



I created a column called GrossPay. Since the employees in this dataset are all salaried, it makes sense that their gross pay was calculated based on their pay frequency.

For employees paid monthly, GrossPay is the Annual Salary ÷ 12.
For employees paid weekly, GrossPay is the Annual Salary ÷ 52.
For employees paid bi-weekly, GrossPay is the Annual Salary ÷ 26.


```excel
=IF(Q3="Monthly", P3/12, 
   IF(Q3="Weekly", P3/52, 
   IF(Q3="Bi-Weekly", P3/26, "")))
```

This way, GrossPay dynamically adjusts depending on each employee’s pay frequency.

<img width="1366" height="418" alt="Image" src="https://github.com/user-attachments/assets/477bde10-f32e-4ddc-b2f5-c62f75f9ae6a" />

For the Current Bonus, I created a column that calculates 10% of the employee’s pay for that specific pay period. 

So, for employees paid monthly, the bonus is 10% of their monthly salary (Annual Salary ÷ 12 × 0.1).
For those paid bi-weekly, it’s 10% of their bi-weekly salary (Annual Salary ÷ 26 × 0.1).
And for employees paid weekly, it’s 10% of their weekly salary (Annual Salary ÷ 52 × 0.1).

Here’s the formula I used:

```excel
=IF(Q2="Monthly", (P2/12)*0.1,
 IF(Q2="Bi-Weekly", (P2/26)*0.1,
 IF(Q2="Weekly", (P2/52)*0.1, 0)))
```

With this, the bonus is automatically tailored to each employee’s pay frequency, without needing to calculate it manually.

<img width="528" height="273" alt="Image" src="https://github.com/user-attachments/assets/aa64e48b-1919-4ec4-94e9-633040f839b1" />


## Tax

My Formula:

```excel
=MIN(AJ2+AK2,1000)*0.05
 +MAX(MIN(AJ2+AK2-1000,2000),0)*0.1
 +MAX(MIN(AJ2+AK2-3000,2000),0)*0.15
 +MAX(AJ2+AK2-5000,0)*0.2
```

Breakdown:

* `MIN(AJ2+AK2,1000)*0.05`  takes the first 1000 and taxes it at 5%
* `MAX(MIN(AJ2+AK2-1000,2000),0)*0.1`  takes the next 2000 (1001–3000) and taxes it at 10%
* `MAX(MIN(AJ2+AK2-3000,2000),0)*0.15`  takes the next 2000 (3001–5000) and taxes it at 15%
* `MAX(AJ2+AK2-5000,0)*0.2`  takes everything above 5000 and taxes it at 20%



<img width="867" height="430" alt="Image" src="https://github.com/user-attachments/assets/962fbe2f-1944-48f3-a675-63b36f479299" />

## Current Retirement Contribution


For the retirement contribution, I implemented the formula:

```excel
=(AJ2+AK2)*0.05
```

Breakdown:

* `AJ2` → GrossPay
* `AK2` → Current Bonus
* `(AJ2+AK2)` → Total Earnings

The formula multiplies total earnings by `0.05`, which means 5% of employee’s salary plus bonus is automatically deducted as their retirement contribution.

<img width="1119" height="389" alt="Image" src="https://github.com/user-attachments/assets/3e828bb4-91f6-43d8-890d-4e7351561944" />


## NetPay

For the NetPay column, I used the formula:

```excel
=SUM(AJ2:AK2) - SUM(AL2:AM2)
```

Breakdown:

* `AJ2` → GrossPay
* `AK2` → Current Bonus
* `AL2` → Tax
* `AM2` → Current Retirement Contribution

The formula first adds GrossPay + Current Bonus to get Total Earnings, then subtracts Tax + Retirement Contribution (Total Deductions).

<img width="1123" height="398" alt="Image" src="https://github.com/user-attachments/assets/6a9466c7-8e31-40c4-9015-01bf5b5fb23f" />



## Date Processed

For tracking when payroll is processed, I used the formula:

```excel
=TODAY()
```

The `TODAY()` function  returns the current date based on the system clock. Each time the payroll sheet is opened or recalculated, the **Date Processed** column reflects the correct processing date without manual input.

<img width="752" height="366" alt="Image" src="https://github.com/user-attachments/assets/1959d58b-0736-4aef-81ac-4265c22220d7" />


## Month/Period Calculation

To dynamically display the correct pay period based on an employee’s pay frequency, I implemented the following formula:

```excel
=IF(Q2="Monthly", 
     TEXT(AO2,"mmmm yyyy"),
 IF(Q2="Weekly", 
     "Week " & WEEKNUM(AO2) & ", " & YEAR(AO2),
 IF(Q2="Bi-Weekly", 
     "Bi-Week: " & TEXT(AO2 - MOD(DAY(AO2)-1,14),"mmm d, yyyy") & " - " & TEXT(AO2 - MOD(DAY(AO2)-1,14)+13,"mmm d, yyyy"),
 "")))
```

Breakdown:

The formula adjusts the output based on pay frequency:

* If Monthly, it displays the full month and year.
* If Weekly, it displays the week number and year.
* If Bi-Weekly, it displays the exact 14-day range.
* If none of these apply, the cell remains blank.

This setup ensures the Month/Period column always reflects the correct payroll cycle automatically.

<img width="868" height="390" alt="Image" src="https://github.com/user-attachments/assets/e41cd0c5-c502-4779-9fb3-97bddf5c1a13" />

## Sheets Protection

I protected the sheet by setting it to display a warning before any edits, ensuring no one accidentally disrupts our functions.

<img width="787" height="335" alt="Image" src="https://github.com/user-attachments/assets/b31f7fd0-d014-407c-ab60-64c083f38c24" />







