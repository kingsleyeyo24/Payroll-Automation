# Payroll-Automation

Please note that this is synthetic data and does not contain any real information about my clients or their company.

First, I duplicated the dataset. As data professionals, we often do this to preserve the original copy for reference in case anything goes wrong. Issues rarely occur, but keeping an untouched version is considered best practice.


<img width="1366" height="547" alt="Image" src="https://github.com/user-attachments/assets/eed4382d-0e53-4b7a-852a-b0028fd68223" />

This is the data we got. It’s ugly and messy, so we’re going to fix it.

Next, I converted the dataset into a table and applied a purple theme to match the company’s colors. I reformatted columns like **DateOfBirth, HireDate, TerminationDate, and LastPromotion** into proper date types and added placeholders where needed.

Still within **General Data**, I changed columns such as **JobLevel, EmploymentType, PayFrequency, BenefitPlan, and IsActive** into dropdown lists for easier data entry and consistency.

<img width="1366" height="484" alt="Image" src="https://github.com/user-attachments/assets/d2dc8a91-a63f-484b-9df9-1f10cc528da7" />

As you can see, our data is already starting to look much better.
The dataset is already quite clean except for a few blank cells, so we can skip extensive cleaning and move straight to calculations.


As requested by the client, I created three new sheets. The first is **Employee Master Data**, which contains mostly static personal and employment details that rarely change. The columns include:
**EmployeeID, FirstName, LastName, Email, PhoneNumber, DateOfBirth, Gender, MaritalStatus, SSN, JobTitle, JobLevel, Department, EmploymentType, HireDate, TerminationDate, CompanyHireDate, LastPromotionDate, StreetAddress, City, State, ZipCode, IBAN, RoutingNumber, and BankAccountNumber.**

For the **EmployeeID** column, I used the formula 
```excel
='General Data'!A2
```
to pull the ID directly from the **General Data** sheet, ensuring consistency across both sheets.

<img width="1361" height="423" alt="Image" src="https://github.com/user-attachments/assets/bac48513-6a6e-482d-9ede-3cfaf96234db" />

For the **FirstName** column, I used the formula

```excel
=INDEX('General Data'!B:B, MATCH(A2, 'General Data'!A:A, 0))
```

This finds the row where the **EmployeeID** in A2 appears in column A of the General Data sheet and returns the **FirstName** from column B of that row.

I used this **INDEX MATCH** method for multiple columns including EmployeeID, FirstName, LastName, Email, PhoneNumber, DateOfBirth, Gender, MaritalStatus, SSN, JobTitle, JobLevel, Department, EmploymentType, HireDate, StreetAddress, City, State, ZipCode, IBAN, RoutingNumber, and BankAccountNumber.

MATCH finds the correct row, and INDEX returns the value from the specified column in that row.

You might be thinking, why not use VLOOKUP? VLOOKUP can only look to the right of the lookup column. Even though EmployeeID is in the first column, inserting a new column like MiddleName would break the formulas. INDEX MATCH avoids this because it specifies which column to return, making it more flexible and reliable.

For the **TerminationDate** and **LastPromotionDate** columns, I did things a bit differently. Using the regular formula caused an error because some cells were blank, which makes sense since active employees don’t have a termination date. To fix this, I used

```excel
=IFERROR(INDEX('General Data'!O:O, MATCH(A2, 'General Data'!A:A, 0)),"")
```

This way, any empty values in the column return "" instead of an error.

<img width="1366" height="413" alt="Image" src="https://github.com/user-attachments/assets/0a353c8c-d23e-452e-a77e-d3d990d4e84f" />

As you can see, our HR data is looking nice.

<img width="1326" height="456" alt="Image" src="https://github.com/user-attachments/assets/f36f2d8b-0046-44da-8f5a-66c40e282f2d" />

All of this data is referenced from the General Data sheet, so General Data remains the ultimate source. Any changes made there will automatically be reflected here, as requested by the client.




Next, I created a sheet called **PayrollData**, which includes columns such as EmployeeID, FirstName, LastName, Email, PhoneNumber, DateOfBirth, Gender, MaritalStatus, SSN, JobTitle, JobLevel, Department, EmploymentType, HireDate, TerminationDate, AnnualSalary, PayFrequency, HourlyRate, BenefitPlan, BankAccountNumber, RoutingNumber, StreetAddress, City, State, ZipCode, CompanyHireDate, Column 1, IsActive, IBAN, CreditCardNumber, CreditCardProvider, LastPromotionDate, BonusPaidYTD, TaxWithheldYTD, RetirementContributionYTD, GrossPay, CurrentBonus, Tax, CurrentRetirementContribution, NetPay, DateProcessed, and Month/Period.

To pull data from **General Data** into a separate Payroll sheet, I wrote a script using **JavaScript** in **Apps Script**. The goal was to only bring in active employees, keeping the payroll data clean and up to date.

The script starts by accessing the spreadsheet and defining the source and target sheets:

```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();
var source = ss.getSheetByName('General Data');
var target = ss.getSheetByName('Sheet5');
```

It then grabs all the data from the source sheet and identifies the **IsActive** column. This column tells whether an employee is active. If it’s missing, the script stops because filtering would be impossible:

```javascript
var data = source.getDataRange().getValues();
var headers = data[0];
var isActiveCol = headers.indexOf("IsActive"); 
if (isActiveCol === -1) throw new Error("IsActive column not found.");
```

Next, it filters only the rows where **IsActive = "Yes"**, ensuring that only active employees are pulled:

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

<img width="1318" height="434" alt="Image" src="https://github.com/user-attachments/assets/090c2e36-c399-4dfe-b77b-bca944e8ac67" />



I created a column called **GrossPay**. Since the employees in this dataset are all salaried, their gross pay was calculated based on their pay frequency.

* For employees paid **monthly**, GrossPay is the **Annual Salary ÷ 12**.
* For employees paid **weekly**, GrossPay is the **Annual Salary ÷ 52**.
* For employees paid **bi-weekly**, GrossPay is the **Annual Salary ÷ 26**.


```excel
=IF(Q3="Monthly", P3/12, 
   IF(Q3="Weekly", P3/52, 
   IF(Q3="Bi-Weekly", P3/26, "")))
```

This way, GrossPay dynamically adjusts depending on each employee’s pay frequency.

<img width="1366" height="418" alt="Image" src="https://github.com/user-attachments/assets/477bde10-f32e-4ddc-b2f5-c62f75f9ae6a" />


