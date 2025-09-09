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

These sheet largely cover the deliverables requested by the client.









