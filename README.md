# healthappvba
HealthApp VBA

HealthApp.xlsm

-	Double click file to open 
-	Enter password: healthapp23
-	To view code in editor: ALT + F11 in the app

-	HealthApp features:
-	Client Manager (Main interface - Click “CLIENT MANAGER” button to open app)
-	Userform - Client contact / joined information collected from “ADD CLIENT / CLIENT DATA” form. Input validation exists where all fields must be completed prior to form submittal. 

![ClientData Userform](https://github.com/jhchiu1/healthappvba/blob/master/ClientData.png)

-	Userform (continued) - Once “ADD CLIENT / CLIENT DATA” form is submitted, a new form will load with the next consecutive Client ID available. Userform has “EXIT” button to prevent user from closing the app using “X” in upper right corner.

![ClientData Form Refreshed](https://github.com/jhchiu1/healthappvba/blob/master/ClientData2.png)

-	ClientData spreadsheet - ClientIDFull field is concatenated using Client ID label, First Name, and Last Name fields.

![ClientData Spreadsheet](https://github.com/jhchiu1/healthappvba/blob/master/ClientDataSpreadsheet.png)

-	Health Data form - Select Client combobox populates using “ClientIDFull” - Column B from ClientData spreadsheet. 

![Health Data Userform](https://github.com/jhchiu1/healthappvba/blob/master/HealthData.png)

-	HealthData spreadsheet - Body Mass Index (BMI) column uses conditional formatting where red cells with values greater than 30 indicate obesity and orange cells with values greater than 25 indicate overweight. Balance Due column is calculated using Time column multiplied by trainer hourly rate of $20.00 / hour. The hourly rate can be adjusted in the formula.

![Pivot Table](https://github.com/jhchiu1/healthappvba/blob/master/Pivot.png)

-	Monthly Billing Pivot Table - Analyzes the following fields from HealthData spreadsheet: ClientID, Date, Time, Balance Due. Filter pivot table by ClientID, view all clients is an option as well. Pivot table allows trainer to easily view a client’s total number of visits by date, total time spent training, plus total balance due. The HealthData spreadsheet contains a named range that dynamically adjusts when the “REFRESH” button is clicked as well as applies the currency format  to Column C / Sum of Balance Due column. 
