# Employee Service Portal

An all-in-one employee self-service tool built with Google Apps Script for managing leave, assets, payroll, and payslip downloads.

---

## 1. Make a Copy of the Portal

Click the link below to create your own copy of the Employee Service Portal spreadsheet (this will directly prompt you to make a copy):  
[**Make a Copy of Employee Service Portal**](https://docs.google.com/spreadsheets/d/1d2ZuA9ZUxd_v6O6U9Ncr7jWnpOJuHtiuqheUQuH3f3A/copy)

---

## 2. Publish as a Web App

1. Go to **Extensions → Apps Script** in your copied spreadsheet.
2. Click **Deploy → New deployment**.
3. Select **Web app**.
4. Set **Execute as** → `Me` and **Who has access** → `Anyone` (or as required).
5. Deploy and copy the Web App URL for use.

---

## 3. Customize the Payslip Template

Make a copy of the payslip template to match your company branding and style:  
[**Make a Copy of Payslip Template**](https://docs.google.com/spreadsheets/d/1d2ZuA9ZUxd_v6O6U9Ncr7jWnpOJuHtiuqheUQuH3f3A/copy)

---

## 4. Configure Payslip Storage Path

In **`config.gs`**, set the variable `PAYSLIP_FOLDER_PATH` to the Google Drive folder path **ID** where all generated payslips should be stored.

---

## 5. Fill in Data

Populate the following tabs in the spreadsheet:
- **Employees** – Employee details and IDs
- **Leaves** – Leave records and balances
- **Assets** – Asset allocation and tracking
- **Payroll** – Salary components and payment details

---

## 6. Need Help?

For customization, setup assistance, or hiring services:  
📧 **support@tabgraf.com**  
🌐 [https://tabgraf.com](https://tabgraf.com)

---

© 2025 **Tabgraf Technologies LLP**. All rights reserved.
