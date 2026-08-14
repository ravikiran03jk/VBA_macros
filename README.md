# Poultry Farm Equipment Inspection Tracker – Excel VBA

## Project Overview

This project is an Excel VBA automation tool designed to monitor equipment inspection dates in a poultry farm environment.

The macro automatically checks the **Last Checked** date of each machine and highlights the entire row when the equipment has not been inspected for more than **60 days**.

This helps maintenance teams quickly identify equipment that may require inspection or follow-up.

## Objective

The main objective of this project is to automate the identification of overdue equipment inspections and reduce the need for manually reviewing inspection dates.

## Dataset

The Excel dataset contains equipment-related information such as:

* Item Name
* Category
* Part ID
* Location
* Manufacturer
* Purchase Year
* Last Checked

## How the VBA Automation Works

1. The macro starts from the equipment data.
2. It reads the **Last Checked** date for each machine.
3. It compares the date with the current date.
4. It calculates the number of days since the last inspection.
5. If the equipment has not been checked for more than **60 days**, the complete row is highlighted.
6. Equipment requiring attention can therefore be identified quickly.

## Key VBA Concepts Used

* VBA loops
* `Range` and `Offset`
* Date calculations
* Conditional logic using `If...Then`
* Dynamic row processing
* Excel cell formatting
* Automated equipment monitoring

## Business Value

This automation can help:

* Reduce manual inspection-date checking
* Identify overdue equipment quickly
* Improve maintenance follow-up
* Reduce the possibility of missed inspections
* Make equipment monitoring more efficient

## Tools Used

* Microsoft Excel
* VBA (Visual Basic for Applications)

## Project Outcome

The final Excel workbook provides an automated way to identify equipment that has exceeded the 60-day inspection interval by highlighting the relevant rows.


