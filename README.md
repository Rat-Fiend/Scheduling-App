# Scheduling-App
A scheduling application, including a makeshift database, made in excel.

## Set-up
In order to use this you will first need to go to the "DatabaseVarsAndMethods" module in the VBA editor (Alt + F11) and change the JOB_DATABASE_PATH and JOB_DATABASE_REFERENCE_PATH to the actual path of your instance of the database workbook.

If you feel like looking through the code, I recommend you download and instal the [Rubberduck plugin](https://rubberduckvba.com/). It allows for a seudo file structure and some other features.

## Usage
This app is meant to be used to schedule manufacturing procedures. You can create parts, product orders, parts of product orders, and jobs that correlate to those product order parts. There is a bar chart to display it all and some other functionality that is nice for scheduling. You can also create copies of the client file and work on the same database at the same time. Just be sure to save and load frequently.
