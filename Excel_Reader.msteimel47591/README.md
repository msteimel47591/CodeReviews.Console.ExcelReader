# Excel Reader

Excel reader is a project for The C# Academy roadmap. It reads data from an excel spreadsheet and populates it into a Sqlite database and then reads the data from the database and displays it to the user.

## Project Requirements

- This is an application that will read data from an Excel spreadsheet into a database
- When the application starts, it should delete the database if it exists, create a new one, create all tables, read from Excel, seed into the database.
- You need to use EPPlus package
- You shouldn't read into Json first.
- You can use SQLite or SQL Server (or MySQL if you're using a Mac)
- Once the database is populated, you'll fetch data from it and show it in the console.
- You don't need any user input
- You should print messages to the console letting the user know what the app is doing at that moment (i.e. reading from excel; creating tables, etc)
- The application will be written for a known table, you don't need to make it dynamic.
- When submitting the project for review, you need to include an xls file that can be read by your application.


## Usage

When the application starts it will create a spreadsheet named DbInfo.xlsx in the project folder if the spreadsheet doesn't already exist and seed it with data. Once the spreadsheet exists you can modify it by adding rows/columns, renaming columns etc.
