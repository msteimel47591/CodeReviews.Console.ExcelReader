using Excel_Reader.Models;
using Microsoft.Data.Sqlite;
using OfficeOpenXml;

namespace Excel_Reader.Control;

internal static class Operations
{
    public static void CreateExcelFile()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Non-Commercial");

        string excelFilePath = Path.Combine(Helpers.GetPath(), "DbInfo.xlsx");

        Helpers.Logger("Checking if DbInfo.xlsx exists...", ConsoleColor.Green);

        if (!File.Exists(excelFilePath))
        {
            Helpers.Logger("DbInfo.xlsx does not exist. Creating a new file...", ConsoleColor.Green);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Occupation";
                worksheet.Cells[1, 3].Value = "Years With Company";

                List<Employee> employees = Helpers.GetEmployees();

                for (int i = 0; i < employees.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = employees[i].Name;
                    worksheet.Cells[i + 2, 2].Value = employees[i].Occupation;
                    worksheet.Cells[i + 2, 3].Value = employees[i].YearsWithCompany;
                }
                worksheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(excelFilePath));

                Helpers.Logger("DbInfo.xlsx created successfully.", ConsoleColor.Green);

                return;
            }
        }
        Helpers.Logger("DbInfo.xlsx already exists...", ConsoleColor.Green);
    }

    public static void CreateDatabase()
    {
        string databaseFilePath = Path.Combine(Helpers.GetPath(), "ExcelReader.db");

        if (File.Exists(databaseFilePath))
        {
            Helpers.Logger("Deleting existing database...", ConsoleColor.Green);
            File.Delete(databaseFilePath);
        }

        Helpers.Logger("Creating a new database...", ConsoleColor.Green);

        using (var connection = new SqliteConnection($"Data Source={databaseFilePath};"))
        {
            connection.Open();
            Helpers.Logger("Database file created successfully.", ConsoleColor.Green);
        }
    }

    public static void CreateTable()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Non-Commercial");
        string basePath = Helpers.GetPath();
        string excelFilePath = Path.Combine(basePath, "DbInfo.xlsx");
        string databaseFilePath = Path.Combine(basePath, "ExcelReader.db");

        Helpers.Logger("Creating table in the database...", ConsoleColor.Green);

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            if (worksheet == null || worksheet.Dimension == null)
            {
                throw new InvalidOperationException("The worksheet is empty or does not have any data.");
            }

            int totalColumns = worksheet.Dimension.Columns;
            int totalRows = worksheet.Dimension.Rows;

            string tableName = "DynamicTable";
            var createTableCmd = $"CREATE TABLE {tableName} (";
            for (int col = 1; col <= totalColumns; col++)
            {
                string columnName = worksheet.Cells[1, col].Text;
                columnName = $"\"{columnName}\"";
                createTableCmd += $"{columnName} TEXT";
                if (col < totalColumns) createTableCmd += ", ";
            }
            createTableCmd += ");";

            using (var connection = new SqliteConnection($"Data Source={databaseFilePath};"))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = createTableCmd;
                    command.ExecuteNonQuery();
                }

                for (int row = 2; row <= totalRows; row++)
                {
                    var insertCmd = $"INSERT INTO {tableName} VALUES (";
                    for (int col = 1; col <= totalColumns; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text.Replace("'", "''");
                        insertCmd += $"'{cellValue}'";
                        if (col < totalColumns) insertCmd += ", ";
                    }
                    insertCmd += ");";

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = insertCmd;
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }

    public static List<Dictionary<string, object>> ReadDatabase()
    {
        string databaseFilePath = Path.Combine(Helpers.GetPath(), "ExcelReader.db");
        var rows = new List<Dictionary<string, object>>();

        Helpers.Logger("Reading data from the database...", ConsoleColor.Green);

        using (var connection = new SqliteConnection($"Data Source={databaseFilePath};"))
        {
            connection.Open();
            string query = "SELECT * FROM DynamicTable";
            using (var command = new SqliteCommand(query, connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var row = new Dictionary<string, object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[reader.GetName(i)] = reader.GetValue(i);
                        }
                        rows.Add(row);
                    }
                }
            }

            return rows;
        }
    }
}
