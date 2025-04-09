using Excel_Reader.Models;

namespace Excel_Reader.Control;

internal static class Helpers
{
    public static string GetPath()
    {
        Logger("Getting file path...", ConsoleColor.Green);

        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string solutionPath = Directory.GetParent(baseDirectory).Parent.Parent.Parent.Parent.FullName;

        return solutionPath;
    }

    public static List<Employee> GetEmployees()
    {
        var employees = new List<Employee>();

        Employee employee = new Employee { Name = "John Doe", Occupation = "Software Engineer", YearsWithCompany = 5 };
        employees.Add(employee);
        employee = new Employee { Name = "Jane Smith", Occupation = "Project Manager", YearsWithCompany = 3 };
        employees.Add(employee);
        employee = new Employee { Name = "Sam Brown", Occupation = "Data Analyst", YearsWithCompany = 2 };
        employees.Add(employee);
        employee = new Employee { Name = "Lisa White", Occupation = "UX Designer", YearsWithCompany = 4 };
        employees.Add(employee);
        employee = new Employee { Name = "Tom Green", Occupation = "DevOps Engineer", YearsWithCompany = 1 };
        employees.Add(employee);

        return employees;
    }

    public static void Logger(string message, ConsoleColor color)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(message);
        Console.ResetColor();
    }
}
