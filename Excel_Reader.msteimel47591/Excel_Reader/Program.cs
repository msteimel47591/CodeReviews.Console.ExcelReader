using Excel_Reader.Control;
using Excel_Reader.UI;

namespace Excel_Reader;

class Program
{
    static void Main(string[] args)
    {
        Helpers.Logger("Starting the program...", ConsoleColor.Green);
        Operations.CreateExcelFile();
        Operations.CreateDatabase();

        try
        {
            Operations.CreateTable();
        }
        catch (Exception e)
        {

            Helpers.Logger("There was a problem creating the database.....", ConsoleColor.Red);
            Helpers.Logger($"{e.Message}", ConsoleColor.Red);
            Helpers.Logger("Exiting the program...", ConsoleColor.Red);
            Environment.Exit(1);
        }

        List<Dictionary<string, object>> dbObjects = Operations.ReadDatabase();

        if (dbObjects.Count > 0)
        {
            View.PrintTable(dbObjects);
        }
        else
        {
            Helpers.Logger("No data found in the database.", ConsoleColor.Red);
        }

        Helpers.Logger("Program completed successfully, exiting program...", ConsoleColor.Green);
    }
}