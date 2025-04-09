using Excel_Reader.Control;
using Spectre.Console;

namespace Excel_Reader.UI;

internal static class View
{
    public static void PrintTable(List<Dictionary<string, object>> dbObjects)
    {
        Helpers.Logger($"Creating display table....\n\n", ConsoleColor.Green);
        var table = new Table { Border = TableBorder.Rounded, BorderStyle = new Style(Color.Green) };

        foreach (var key in dbObjects[0].Keys)
        {
            table.AddColumn(key);
        }

        foreach (var obj in dbObjects)
        {
            var rowValues = new List<string>();
            foreach (var value in obj.Values)
            {
                rowValues.Add(value.ToString());
            }
            table.AddRow(rowValues.ToArray());
        }
        AnsiConsole.Write(table);
        Console.WriteLine("\n\n");
    }
}
