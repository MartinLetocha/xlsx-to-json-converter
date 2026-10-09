using Converter.Lib;

namespace Converter.App;

class Program
{
    static void Main(string[] args)
    {
        var converter = new XlsxToJsonConverter();
        var deepDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
        var directory = deepDirectory.Parent.Parent.Parent;
        var spreadsheet = directory.GetFiles("*.xlsx").FirstOrDefault();

        if (spreadsheet == null)
        {
            Console.WriteLine("There is no .xlsx file in the 'Converter.App' directory.");
            Console.ReadKey();
            return;
        }

        string json;
        using (FileStream fileStream = new FileStream(spreadsheet.FullName, FileMode.Open))
        {
            json = converter.Convert(fileStream);
        }

        File.WriteAllText($"{directory.FullName}\\Result.json", json);
        Console.WriteLine(json);
        Console.ReadKey();
    }
}