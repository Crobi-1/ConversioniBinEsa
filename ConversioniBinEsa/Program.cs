
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string filePath = "Conversioni.xlsx";

        Console.WriteLine("Inserisci un numero binario di 8 cifre:");
        string binario = Console.ReadLine();

        int decimale = 0;

        for (int i = 0; i < 8; i++)
        {
            decimale = decimale * 2 + (binario[i] - '0');
        }

        string esadecimale = decimale.ToString("X");

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Conversioni");

            worksheet.Cell(1, 1).Value = "BINARIO";
            worksheet.Cell(1, 2).Value = "DECIMALE";
            worksheet.Cell(1, 3).Value = "ESADECIMALE";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.Yellow;
            headerRow.Style.Font.FontColor = XLColor.Blue;

            worksheet.Cell(2, 1).Value = binario;
            worksheet.Cell(2, 2).Value = decimale;
            worksheet.Cell(2, 3).Value = esadecimale;

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);

            Console.WriteLine("File Excel creato con successo!");
        }
    }
}
