using System;
using Aspose.Cells;

namespace WorkbookMergingConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook(); // using Workbook() constructor
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (empty workbook) and add some data
            Workbook destWorkbook = new Workbook(); // using Workbook() constructor
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "DestinationSheet";
            destSheet.Cells["B2"].PutValue("Destination Data");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook); // using Workbook.Combine method

            // Save the combined workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx); // using Workbook.Save method

            Console.WriteLine($"Workbooks combined and saved to '{outputPath}'.");
        }
    }
}