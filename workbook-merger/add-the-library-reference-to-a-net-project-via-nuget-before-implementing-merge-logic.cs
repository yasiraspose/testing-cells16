using System;
using Aspose.Cells;

namespace MergeDemo
{
    class Program
    {
        // Before running this code, add the Aspose.Cells library to the project via NuGet:
        //   Install-Package Aspose.Cells
        // or using the .NET CLI:
        //   dotnet add package Aspose.Cells

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put some sample data
            cells[0, 0].PutValue("Merged Header");
            cells[1, 0].PutValue("Row 2");
            cells[2, 0].PutValue("Row 3");

            // Merge cells from row 0, column 0 spanning 3 rows and 2 columns
            cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 3, totalColumns: 2);

            // Save the workbook
            workbook.Save("MergedOutput.xlsx");
        }
    }
}