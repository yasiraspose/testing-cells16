using System;
using Aspose.Cells;

namespace WorkbookMergingExample
{
    public class Merger
    {
        /// <summary>
        /// Merges multiple Excel files into a single workbook.
        /// Each source file is loaded using the Workbook(string) constructor,
        /// then combined into the destination workbook via Workbook.Combine.
        /// </summary>
        /// <param name="sourceFiles">Array of file paths to merge.</param>
        /// <param name="outputFile">Path where the merged workbook will be saved.</param>
        public static void MergeWorkbooks(string[] sourceFiles, string outputFile)
        {
            // Create an empty destination workbook.
            // This uses the default constructor (Workbook()).
            Workbook destination = new Workbook();

            // Iterate over each source file, load it, and combine it into the destination.
            foreach (string filePath in sourceFiles)
            {
                // Load the source workbook using the constructor that accepts a file path.
                Workbook source = new Workbook(filePath);

                // Combine the source workbook into the destination workbook.
                // The Combine method merges worksheets, styles, etc.
                destination.Combine(source);
            }

            // Save the merged workbook to the specified output path.
            // SaveFormat.Xlsx ensures the file is saved in the modern Excel format.
            destination.Save(outputFile, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            // Define the files to be merged.
            string[] filesToMerge = new string[]
            {
                "Report_January.xlsx",
                "Report_February.xlsx",
                "Report_March.xlsx"
            };

            // Define the output file path.
            string mergedFile = "QuarterlyReport.xlsx";

            // Perform the merge.
            MergeWorkbooks(filesToMerge, mergedFile);

            Console.WriteLine($"Merged workbook saved to: {mergedFile}");
        }
    }
}