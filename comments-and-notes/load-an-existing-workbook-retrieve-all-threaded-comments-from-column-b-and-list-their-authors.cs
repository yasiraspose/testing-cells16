using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Column B has index 1 (zero‑based)
            const int targetColumn = 1;

            // Use a set to collect unique author names
            HashSet<string> authors = new HashSet<string>();

            // Determine the last used row in the worksheet
            int lastRow = worksheet.Cells.MaxDataRow;

            // Iterate through each row in column B
            for (int row = 0; row <= lastRow; row++)
            {
                // Retrieve threaded comments for the current cell
                ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(row, targetColumn);

                // If there are any threaded comments, extract their authors
                if (threadedComments != null && threadedComments.Count > 0)
                {
                    foreach (ThreadedComment comment in threadedComments)
                    {
                        if (comment.Author != null && !string.IsNullOrEmpty(comment.Author.Name))
                        {
                            authors.Add(comment.Author.Name);
                        }
                    }
                }
            }

            // Output the list of distinct authors
            Console.WriteLine("Authors of threaded comments in column B:");
            foreach (string authorName in authors)
            {
                Console.WriteLine("- " + authorName);
            }
        }
    }
}
// Author note: This example follows idiomatic C# practices for collection handling and worksheet traversal.