using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a regular comment to cell A1
            int commentIdx = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIdx];
            comment.Note = "Root comment";

            // Access the threaded comment collection of the comment
            ThreadedCommentCollection threadedComments = comment.ThreadedComments;

            // Get the collection of authors for threaded comments
            ThreadedCommentAuthorCollection authors = workbook.Worksheets.ThreadedCommentAuthors;

            // Add an author (if not already present)
            int authorIdx = authors.Add("John Doe", "john.doe@example.com", "ID001");
            ThreadedCommentAuthor author = authors[authorIdx];

            // Add some threaded comments
            threadedComments.Add("First reply", author);
            threadedComments.Add("Second reply", author);

            // Iterate through the threaded comments and display details
            for (int i = 0; i < threadedComments.Count; ++i)
            {
                ThreadedComment tc = threadedComments[i];
                Console.WriteLine($"Comment #{i + 1}");
                Console.WriteLine($"  Text   : {tc.Notes}");
                Console.WriteLine($"  Author : {tc.Author.Name}");
                Console.WriteLine($"  Created: {tc.CreatedTime}");
                Console.WriteLine();
            }

            // Save the workbook (optional)
            workbook.Save("ThreadedCommentsDemo.xlsx");
        }
    }
}

// Author note: This example demonstrates how to traverse a ThreadedCommentCollection
// and output each comment's text, author name, and creation timestamp.