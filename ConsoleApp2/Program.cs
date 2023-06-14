using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the path to the input DOCX file:");
        string docxFilePath = Console.ReadLine(); // Read the input DOCX file path from the console
        string outputFilePath = "output.xlsx"; // Path for the output Excel file
         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        // Read the DOCX file
        List<string> lines = ReadDocxFile(docxFilePath);

        // Create a new Excel package
        using (ExcelPackage package = new ExcelPackage())
        {
            // Create a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

            int row = 1; // Start from the first row

            // Process the lines
            for (int i = 0; i < lines.Count; i += 6)
            {
                // Get the question (which can be multiple lines)
                List<string> questionLines = new List<string> ();

                // Check if the question spans multiple lines
                if (i + 5 < lines.Count && !lines[i + 5].StartsWith("A)"))
                {
                    questionLines.Add(lines[i]);
                }
                Console.WriteLine(questionLines.Count);
                string question = string.Join(Environment.NewLine, questionLines);

                string optionA = "A)" + GetOption(lines, i + 1);
                string optionB = "B)" + GetOption(lines, i + 2);
                string optionC = "C)" + GetOption(lines, i + 3);
                string optionD = "D)" + GetOption(lines, i + 4);
                string optionE = "E)" + GetOption(lines, i + 5);

                // Write the question and options to the Excel worksheet
                worksheet.Cells[row, 1].Value = $"{i/6+1}. {question}";
                worksheet.Cells[row, 2].Value = optionA;
                worksheet.Cells[row, 3].Value = optionB;
                worksheet.Cells[row, 4].Value = optionC;
                worksheet.Cells[row, 5].Value = optionD;
                worksheet.Cells[row, 6].Value = optionE;

                row++; // Move to the next row
            }

            // Save the Excel package to a file
            package.SaveAs(new FileInfo(outputFilePath));
        }

        Console.WriteLine("Excel file created successfully.");
        Console.ReadLine();
    }

    static List<string> ReadDocxFile(string filePath)
    {
        List<string> lines = new List<string>();

        using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            Body body = mainPart.Document.Body;

            foreach (Paragraph paragraph in body.Elements<Paragraph>())
            {
                string text = paragraph.InnerText.Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    lines.Add(text);
                }
            }
        }

        return lines;
    }

    static string GetOption(List<string> lines, int index)
    {
        if (index < lines.Count)
        {
            return lines[index];
        }

        return string.Empty;
    }
}
