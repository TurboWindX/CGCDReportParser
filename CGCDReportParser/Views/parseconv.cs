using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


namespace avalonia_docxpdf
{
    internal class parseconv
    {
        public static void InstallDocx2Pdf()
        {
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = "python";
            start.Arguments = "-m pip install docx2pdf";
            start.UseShellExecute = false; //no system shell.
            start.RedirectStandardOutput = true; // Any output
            start.RedirectStandardError = true; // Any error 
            start.CreateNoWindow = true; // Don't create a new window

            using (Process process = Process.Start(start))
            {
                string stderr = process.StandardError.ReadToEnd();
                process.WaitForExit();
            }
        }

        public static void ConvertDocxToPdf(string docxPath)
        {
            // Escape backslashes in the paths.
            string docxPathEscaped = docxPath.Replace("\\", "\\\\");
            string pdfPathEscaped = System.IO.Path.ChangeExtension(docxPath, ".pdf").Replace("\\", "\\\\");

            // Define Python script.
            string pythonScript = $"import sys\nfrom docx2pdf import convert\nconvert('{docxPathEscaped}', '{pdfPathEscaped}')";

            // Write Python script to a temporary file.
            string tmpPythonScriptPath = System.IO.Path.GetTempFileName() + ".py";
            System.IO.File.WriteAllText(tmpPythonScriptPath, pythonScript);

            // Setup Python process.
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "python", // or the full path of your python interpreter
                Arguments = tmpPythonScriptPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };

            // Start Python process.
            Process process = new Process { StartInfo = processStartInfo };
            process.Start();
            process.WaitForExit();

            // Print Python output.
            string output = process.StandardOutput.ReadToEnd();
            Console.WriteLine(output);

            // Print Python errors.
            string errors = process.StandardError.ReadToEnd();
            if (!string.IsNullOrEmpty(errors))
            {
                Console.WriteLine(errors);
            }
        }
        public static void SplitDocument(string filepath)
        {
            string directory = Path.GetDirectoryName(filepath);
            string filename = Path.GetFileNameWithoutExtension(filepath);
            string extension = Path.GetExtension(filepath);
            int counter = 0;
            string newfilename = Path.Combine(directory, $"{filename}_{counter}{extension}");

            List<OpenXmlElement> paragraphs = null;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, true))
            {
                paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements().ToList();
            }

            WordprocessingDocument newDoc = null;
            DocumentFormat.OpenXml.Wordprocessing.Body newDocBody = null;
            int heading2Counter = 0; // Counter for Heading 2

            foreach (var element in paragraphs)
            {
                Paragraph para = element as Paragraph;

                if (para != null)
                {
                    ParagraphProperties paraProps = para.GetFirstChild<ParagraphProperties>();

                    if (paraProps != null)
                    {
                        // Check if it's a Heading 4
                        if (paraProps.ParagraphStyleId != null && paraProps.ParagraphStyleId.Val.Value == "Heading4")
                        {
                            string heading4Text = para.InnerText;

                            // Replace any characters that are not valid in file names
                            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                            {
                                heading4Text = heading4Text.Replace(c, '_');
                            }

                            // Save and close current document
                            if (newDoc != null)
                            {
                                newDoc.MainDocumentPart.Document.Save();
                                newDoc.Dispose();
                                ConvertDocxToPdf(newfilename);
                            }

                            //newfilename = Path.Combine(directory, $"{filename}_{counter}{extension}");
                            newfilename = Path.Combine(directory, $"{heading4Text}{extension}");

                            File.Copy(filepath, newfilename, true);

                            newDoc = WordprocessingDocument.Open(newfilename, true);
                            newDocBody = newDoc.MainDocumentPart.Document.Body;

                            // Clear the new document
                            newDocBody.RemoveAllChildren();

                            counter++;
                        }
                        // If it's a Heading 3, skip it
                        else if (paraProps.ParagraphStyleId != null && paraProps.ParagraphStyleId.Val.Value == "Heading3")
                        {
                            continue;
                        }
                        // If it's a Heading 2, increment counter, and if it's the second Heading 2, break the loop
                        else if (paraProps.ParagraphStyleId != null && paraProps.ParagraphStyleId.Val.Value == "Heading2")
                        {
                            heading2Counter++;

                            if (heading2Counter == 3)
                            {
                                break;
                            }
                        }
                    }
                }

                // If we have an open document, import the paragraph
                if (newDoc != null)
                {
                    newDocBody.Append(element.CloneNode(true));
                }

            }

            // Save and close the last document
            if (newDoc != null)
            {
                newDoc.MainDocumentPart.Document.Save();
                newDoc.Dispose();
            }
            ConvertDocxToPdf(newfilename);

        }

    }

}