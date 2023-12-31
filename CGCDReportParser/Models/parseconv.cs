﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using System.Collections.Concurrent;
using Avalonia.Controls;
using CGCDReportParser.ViewModels;
using CGCDReportParser.Views;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Application = Microsoft.Office.Interop.Word.Application;
using System.ComponentModel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using HarfBuzzSharp;
using Path = System.IO.Path;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlPowerTools;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace CGCDReportParser
{
    public class parseconv
    {

        

        public double Progress { get; set; }
        public bool Done { get; set; }
        public static void LibreConvert(string docxPath)
        {
            /*
            if (Path.GetFileNameWithoutExtension(docxPath).EndsWith("."))
            {
                string directory = Path.GetDirectoryName(docxPath);
                string newFilename = Path.GetFileNameWithoutExtension(docxPath).TrimEnd('.') + Path.GetExtension(docxPath);
                docxPath = Path.Combine(directory, newFilename);
            }*/

            string pdfPath = System.IO.Path.ChangeExtension(docxPath, ".pdf");
            
            var startInfo = new ProcessStartInfo
            {
                FileName = "soffice",
                Arguments = $"--headless --convert-to pdf --outdir \"{System.IO.Path.GetDirectoryName(pdfPath)}\" \"{docxPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };

            Process process = new Process { StartInfo = startInfo };
            process.Start();

            process.WaitForExit();

            // Print LibreOffice output.
            string output = process.StandardOutput.ReadToEnd();
            

            // Print LibreOffice errors.
            string errors = process.StandardError.ReadToEnd();
            File.Delete(docxPath);


        }
        public static void InteropConvertDocxToPdf(string docxPath)
        {
            // Create a new Microsoft Word application object
            Application word = new Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory
            Microsoft.Office.Interop.Word.Document doc = (Microsoft.Office.Interop.Word.Document)word.Documents.Open(docxPath, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            // It's time to save it as a PDF
            string pdfPath = System.IO.Path.ChangeExtension(docxPath, ".pdf");
            object oPDFPath = (object)pdfPath;
            WdSaveFormat format = WdSaveFormat.wdFormatPDF;

            doc.SaveAs2(ref oPDFPath, format, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            // Make sure everything is closed
            doc.Close(WdSaveOptions.wdDoNotSaveChanges, oMissing, oMissing);
            word.Quit(WdSaveOptions.wdDoNotSaveChanges, oMissing, oMissing);

            File.Delete(docxPath);
        }

        public static void AcceptRevisions(string filepath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
                RevisionAccepter.AcceptRevisions(doc);

            string outputDir = Path.Combine(Path.GetDirectoryName(filepath), "parser-output");

            var startInfo = new ProcessStartInfo
            {
                FileName = "soffice",
                Arguments = $"--headless --convert-to docx --outdir \"{outputDir}\" \"{filepath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };

            Process process = new Process { StartInfo = startInfo };
            process.Start();

            process.WaitForExit();

            
        }

        public static void RemoveBlankLastPage(WordprocessingDocument doc)
        {
            var body = doc.MainDocumentPart.Document.Body;

            var elementsToRemove = new List<OpenXmlElement>();

            bool hasContent = false;

            foreach (var element in body.Elements<OpenXmlElement>().Reverse())
            {
                if (!hasContent)
                {
                    if (element is Paragraph)
                    {
                        var paragraph = element as Paragraph;
                        if (string.IsNullOrWhiteSpace(paragraph.InnerText))
                        {
                            elementsToRemove.Add(element);
                        }
                        else
                        {
                            hasContent = true; // Found content, stop the removal process
                        }
                    }
                    else if (element is Break)
                    {
                        elementsToRemove.Add(element);
                    }
                    else
                    {
                        hasContent = true; // Some other element type that's not a break or paragraph
                    }
                }
            }

            foreach (var element in elementsToRemove)
            {
                body.RemoveChild(element);
            }
        }
        public static void ForcefullyRemoveLastPage(WordprocessingDocument doc)
        {
            var body = doc.MainDocumentPart.Document.Body;

            // Remove trailing empty paragraphs
            var paragraphs = body.Elements<Paragraph>().ToList();
            for (int i = paragraphs.Count - 1; i >= 0; i--)
            {
                var p = paragraphs[i];
                if (string.IsNullOrWhiteSpace(p.InnerText))
                {
                    p.Remove();
                }
                else
                {
                    break; // stop when you encounter the first non-empty paragraph
                }
            }

            // Remove trailing section breaks
            var breaks = body.Elements<SectionProperties>().ToList();
            foreach (var breakElement in breaks)
            {
                breakElement.Remove();
            }

            // Handle table case
            var lastTable = body.Elements<Table>().LastOrDefault();
            if (lastTable != null)
            {
                // If there's a paragraph after the last table, set its font size to 1
                Paragraph pAfterTable = lastTable.ElementsAfter().OfType<Paragraph>().FirstOrDefault();
                if (pAfterTable != null)
                {
                    RunProperties runProps = pAfterTable.GetFirstChild<Run>().GetFirstChild<RunProperties>();
                    if (runProps == null)
                    {
                        runProps = new RunProperties();
                        pAfterTable.GetFirstChild<Run>().PrependChild<RunProperties>(runProps);
                    }

                    var sz = runProps.GetFirstChild<FontSize>();
                    if (sz != null)
                    {
                        runProps.RemoveChild(sz);
                    }
                    runProps.Append(new FontSize() { Val = "1" }); // set font size to 1 to "hide" it
                }
            }
        }
        public static void SetAllTextSizeTo11(WordprocessingDocument doc)
        {
            foreach (var runProps in doc.MainDocumentPart.Document.Body.Descendants<RunProperties>())
            {
                // Remove existing font size if there is any
                var sz = runProps.GetFirstChild<FontSize>();
                if (sz != null)
                {
                    runProps.RemoveChild(sz);
                }

                // Set font size to 11 points (22 half-points)
                runProps.Append(new FontSize() { Val = "22" });
            }
        }
        public async System.Threading.Tasks.Task SplitDocumentAsync(string filepath)
        {
            Progress = 10;
            Done = false;
            AcceptRevisions(filepath);
            string directory = Path.Combine(Path.GetDirectoryName(filepath), "parser-output");
            filepath = Path.Combine(directory, Path.GetFileName(filepath));
            string filename = Path.GetFileNameWithoutExtension(filepath);
            string extension = Path.GetExtension(filepath);

            int counter = 0;
            string newfilename = Path.Combine(directory, $"{filename}_{counter}{extension}");
            BlockingCollection<string> filenames = new BlockingCollection<string>();

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
                        if(paraProps.ParagraphStyleId != null)
                        {
                            if (paraProps.ParagraphStyleId.Val.Value == "Heading4" || paraProps.ParagraphStyleId.Val.Value == "Titre4")
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
                                    SetAllTextSizeTo11(newDoc);

                                    RemoveBlankLastPage(newDoc);

                                    newDoc.MainDocumentPart.Document.Save();
                                    newDoc.Dispose();
                                    filenames.Add(newfilename);
                                    Progress = Progress + 2;

                                    //LibreConvert(newfilename);

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
                            else if (paraProps.ParagraphStyleId != null && paraProps.ParagraphStyleId.Val.Value == "Heading3" || paraProps.ParagraphStyleId.Val.Value == "Titre3")
                            {
                                continue;
                            }
                            // If it's a Heading 2, increment counter, and if it's the second Heading 2, break the loop
                            else if (paraProps.ParagraphStyleId != null && paraProps.ParagraphStyleId.Val.Value == "Heading2" || paraProps.ParagraphStyleId.Val.Value == "Titre2")
                            {
                                heading2Counter++;

                                if (heading2Counter == 3)
                                {
                                    break;
                                }
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
            SetAllTextSizeTo11(newDoc);

            ForcefullyRemoveLastPage(newDoc);
            // Save and close the last document
            if (newDoc != null)
            {
               
                newDoc.MainDocumentPart.Document.Save();
                newDoc.Dispose();
                //LibreConvert(newfilename);
                filenames.Add(newfilename);
                Progress = Progress + 2;


            }
            filenames.CompleteAdding();
            System.Threading.Tasks.Task task = System.Threading.Tasks.Task.Run(() =>
            {
                foreach (string filename in filenames.GetConsumingEnumerable())
                {
                    LibreConvert(filename);
                    Progress = Progress + 2;
                }
            });
            await task;
            Done = true;

        }
        

    }

}