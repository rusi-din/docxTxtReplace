using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace docxTxtReplace
{

    class Program
    {

        static int Main(string[] args)
        {
            string? filePath = null;
            string? search = null;
            string? replace = null;
            string? outputPath = null;
            string? filePass = null;


            foreach (var item in args)
            {
                if (item == "-file")
                {
                    int itemIndex = args.ToList().IndexOf(item) + 1;
                    string currentPath = Directory.GetCurrentDirectory();
                    Console.WriteLine(currentPath);


                    if (args[itemIndex][0] == '.')
                    {
                        string userpath = args[itemIndex];
                        userpath = userpath.Remove(0, 1);
                        filePath = currentPath + userpath;

                        Console.WriteLine(filePath);
                    }
                    else
                    {
                        filePath = args[itemIndex];
                        Console.WriteLine(filePath);
                    }

                    if (File.Exists(filePath))
                    {
                        string ext = Path.GetExtension(filePath);
                        Console.WriteLine("extension: " + ext);

                        if (ext == ".docx")
                        {
                            Console.WriteLine("File is .docx");
                        }
                        else
                        {
                            Console.WriteLine("File is not .docx");
                            return 2;
                        }
                        Console.WriteLine("File exists");
                    }
                    else
                    {
                        Console.WriteLine("File does not exist");
                        return 1;
                    }
                }
                else if (item == "-search")
                {
                    int itemIndex = args.ToList().IndexOf(item) + 1;
                    search = args[itemIndex];
                    Console.WriteLine(search);
                }
                else if (item == "-replace")
                {
                    int itemIndex = args.ToList().IndexOf(item) + 1;
                    replace = args[itemIndex];
                    Console.WriteLine(replace);
                }
                else if (item == "-output")
                {
                    int itemIndex = args.ToList().IndexOf(item) + 1;
                    string currentPath = Directory.GetCurrentDirectory();
                    Console.WriteLine(currentPath);

                    if (args[itemIndex][0] == '.')
                    {
                        string userpath = args[itemIndex];
                        userpath = userpath.Remove(0, 1);
                        outputPath = currentPath + userpath;

                        Console.WriteLine(outputPath);
                    }
                    else
                    {
                        outputPath = args[itemIndex];
                        Console.WriteLine(outputPath);
                    }

                    string ext = Path.GetExtension(outputPath);

                    if (ext == ".docx")
                    {
                        Console.WriteLine("File is .docx");
                    }
                    else
                    {
                        Console.WriteLine("File is not .docx");
                        Console.WriteLine("Please add the extension .docx");
                        return 3;
                    }


                }
                else if (item == "-pass")
                {
                    int itemIndex = args.ToList().IndexOf(item) + 1;
                    filePass = args[itemIndex];
                    Console.WriteLine(filePass);
                    Console.WriteLine("File encrypting and decrypting is not supported yet");
                }
            }



            if (filePath == null)
            {
                Console.WriteLine("No file path provided");
                return 4;
            }

            if (search == null)
            {
                Console.WriteLine("No search string provided");
                return 5;
            }

            if (replace == null)
            {
                Console.WriteLine("No replace string provided");
                return 6;
            }

            if (outputPath == null)
            {
                Console.WriteLine("No output path provided, file path will be used as output path");
                outputPath = filePath;
            }
            else
            {
                Console.WriteLine("output path: " + outputPath);
                File.Copy(filePath, outputPath, true);
            }




            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
            {
                // Access the main document part (the main text of the document)
                MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

                // Iterate over all paragraphs in the document
                foreach (var para in mainPart.Document.Body.Elements<Paragraph>())
                {
                    // Iterate over all runs in the paragraph
                    foreach (var run in para.Elements<Run>())
                    {
                        // Iterate over all Text elements in the run
                        foreach (var text in run.Elements<Text>())
                        {
                            // If the text contains the text to replace
                            if (text.Text.Contains(search))
                            {
                                // Replace the text
                                text.Text = text.Text.Replace(search, replace);
                            }
                        }
                    }
                }

                // Save the changes to the document
                mainPart.Document.Save();
            }




            return 000000;
        }
    }
}