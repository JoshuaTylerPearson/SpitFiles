/*
 * 
 * SpitFiles
 * regex based PDF spliter
 * by Joshua Pearson
 * started: 2020-03-01 
 * 
 */

using System;
using System.IO;
using System.Text.RegularExpressions;
using CommandLine;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace SpitFiles
{

    class Options
    {
        [Option('v', "verbose", Required = false, HelpText = "Output verbose process info.")]
        public bool verbose { get; set; }

        [Option('i', "input", Required = true, HelpText = "Relative path and name of input PDF file to split.")]
        public string inputFilename { get; set; }

        [Option('o', "output", Required = true, HelpText = "Relative path to send output files to.")]
        public string outputDirname { get; set; }

        [Option('m', "matching-key-split", Required = false, HelpText = "Split pages based on having the same regex match.")]
        public string keySplitRegex { get; set; }

        [Option('d', "dated", Required = false, HelpText = "Start output filenames with dates.")]
        public bool datedFileNames { get; set; }

        [Option('e', "explicit", Required = false, HelpText = "Explicitly show regex groups for debugging.")]
        public bool showRegexMatches { get; set; }

    }


    class Program
    {

        static int Main(string[] args)
        {

            #region var declaration
            bool verbose = false;
            bool datedFileNames = false;
            bool showRegexMatches = false;
            string inputFilename = "";
            bool inputFileExists;
            string extension;
            bool isPDF = false;
            string outputDirname = "";
            bool outputDirExsists = false;

            uint splitType = 0x0;

            string keySplitRegex = "";
            #endregion

            #region arg intake
            // intake args and load values in to scope
            var result = Parser.Default.ParseArguments<Options>(args);

            result.WithParsed<Options>(o =>
            {
                inputFilename = o.inputFilename;

                outputDirname = o.outputDirname;


                if (o.verbose)
                {
                    verbose = true;
                }

                if (o.datedFileNames)
                {
                    datedFileNames = true;
                }

                if (o.showRegexMatches)
                {
                    showRegexMatches = true;
                }

            if (!String.IsNullOrEmpty(o.keySplitRegex) && o.keySplitRegex.Length > 0)
                {
                    splitType = splitType | 0x1;
                    keySplitRegex = o.keySplitRegex;
                }
            });

            #endregion

            #region verify input

            // input file location
            if (verbose) Console.WriteLine("Input File:\t" + inputFilename);
            inputFileExists = File.Exists(inputFilename);
            if (verbose) Console.WriteLine(inputFileExists ? "File exists:\tTrue" : "File exists:\tFalse");
            if (!inputFileExists)
            {
                if (verbose) Console.WriteLine("Input File does not exsist; Exiting with error code 1.");
                #if DEBUG
                Console.ReadKey();
                #endif
                return 1;
            }

            // input file format
            extension = System.IO.Path.GetExtension(inputFilename).ToLower();
            if (verbose) Console.WriteLine("File format:\t" + extension);
            isPDF = string.Equals(extension, ".pdf");
            if (verbose) Console.WriteLine(isPDF ? "Correct Format:\tTrue" : "Correct Format:\tFalse");
            if (!isPDF)
            {
                if (verbose) Console.WriteLine("Input File is not a PDF; Exiting with error code 2.");
                #if DEBUG
                Console.ReadKey();
                #endif
                return 2;
            }

            // output directory exsistance
            if (verbose) Console.WriteLine("Output to:\t" + outputDirname);
            outputDirExsists = Directory.Exists(outputDirname);
            if (verbose) Console.WriteLine(outputDirExsists ? "Output valid:\tTrue" : "Output valid:\tFalse");
            if (!outputDirExsists)
            {
                if (verbose) Console.WriteLine("Output dir does not exsist; Exiting with error code 3.");
                #if DEBUG
                Console.ReadKey();
                #endif
                return 3;
            }


            #endregion

            // Split
            switch (splitType)
            {
                case 0x1: // key match
                    if (verbose) Console.WriteLine("split type:\tKey");

                    if (verbose) Console.WriteLine("Key regex:\t" + keySplitRegex);

                    Regex regex = new Regex(keySplitRegex, RegexOptions.Compiled | RegexOptions.Multiline);

                    PdfReader reader = new PdfReader(inputFilename);

                    PdfReaderContentParser parser = new PdfReaderContentParser(reader);

                    string regexKeyMatch = "";
                    int docPageStart = 1;
                    string newDocName = "";

                    for (int page = 1; page <= reader.NumberOfPages; page++)
                    {

                        if (showRegexMatches) Console.WriteLine("Page: " + page);

                        ITextExtractionStrategy strategy = parser.ProcessContent
                            (page, new SimpleTextExtractionStrategy());

                        int matchCount = 0;

                        Match match = regex.Match(strategy.GetResultantText());
                        {
                            if (showRegexMatches) Console.WriteLine("Match: " + (++matchCount));
                            for (int x = 1; x <= 2; x++)
                            {
                                Group group = match.Groups[x];
                                if (showRegexMatches) Console.WriteLine("Group " + x + " = '" + group + "'");
                                CaptureCollection cc = group.Captures;
                                for (int y = 0; y < cc.Count; y++)
                                {
                                    Capture capture = cc[y];

                                    string captureS = capture.ToString();

                                    if (!string.Equals(captureS, regexKeyMatch))
                                    {
                                        // if not first instance print last doc
                                        if (page > 1)
                                        {
                                            ExtractPages(inputFilename, outputDirname + newDocName, docPageStart, (page - 1));
                                        }

                                        // reset the count
                                        regexKeyMatch = captureS;
                                        if (datedFileNames)
                                        {
                                            newDocName = DateTime.Now.ToString("yyyyMMdd") + "_" + captureS + ".pdf";
                                        }
                                        else
                                        {
                                            newDocName = captureS + ".pdf";
                                        }
                                        
                                        docPageStart = page;

                                        if (verbose) System.Console.WriteLine("New document at page:\t" + docPageStart);
                                    }

                                    if (showRegexMatches) System.Console.WriteLine("Capture " + y + " = '" + capture + "', Position=" + capture.Index);
                                }
                            }
                            match = match.NextMatch();
                        }
                    }

                    break;
                default:
                    if (verbose) Console.WriteLine("No valid split type selected; Exiting with error code 4.");
                    #if DEBUG
                    Console.ReadKey();
                    #endif
                    return 4;
            }


            #if DEBUG
            Console.ReadKey();
            #endif

            return 0;
        }

        // FUNCTIONS //

        private static void ExtractPages(string sourcePDFpath, string outputPDFpath, int startpage, int endpage)
        {

            PdfReader reader = new PdfReader(sourcePDFpath);
            Document sourceDocument = new Document(reader.GetPageSizeWithRotation(startpage));
            PdfCopy pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPDFpath, System.IO.FileMode.Create));

            sourceDocument.Open();

            for (int i = startpage; i <= endpage; i++)
            {
                PdfImportedPage importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                pdfCopyProvider.AddPage(importedPage);
            }
            sourceDocument.Close();
            reader.Close();
        }

    }
}
