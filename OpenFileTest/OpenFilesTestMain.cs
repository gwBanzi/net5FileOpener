using System;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace OpenFileTest
{

    public enum FileTypes
    {
        Unassigned,
        WordBinary,
        ExcelBinary,
        PowerPointBinary,
        WordXml,
        ExcelXml,
        PowerPointXml,
        PDF
    };
    public enum ReturnStatus
    {
        FileOpened = 0,
        DirOpened = 0,
        FileCouldNotOpen = 1,
        DirContainsFilesCouldNotOpen = 1,
        Error,
        BadArgs,
        UndefinedStatus
    };

    class OpenFilesTestMain
    {
        static int Main(string[] args)
        {
            try
            {
                ReturnStatus status = ReturnStatus.UndefinedStatus;
                OpenFilesTestOptions opt = new OpenFilesTestOptions();
                StreamWriter loggingObject = null;
                TestReporting testReport = null;
                string outputDirectory = "";

                if (CommandLine.Parser.Default.ParseArguments(args, opt))
                {

                    if (String.IsNullOrEmpty(opt.xmlreport))
                    {
                        testReport = new TestReporting();
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(opt.outputDirectory))
                        {
                            outputDirectory = @"output";
                        }
                        else
                        {
                            outputDirectory = opt.outputDirectory;
                        }

                        string xmlreport = outputDirectory + "//" + opt.xmlreport;
                        testReport = new TestReporting(xmlreport);
                    }

                    testReport.options = opt;

                    if (ValidateArgs(opt) == false)
                        return (int)ReturnStatus.BadArgs;

                    if (opt.log)
                    {
                        loggingObject = InitializeLogging(outputDirectory);
                    }

                    testReport.startReport();

                    if (opt.directory != null)
                    {
                        status = OpenDirectory(opt, ref testReport, ref loggingObject);
                    }

                    if (opt.file != null)
                    {
                        int passCount = 0; //Redundant Variable now
                        FileInfo fileInfo = new FileInfo(opt.file);
                        FileTypes fileTypeSelected = SelectFileType(fileInfo.Extension.ToLower());
                        status = OpenFile(fileTypeSelected, fileInfo, ref testReport, ref passCount, ref loggingObject);
                    }

                    if (testReport != null)
                    {
                        testReport.endReport();
                    }
                    if (loggingObject != null)
                    {
                        loggingObject.Close();
                    }

                    return (int)status;
                }
                else
                {
                    Console.WriteLine("Couldn't Parse Args");
                    return (int)ReturnStatus.BadArgs;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("General exception occurred: {0}", ex.Message);
            }
            return (int)ReturnStatus.UndefinedStatus;
        }

        public static bool ValidateArgs(OpenFilesTestOptions opt)
        {
            string[] validFileTypes = {"doc", "xls", "ppt",
                "docx", "xslx", "pptx", "docm", "xlsm", "pptm",
                "pdf","*" };

            if (!validFileTypes.Any(opt.filetype.ToLower().Contains))
            {
                Console.Error.WriteLine("Filetype must be one of doc|xls|ppt|docx|xlsx|pptx|docm|xlsm|pptm|pdf");
                return false;
            }

            if (opt.directory != null && opt.file != null)
            {
                Console.Error.WriteLine("Cannot Use Both File and Directory Options. Please select 1.");
                return false;
            }

            return true;
        }

        public static ReturnStatus OpenDirectory(OpenFilesTestOptions opt, ref TestReporting testReport, ref StreamWriter loggingObject)
        {
            int fileCount = 0;
            int passCount = 0;

            DirectoryInfo sourceDirectory = new DirectoryInfo(opt.directory);

            if (!sourceDirectory.Exists)
            {
                Console.Error.WriteLine("Directory '{0}' does not exist.", sourceDirectory.FullName);
                return ReturnStatus.Error;
            }

            foreach (FileInfo fileinfo in sourceDirectory.EnumerateFiles("*." + opt.filetype, SearchOption.AllDirectories))
            {
                fileCount++;
                FileTypes fileTypeSelected = SelectFileType(fileinfo.Extension.ToLower());
                OpenFile(fileTypeSelected, fileinfo, ref testReport, ref passCount, ref loggingObject);
            }

            Console.WriteLine("Files opened successfully {0} of {1}", passCount, fileCount);

            if (loggingObject is object)
            {
                loggingObject.WriteLine("Files opened successfully {0} of {1}", passCount, fileCount);
            }

            return fileCount == passCount ? ReturnStatus.DirOpened : ReturnStatus.DirContainsFilesCouldNotOpen;
        }

        public static string GetThisApplicationVersion()
        {
            return Assembly.GetEntryAssembly().GetName().Version.ToString();
        }

        public static string GetThisApplicationName()
        {
            string name = Process.GetCurrentProcess().MainModule.FileName;
            return Path.GetFileName(name);
        }
        public static StreamWriter InitializeLogging(string outDirectory)
        {
            try
            {
                DirectoryInfo outputDirectory = new DirectoryInfo(outDirectory);
                if (!outputDirectory.Exists)
                {
                    Directory.CreateDirectory(outputDirectory.FullName);
                }
                string FILE_NAME = outputDirectory + "\\Results.txt";
                var objWriter = new StreamWriter(FILE_NAME);
                return objWriter;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Environment.Exit(-1);
            }
            return null;
        }

        public static FileTypes SelectFileType(string filestring)
        {
            FileTypes newFileType = FileTypes.Unassigned;

            filestring = filestring.TrimStart('.');

            /* binary cfb files */
            if (filestring == "doc")
                newFileType = FileTypes.WordBinary;
            if (filestring == "xls")
                newFileType = FileTypes.ExcelBinary;
            if (filestring == "ppt")
                newFileType = FileTypes.PowerPointBinary;

            /* xml opc files */
            if (filestring == "docx")
                newFileType = FileTypes.WordXml;
            if (filestring == "xlsx")
                newFileType = FileTypes.ExcelXml;
            if (filestring == "pptx")
                newFileType = FileTypes.PowerPointXml;

            /* macro enabled xml opc files */
            if (filestring == "docm")
                newFileType = FileTypes.WordXml;
            if (filestring == "xlsm")
                newFileType = FileTypes.ExcelXml;
            if (filestring == "pptm")
                newFileType = FileTypes.PowerPointXml;

            /* pdf files */
            if (filestring == "pdf")
                newFileType = FileTypes.PDF;

            return newFileType;
        }

        public static ReturnStatus OpenFile(FileTypes fileTypeSelected, FileInfo fileinfo, ref TestReporting testReport, ref int passCount, ref StreamWriter loggingObject)
        {
            bool is_opened = false;
            switch (fileTypeSelected)
            {
                case FileTypes.WordBinary:
                case FileTypes.WordXml:

                    is_opened = TestOpenOfficeDocument.TestWordFileOpensSpire(fileinfo.FullName, fileTypeSelected, ref testReport, ref passCount, ref loggingObject);
                    break;

                case FileTypes.ExcelBinary:
                case FileTypes.ExcelXml:

                    is_opened = TestOpenOfficeExcel.TestExcelFileOpensSpire(fileinfo.FullName, fileTypeSelected, ref testReport, ref passCount, ref loggingObject);
                    break;

                case FileTypes.PowerPointBinary:
                case FileTypes.PowerPointXml:

                    is_opened = TestOpenOfficePowerPoint.TestPowerpointFileOpensSpire(fileinfo.FullName, fileTypeSelected, ref testReport, ref passCount, ref loggingObject);
                    break;

                case FileTypes.PDF:

                    is_opened = TestOpenPDF.TestPDFFileOpen(fileinfo.FullName, fileTypeSelected, ref testReport, ref passCount, ref loggingObject);
                    break;

                default:
                    testReport.fileFailedToOpen(fileinfo.FullName, string.Format("Failed to Open - Unknown Filetype {0}", fileinfo.Extension), fileTypeSelected, TestReporting.Result.eFail, ref loggingObject);
                    break;
            }

            if (is_opened)
            {
                return ReturnStatus.FileOpened;
            }
            return ReturnStatus.FileCouldNotOpen;
        }

    }
}
