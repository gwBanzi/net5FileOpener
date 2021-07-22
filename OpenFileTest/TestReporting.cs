using System;
using System.Xml;
using System.Collections.Generic;
using System.IO;


namespace OpenFileTest
{
    public class TestReporting
    {
        internal enum Result
        {
            eSuccess,
            eFail
        };

        private struct TestResult
        {
            public string testDescription;
            public string filePath;
            public Result outcome;
        }

        private string xmlFilePath = null;

        private Dictionary<FileTypes, List<TestResult>> results = null;

        public OpenFilesTestOptions options { get; set; }

        public TestReporting() { }

        public TestReporting(string xmlReportPath)
        {
            xmlFilePath = xmlReportPath;
            results = new Dictionary<FileTypes, List<TestResult>>();
        }

        private void WriteXmlReport()
        {
            var xmlWriter = XmlWriter.Create(xmlFilePath);
            int idValue = 0;
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("TestCases");

            foreach (var res in results)
            {
                idValue++;
                xmlWriter.WriteStartElement("TestCase");

                xmlWriter.WriteAttributeString("Description", "Programatically open a set of files and report whether or not they open correctly.");
                xmlWriter.WriteAttributeString("ID", idValue.ToString());
                xmlWriter.WriteAttributeString("Name", String.Format("{0} - File open test", getFileExtension(res.Key).ToUpper()));

                Result overallResult = Result.eSuccess;
                /* Determine overall test result */
                foreach (var result in res.Value)
                {
                    if (result.outcome == Result.eFail)
                        overallResult = Result.eFail;
                }

                /* Set the overall (test case) result*/
                if (overallResult == Result.eFail)
                    xmlWriter.WriteAttributeString("Result", "Fail");
                else
                    xmlWriter.WriteAttributeString("Result", "Success");

                /* Write result for each test */
                foreach (var result in res.Value)
                {
                    xmlWriter.WriteStartElement("Test");

                    xmlWriter.WriteAttributeString("Name", String.Format("Opening file {0}.", result.filePath));
                    xmlWriter.WriteAttributeString("Description", result.testDescription);
                    xmlWriter.WriteAttributeString("Result", result.outcome == Result.eSuccess ? "Success" : "Fail");

                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteEndElement();
            }
            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndDocument();
            xmlWriter.Close();
        }

        public static string getFileExtension(FileTypes filetype)
        {
            string fileExtension = "None";

            switch (filetype)
            {
                case FileTypes.WordBinary:
                    fileExtension = "doc";
                    break;

                case FileTypes.ExcelBinary:
                    fileExtension = "xls";
                    break;

                case FileTypes.PowerPointBinary:
                    fileExtension = "ppt";
                    break;

                case FileTypes.WordXml:
                    fileExtension = "docx";
                    break;

                case FileTypes.ExcelXml:
                    fileExtension = "xlsx";
                    break;

                case FileTypes.PowerPointXml:
                    fileExtension = "pptx";
                    break;

                case FileTypes.PDF:
                    fileExtension = "pdf";
                    break;

                default:
                case FileTypes.Unassigned:
                    fileExtension = "Uknown";
                    break;
            }

            return fileExtension;
        }

        public static string getFileTypeName(FileTypes filetype)
        {
            switch (filetype)
            {
                case FileTypes.WordBinary:
                case FileTypes.WordXml:
                    return "Document";

                case FileTypes.ExcelBinary:
                case FileTypes.ExcelXml:

                    return "Workbook";

                case FileTypes.PowerPointBinary:
                case FileTypes.PowerPointXml:

                    return "Presentation";

                case FileTypes.PDF:

                    return "PDF";

                default:
                case FileTypes.Unassigned:
                    return "Unsupported";
            }

        }

        internal void startReport()
        {
            Console.WriteLine("Running {0} {1}", OpenFilesTestMain.GetThisApplicationName(), OpenFilesTestMain.GetThisApplicationVersion());
            Console.WriteLine("Searching {0} for files of type {1}", options.directory, "*." + options.filetype);
        }

        internal void endReport()
        {
            if (results != null)
                WriteXmlReport();
        }

        private void AddResult(string filename, string testDesc, FileTypes filetype, Result testOutcome)
        {
            if (results != null)
            {
                TestResult result = new TestResult()
                {
                    filePath = filename,
                    outcome = testOutcome,
                    testDescription = testDesc
                };

                List<TestResult> values;

                if (results.TryGetValue(filetype, out values))

                    values.Add(result);
                else
                    results.Add(filetype, new List<TestResult>() { result });
            }
        }

        internal void fileOpenedProtected(string filename, string testDesc, FileTypes filetype, Result testOutcome)
        {
            if (results != null)
                AddResult(filename, testDesc, filetype, testOutcome);

            Console.WriteLine("{0}: {1} - Opened Protected", getFileTypeName(filetype), filename);
        }

        internal void fileOpenedNormally(string filename, string testDesc, FileTypes filetype, Result testOutcome, ref int passCount, ref StreamWriter log)
        {
            if (results != null)
            {
                AddResult(filename, testDesc, filetype, testOutcome);
            }
            passCount++;
            Console.WriteLine("{0}: {1} - Opened OK", getFileTypeName(filetype), filename);
            if (log is object)
            {
                log.WriteLine("{0}: {1} - Opened OK", getFileTypeName(filetype), filename);
            }
        }

        internal void fileFailedToOpen(string filename, string reason, FileTypes filetype, Result testOutcome, ref StreamWriter log)
        {
            if (results != null)
                AddResult(filename, String.Format("File Open - Failed with '{0}'", reason), filetype, testOutcome);

            Console.WriteLine("{0}: {1} - Error: {2}", getFileTypeName(filetype), filename, reason);
            if (log is object)
            {
                log.WriteLine("{0}: {1} - Error: {2}", getFileTypeName(filetype), filename, reason);
            }
        }
    }
}
