using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace e2020_tool_data
{
    public class Program
    {
        public static string ResultFolder = @"C:\e2020_tool_data\";
        public static string InputFolder = @"Input\";
        public static string OutputFolder = @"Output\";
        public static string Sheet1Name = @"Extraction Sheet";
        public static string Sheet2Name = @"Video QA Sheet";
        public static string LogFileName = "System.docx";
        public static void Main(string[] args)
        {
            string folderName = ResultFolder;
            if (!Directory.Exists(ResultFolder + InputFolder))
            {
                Directory.CreateDirectory(ResultFolder + InputFolder);
            }
            folderName = @"Input";
            if (!Directory.Exists(ResultFolder + InputFolder))
            {
                Directory.CreateDirectory(ResultFolder + InputFolder);
            }
            folderName = @"Output";
            if (!Directory.Exists(ResultFolder + OutputFolder))
            {
                Directory.CreateDirectory(ResultFolder + OutputFolder);
            }

            string pathString = ResultFolder + InputFolder;
            DirectoryInfo dir = new DirectoryInfo(pathString);

            if (dir.GetFiles().Select(x => x.Name).FirstOrDefault() != null)
            {
                String[] fileName = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault(name => name != "Thumbs.db").Split('.');
                folderName = ResultFolder + OutputFolder;
                pathString = System.IO.Path.Combine(folderName, fileName[0]);
                if (fileName[0] != "" && (fileName[1] == "ppt" || fileName[1] == "pptx"))
                {
                    ProcessingFolder(dir);
                    ProcessingDocument(dir);
                    ProcessingSheet1(dir, pathString, folderName, Sheet1Name);
                    ProcessingSheet2(dir, pathString, folderName, Sheet2Name);
                }
                else
                {
                    if ( fileName[0] == null && fileName[0] == "")
                    {
                        Console.WriteLine("Opss...Error : there is no file exits in the path of " + ResultFolder + InputFolder + "!!");
                        Console.ReadKey();
                    }
                    else if ((fileName[1] != "ppt" && fileName[1] != "pptx"))
                    {
                        Console.WriteLine("Opss...Error : there is no ppt file exits in the path of " + ResultFolder + InputFolder + "!!");
                        Console.ReadKey();
                    }                    
                }
            }
            else
            {
                Console.WriteLine("Opss...Error : there is no file exits in the path of "+ResultFolder+InputFolder+"!!");
                Console.ReadKey();
            }
        }
        public static void ProcessingFolder(DirectoryInfo dir)
        {

            String[] fileName = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault(name => name != "Thumbs.db").Split('.');

            string folderName = ResultFolder + OutputFolder;

            string pathString = System.IO.Path.Combine(folderName, fileName[0]);
            Directory.CreateDirectory(pathString);

            string fileNam1 = fileName[0] + "." + fileName[1];
            string fileNam2 = fileName[0] + ".xlsx";
            string sourcePath = ResultFolder + InputFolder;
            string targetPath = pathString;

            string sourceFile = System.IO.Path.Combine(sourcePath, fileNam1);
            string destFile = System.IO.Path.Combine(targetPath, fileNam1);
            System.IO.File.Copy(sourceFile, destFile, true);

            try
            {
                Application powerPoint = new Application();
                String inputFile = ResultFolder + InputFolder + fileName[0] + "." + fileName[1];
                powerPoint.Visible = MsoTriState.msoTrue;
                powerPoint.WindowState = PpWindowState.ppWindowMinimized;
                Presentations oPresSet = powerPoint.Presentations;
                _Presentation oPres = oPresSet.Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                String changeName = "00";
                String newName = "01";
                for (int i = 0; i < oPres.Slides.Count; i++)
                {
                    var str = oPres.Slides[i + 1].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                    if (str.Contains("#framechain"))
                    {
                        if (Convert.ToInt32(changeName) < 5)
                            changeName = "0" + (Convert.ToInt32(changeName) + 5).ToString();
                        else changeName = (Convert.ToInt32(changeName) + 5).ToString();

                        string folderNam = ResultFolder + OutputFolder + fileName[0];
                        pathString = System.IO.Path.Combine(folderNam, fileName[0] + "-" + changeName);
                        System.IO.Directory.CreateDirectory(pathString);
                        newName = "01";
                    }
                    if (str.Contains("#frame"))
                    {
                        string folderNam = ResultFolder + OutputFolder + fileName[0];
                        folderNam = System.IO.Path.Combine(folderNam, fileName[0] + "-" + changeName);
                        pathString = System.IO.Path.Combine(folderNam, fileName[0] + "-" + changeName + "-" + newName);
                        System.IO.Directory.CreateDirectory(pathString);
                        if (Convert.ToInt32(newName) < 9)
                            newName = "0" + (Convert.ToInt32(newName) + 1).ToString();
                        else newName = (Convert.ToInt32(newName) + 1).ToString();
                    }

                }
                powerPoint.Quit();
            }
            catch (Exception e)
            {
                //var errors = e.Message;
                //Console.WriteLine("Opss...Error : "+errors);
                //Console.ReadKey();
                //throw;
            }

        }
        public static void ProcessingDocument(DirectoryInfo dir)
        {

            String[] fileName = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault(name => name != "Thumbs.db").Split('.');
            string folderName = ResultFolder + OutputFolder;
            string pathString = System.IO.Path.Combine(folderName, fileName[0]);
            string fileNam2 = fileName[0] + ".docx";
            string targetPath = pathString;


            try
            {
                //Create an instance for word app
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set status for word application is to be visible or not.
                winword.Visible = false;

                //Create a missing variable for missing value
                object missing = System.Reflection.Missing.Value;

                //Create a new document
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Heading 1";
                para1.Range.set_Style(ref styleHeading1);
                para1.Range.Text = "Extract Vocabulary from ppt";
                para1.Range.InsertParagraphAfter();
                try
                {
                    Application powerPoint = new Application();
                    String inputFile = ResultFolder + InputFolder + fileName[0] + "." + fileName[1];
                    powerPoint.Visible = MsoTriState.msoTrue;
                    powerPoint.WindowState = PpWindowState.ppWindowMinimized;
                    Presentations oPresSet = powerPoint.Presentations;
                    _Presentation oPres = oPresSet.Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    int rows = 0;
                    for (int i = 0; i < oPres.Slides.Count; i++)
                    {
                        var str = oPres.Slides[i + 1].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                        int ii = str.IndexOf("SOURCES:");
                        if (str != "" && ii != -1)
                        {
                            rows++;
                        }
                    }

                    Microsoft.Office.Interop.Word.Table firstTable = document.Tables.Add(para1.Range, rows, 3, ref missing, ref missing);
                    firstTable.Borders.Enable = 1;
                    String[] heading = { "#", "File Name", "Source" };
                    int k = 0;

                    String changeName = "00";
                    String newName = "01";
                    String docFileName = "";
                    int row = 0;
                    for (int i = 0; i < oPres.Slides.Count; i++)
                    {
                        var str = oPres.Slides[i + 1].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                        int ii = str.IndexOf("SOURCES:");
                        String substr = "";
                        if (str.Contains("#framechain") || str.Contains("#frame"))
                        {
                            if (ii > 0)
                            {
                                substr = str.Replace(str.Substring(0, ii) + "SOURCES:", "");

                                if (str.Contains("#framechain"))
                                {
                                    if (Convert.ToInt32(changeName) < 5)
                                        changeName = "0" + (Convert.ToInt32(changeName) + 5).ToString();
                                    else changeName = (Convert.ToInt32(changeName) + 5).ToString();

                                    newName = "01";
                                }
                                if (str.Contains("#frame"))
                                {
                                    docFileName = fileName[0] + "-" + changeName + "-" + newName;
                                    if (Convert.ToInt32(newName) < 9)
                                        newName = "0" + (Convert.ToInt32(newName) + 1).ToString();
                                    else newName = (Convert.ToInt32(newName) + 1).ToString();
                                }

                                // Microsoft.Office.Interop.Word.Table firstTable1 = document.Tables.Add(para1.Range, 1, 3, ref missing, ref missing);
                                //firstTable1.Borders.Enable = 1;
                                var r = firstTable.Rows[row + 1];
                                row++;
                                foreach (Microsoft.Office.Interop.Word.Cell cell in r.Cells)
                                {
                                    if (cell.RowIndex == 1)
                                    {
                                        //cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
                                        cell.Range.Text = heading[k++];
                                        cell.Range.Font.Bold = 1;
                                        cell.Range.Font.Name = "verdana";
                                        cell.Range.Font.Size = 10;
                                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                                        cell.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray25;
                                        //Center alignment for the Header cells
                                        cell.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                                        cell.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                    }
                                    else
                                    {
                                        if (cell.ColumnIndex == 1)
                                        {
                                            cell.Range.Text = newName;
                                        }
                                        else if (cell.ColumnIndex == 2)
                                        {
                                            cell.Range.Text = docFileName;
                                        }
                                        else if (cell.ColumnIndex == 3)
                                        {
                                            cell.Range.Text = substr;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    firstTable.AutoFormat();
                    powerPoint.Quit();
                }
                catch (Exception e)
                {
                    //var errors = e.Message;
                    //Console.WriteLine("Opss...Error : " + errors);
                    //Console.ReadKey();
                    //throw;
                }
                //Save the document
                object filename = System.IO.Path.Combine(targetPath, fileNam2);
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
            catch (Exception e)
            {
                //var errors = e.Message;
                //Console.WriteLine("Opss...Error : " + errors);
                //Console.ReadKey();
                //throw;
            }

        }
        public static void ProcessingSheet1(DirectoryInfo dir, string pathString, string folderName, string sheetName)
        {
            String s = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault();
            String[] fileName = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault(name => name != "Thumbs.db").Split('.');

            folderName = ResultFolder + OutputFolder;

            string fileNam2 = fileName[0] + ".xlsx";
            string targetPath = pathString;

            try
            {
                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.DisplayAlerts = false;
#if DEBUG
                oXL.Visible = true;
#else
                oXL.Visible = false;
#endif

                //Get a new workbook.

                Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Add(Missing.Value);
                Microsoft.Office.Interop.Excel._Worksheet oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Name = sheetName;
                for (int i = 1; i <= 6; i++)
                {
                    oSheet.Cells[i, 1].Font.Bold = true;
                    oSheet.Cells[i, 1].Font.Size = 12;
                }
                oSheet.Cells[1, 1] = "Course";
                oSheet.Cells[2, 1] = "Unit";
                oSheet.Cells[3, 1] = "Lesson Title";
                oSheet.Cells[4, 1] = "Lesson #";
                oSheet.Cells[5, 1] = "Author";
                oSheet.Cells[6, 1] = "Version #";

                for (int i = 1; i <= 6; i++)
                {
                    oSheet.Cells[8, i].Font.Bold = true;
                    oSheet.Cells[8, i].Font.Size = 12;
                }
                oSheet.Cells[8, 1] = "Activity Type";
                oSheet.Cells[8, 2] = "Activity Name";
                oSheet.Cells[8, 3] = "Frame";
                oSheet.Cells[8, 4] = "Slide Title";
                oSheet.Cells[8, 5] = "Slide Type";
                oSheet.Cells[8, 6] = "Slide Layout";
                try
                {
                    Application powerPoint = new Application();
                    String inputFile = ResultFolder + InputFolder + fileName[0] + "." + fileName[1];
                    powerPoint.Visible = MsoTriState.msoTrue;
                    powerPoint.WindowState = PpWindowState.ppWindowMinimized;
                    Presentations oPresSet = powerPoint.Presentations;
                    _Presentation oPres = oPresSet.Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    bool flag = true;
                    int row = 9;
                    String changeName = "00";
                    String newName = "01";
                    for (int i = 0; i < oPres.Slides.Count; i++)
                    {
                        var str = oPres.Slides[i + 1].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                        if (str.Contains("#framechain") || str.Contains("#frame"))
                        {
                            if (str.Contains("#framechain"))
                            {
                                if (Convert.ToInt32(changeName) < 5)
                                    changeName = "0" + (Convert.ToInt32(changeName) + 5).ToString();
                                else changeName = (Convert.ToInt32(changeName) + 5).ToString();

                                newName = "01";
                            }
                            if (str.Contains("#frame"))
                            {
                                if (Convert.ToInt32(newName) < 9)
                                    newName = "0" + (Convert.ToInt32(newName) + 1).ToString();
                                else newName = (Convert.ToInt32(newName) + 1).ToString();
                            }

                            //oSheet.Cells[row, 1] = "Activity Type";
                            //oSheet.Cells[row, 2] = "Activity Name";
                            oSheet.Cells[row, 3] = changeName;
                            //oSheet.Cells[row, 4] = oPres.TitleMaster;
                            //oSheet.Cells[row, 5] = oPres.TitleMaster;
                            //oSheet.Cells[row, 6] = oPres.LayoutDirection;
                            row++;
                        }

                        foreach (var item in oPres.Slides[i + 1].Shapes)
                        {
                            var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                            var type = shape.Type;
                            var tye = shape.Title;
                            var ty = shape.Name;
                            var t = shape.OLEFormat;
                            var rr = type;

                            if (shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    var textRange = shape.TextFrame.TextRange;

                                    var text = textRange.Text;
                                    if (text.Contains("Course:") && flag)
                                    {
                                        String[] lines = text.Split('\r');
                                        foreach (String ss in lines)
                                        {
                                            if (ss.Contains("Course:"))
                                            {
                                                String line = ss.Replace("Course:", " ").Trim();
                                                oSheet.Cells[1, 2] = line;
                                            }
                                            else if (ss.Contains("Unit:"))
                                            {
                                                String line = ss.Replace("Unit:", " ").Trim();
                                                oSheet.Cells[2, 2] = line;
                                            }
                                            else if (ss.Contains("Lesson Title:"))
                                            {
                                                String line = ss.Replace("Lesson Title:", " ").Trim();
                                                oSheet.Cells[3, 2] = line;
                                            }
                                            else if (ss.Contains("Lesson #:"))
                                            {
                                                String line = ss.Replace("Lesson #:", " ").Trim();
                                                oSheet.Cells[4, 2] = line;
                                            }
                                            else if (ss.Contains("Author:"))
                                            {
                                                String line = ss.Replace("Author:", " ").Trim();
                                                oSheet.Cells[5, 2] = line;
                                            }
                                            else if (ss.Contains("Version #:"))
                                            {
                                                String line = ss.Replace("Version #:", " ").Trim();
                                                oSheet.Cells[6, 2] = line;
                                            }
                                        }
                                    }

                                }
                            }
                        }

                    }
                    powerPoint.Quit();
                }
                catch (Exception e)
                {
                    //var errors = e.Message;
                    //Console.WriteLine("Opss...Error : " + errors);
                    //Console.ReadKey();
                    //throw;
                }
                String ReportFile = System.IO.Path.Combine(targetPath, fileNam2);
                oWB.SaveAs(ReportFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                                        Type.Missing, Type.Missing,
                                        false,
                                        false,
                                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing);

                oXL.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                oSheet = null;
                oWB = null;
                oXL = null;
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
            catch (Exception e)
            {
                //var errors = e.Message;
                //Console.WriteLine("Opss...Error : " + errors);
                //Console.ReadKey();
                //throw;
            }
        }
        public static void ProcessingSheet2(DirectoryInfo dir, string pathString, string folderName, string sheetName)
        {
            String[] fileName;
            String s = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault();
            fileName = dir.GetFiles().Select(fi => fi.Name).FirstOrDefault(name => name != "Thumbs.db").Split('.');

            folderName = ResultFolder + OutputFolder;
            string fileNam2 = fileName[0] + ".xlsx";
            string targetPath = pathString;

            try
            {
                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.DisplayAlerts = false;
#if DEBUG
                oXL.Visible = true;
#else
                oXL.Visible = false;
#endif

                String inputFil = System.IO.Path.Combine(targetPath, fileNam2);
                Microsoft.Office.Interop.Excel.Workbook oWB = oXL.Workbooks.Open(inputFil);

                Microsoft.Office.Interop.Excel._Worksheet oSheet = oWB.Sheets[2];
                oSheet.Name = sheetName;
                for (int i = 1; i <= 6; i++)
                {
                    oSheet.Cells[i, 1].Font.Bold = true;
                    oSheet.Cells[i, 1].Font.Size = 12;
                }
                oSheet.Cells[1, 1] = "Unit Name ";
                oSheet.Cells[2, 1] = "Lesson Name ";
                oSheet.Cells[3, 1] = "Filmer Name ";
                oSheet.Cells[4, 1] = "QAer Name ";
                oSheet.Cells[5, 1] = "Space for teacher picture insertion";
                oSheet.Cells[6, 1] = "Review Criteria rules ";

                for (int i = 1; i <= 13; i++)
                {
                    oSheet.Cells[8, i].Font.Bold = true;
                    oSheet.Cells[8, i].Font.Size = 12;
                }
                oSheet.Cells[8, 1] = "Frame Chain ";
                oSheet.Cells[8, 2] = "Frame ";
                oSheet.Cells[8, 3] = "Slide # ";
                oSheet.Cells[8, 4] = "Filename ";
                oSheet.Cells[8, 5] = "Type ";
                oSheet.Cells[8, 6] = "Date Filmed";
                oSheet.Cells[8, 7] = "Length";
                oSheet.Cells[8, 8] = "Status";
                oSheet.Cells[8, 9] = "1st Review Comments";
                oSheet.Cells[8, 10] = "Reshoot Status";
                oSheet.Cells[8, 11] = "2nd Review Comments";
                oSheet.Cells[8, 12] = "2nd Reshoot Status";
                oSheet.Cells[8, 13] = "3rd Review Comments";
                try
                {
                    Application powerPoint = new Application();
                    String inputFile = ResultFolder + InputFolder + fileName[0] + "." + fileName[1];
                    powerPoint.Visible = MsoTriState.msoTrue;
                    powerPoint.WindowState = PpWindowState.ppWindowMinimized;
                    Presentations oPresSet = powerPoint.Presentations;
                    _Presentation oPres = oPresSet.Open(inputFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);


                    bool flag = true;
                    int row = 9;
                    String changeName = "00";
                    String newName = "00";

                    for (int i = 0; i < oPres.Slides.Count; i++)
                    {
                        var str = oPres.Slides[i + 1].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                        if (str.Contains("#framechain") || str.Contains("#frame"))
                        {
                            if (str.Contains("#framechain"))
                            {
                                if (Convert.ToInt32(changeName) < 5)
                                    changeName = "0" + (Convert.ToInt32(changeName) + 5).ToString();
                                else changeName = (Convert.ToInt32(changeName) + 5).ToString();
                                newName = "00";
                            }
                            if (str.Contains("#frame"))
                            {
                                if (Convert.ToInt32(newName) < 9)
                                    newName = "0" + (Convert.ToInt32(newName) + 1).ToString();
                                else newName = (Convert.ToInt32(newName) + 1).ToString();
                            }

                            string filename = "";
                            string type = "";
                            if (str.Contains("Anchor video"))
                            {
                                filename = fileName[0] + "-" + changeName + "-" + newName + "-anchor";
                                type = "Video";
                            }
                            else if (str.Contains("Instructional video"))
                            {
                                filename = fileName[0] + "-" + changeName + "-" + newName + "-instructional";
                                type = "Video";
                            }
                            else if (str.Contains("Entry audio") || str.Contains("Exit audio") || str.Contains("Hint Audio"))
                            {
                                if (str.Contains("Entry audio"))
                                {
                                    filename = fileName[0] + "-" + changeName + "-" + newName + "-entry";
                                }
                                if (str.Contains("Hint Audio"))
                                {
                                    filename += "\n" + fileName[0] + "-" + changeName + "-" + newName + "-hint";
                                }
                                if (str.Contains("Exit audio"))
                                {
                                    filename += "\n" + fileName[0] + "-" + changeName + "-" + newName + "-exit";
                                }
                                type = "Audio";
                            }

                            oSheet.Cells[row, 1] = changeName;
                            oSheet.Cells[row, 2] = newName;
                            oSheet.Cells[row, 3] = i + 1;
                            oSheet.Cells[row, 4] = filename;
                            oSheet.Cells[row, 5] = type;
                            //oSheet.Cells[row, 6] = "Date Filmed";
                            //oSheet.Cells[row, 7] = "Length";
                            //oSheet.Cells[row, 8] = "Status";
                            //oSheet.Cells[row, 9] = "1st Review Comments";
                            //oSheet.Cells[row, 10] = "Reshoot Status";
                            //oSheet.Cells[row, 11] = "2nd Review Comments";
                            //oSheet.Cells[row, 12] = "2nd Reshoot Status";
                            //oSheet.Cells[row, 13] = "3rd Review Comments";
                            row++;
                        }
                        foreach (var item in oPres.Slides[i + 1].Shapes)
                        {
                            var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;

                            if (shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    var textRange = shape.TextFrame.TextRange;

                                    var text = textRange.Text;
                                    if (text.Contains("Course:") && flag)
                                    {
                                        String[] lines = text.Split('\r');
                                        foreach (String ss in lines)
                                        {
                                            if (ss.Contains("Unit:"))
                                            {
                                                String line = ss.Replace("Unit:", " ").Trim();
                                                oSheet.Cells[1, 2] = line;
                                            }
                                            if (ss.Contains("Lesson Title:"))
                                            {
                                                String line = ss.Replace("Lesson Title:", " ").Trim();
                                                oSheet.Cells[2, 2] = line;
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                    powerPoint.Quit();
                }
                catch (Exception e)
                {
                    //var errors = e.Message;
                    //Console.WriteLine("Opss...Error : " + errors);
                    //Console.ReadKey();
                    //throw;
                }
                String ReportFile = System.IO.Path.Combine(targetPath, fileNam2);
                oWB.SaveAs(ReportFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                                        Type.Missing, Type.Missing,
                                        false,
                                        false,
                                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing,
                                        Type.Missing);

                oXL.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                oSheet = null;
                oWB = null;
                oXL = null;
                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
            catch (Exception e)
            {
                //var errors = e.Message;
                //Console.WriteLine("Opss...Error : " + errors);
                //Console.ReadKey();
                //throw;
            }
        }
    }
}
