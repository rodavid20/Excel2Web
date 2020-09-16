using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Excel2Web.Helper;

namespace Excel2Web
{
    public partial class FrmMain : Form
    {
        IWebDriver driver;
//        ChromeDriver driver;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnLoadXML_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openXMLFile = new OpenFileDialog
            {
                Title = "Open XML File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xml",
                Filter = "xml files (*.xml)|*.xml",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            })
            {
                try
                {
                    if (openXMLFile.ShowDialog() == DialogResult.OK)
                    {
                        txtXMLName.Text = openXMLFile.FileName;
                        txtXMLName.SelectionStart = txtXMLName.Text.Length;
                        txtXMLName.SelectionLength = 0;
                        XDocument doc = XDocument.Load(openXMLFile.FileName);
                        XElement root = doc.Element("Template");
                        dgvColumns.Rows.Clear();
                        foreach (XElement column in root.Element("Columns").Elements("Column"))
                        {
                            int rowId = dgvColumns.Rows.Add();
                            DataGridViewRow row = dgvColumns.Rows[rowId];
                            row.Cells[0].Value = column.Element("ExcelColumnNumber").Value;
                            row.Cells[1].Value = column.Element("HtmlName").Value;
                        }

                        XElement info = root.Element("Info");
                        txtDataEntryPage.Text = info.Element("DataEntryPage").Value;
                        txtLoginPage.Text = info.Element("LoginPage").Value;
                        txtStartingRow.Text = info.Element("StartingRow").Value;
                        txtSubmitButtonId.Text = info.Element("SubmitButtonId").Value;
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteToLog(String.Format("Message : {0}{1}Source: {2}{1}Excel: {3}{1}Stack Trace: {4}",
                                ex.Message, Environment.NewLine, ex.Source, txtXMLName.Text, ex.StackTrace),
                                System.Diagnostics.EventLogEntryType.Error);
                    MessageBox.Show("Error XML could not be loaded - " + ex.Message);
                }
            }
        }

        private void btnSaveXML_Click(object sender, EventArgs e)
        {
            XDocument template = new XDocument();
            template.Add(new XElement("Template"));
            XElement info = new XElement("Info",
                new XElement("DataEntryPage", txtDataEntryPage.Text),
                new XElement("LoginPage", txtLoginPage.Text),
                new XElement("StartingRow", txtStartingRow.Text),
                new XElement("SubmitButtonId", txtSubmitButtonId.Text));
            template.Element("Template").Add(info);

            template.Element("Template").Add(new XElement("Columns"));

            foreach (DataGridViewRow row in dgvColumns.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null &&
                    !String.IsNullOrEmpty(row.Cells[0].Value.ToString()) && !String.IsNullOrEmpty(row.Cells[1].Value.ToString()))
                {
                    XElement element = new XElement("Column",
                        new XElement("ExcelColumnNumber", row.Cells[0].Value),
                        new XElement("HtmlName", row.Cells[1].Value));
                    template.Element("Template").Element("Columns").Add(element);
                }
            }
            using (SaveFileDialog saveXMLFile = new SaveFileDialog
            {
                Title = "Save XML File",

                CheckPathExists = true,

                DefaultExt = "xml",
                Filter = "xml files (*.xml)|*.xml",
                FilterIndex = 2,
                RestoreDirectory = true
            })
            {
                if (saveXMLFile.ShowDialog() == DialogResult.OK)
                {
                    template.Save(saveXMLFile.FileName);
                    MessageBox.Show("XML Saved");
                }
            }
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            btnOpenExcel.Enabled = false;
            btnSubmit.Enabled = false;
            if (!String.IsNullOrEmpty(txtStartingRow.Text))
            {
                using (OpenFileDialog openExcelFile = new OpenFileDialog
                {
                    Title = "Open Excel File",

                    CheckFileExists = true,
                    CheckPathExists = true,

                    DefaultExt = "xlsx",
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FilterIndex = 2,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                })
                {
                    PleaseWait pleaseWait = new PleaseWait();

                    //  Allow main UI thread to properly display please wait form.
                    Application.DoEvents();
                    try
                    {
                        if (openExcelFile.ShowDialog() == DialogResult.OK)
                        {
                            // Display form modelessly
                            pleaseWait.Show();
                            txtExcelName.Text = openExcelFile.FileName;
                            xlApp = new Excel.Application();
                            xlWorkBook = xlApp.Workbooks.Open(txtExcelName.Text, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                            // Create Chrome instance using Selenium
                            // Make sure the HttpWatch extension is enabled in the Selenium Chrome session by referencing the CRX file
                            // e.g. C:\Program Files (x86)\HttpWatch\HttpWatchForChrome.crx
                            // The HttpWatchCRXFile property returns the installed location of the CRX file
                            //var options = new ChromeOptions();
                            //options.AddExtension(control.Chrome.HttpWatchCRXFile);
                            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                            service.HideCommandPromptWindow = true;
                            // Start the Chrome browser session
                            //driver = new ChromeDriver(service);
                            driver = new ChromeDriver(service);

                            //// Goto blank start page so that HttpWatch recording can be started
                            //driver.Navigate().GoToUrl("about:blank");
                            //// Set a unique title on the first tab so that HttpWatch can attach to it
                            //var uniqueTitle = Guid.NewGuid().ToString();
                            //driver.ExecuteScript("document.title = '" + uniqueTitle + "'");

                            //// Attach HttpWatch to the instance of Chrome created through Selenium
                            //Plugin plugin = control.AttachByTitle(uniqueTitle);
                            driver.Navigate().GoToUrl(txtLoginPage.Text);
                            txtCurrentRow.Text = txtStartingRow.Text;
                            txtExcelName.SelectionStart = txtExcelName.Text.Length;
                            txtExcelName.SelectionLength = 0;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteToLog(String.Format("Message : {0}{1}Source: {2}{1}Excel: {3}{1}Stack Trace: {4}",
                                ex.Message, Environment.NewLine, ex.Source, txtExcelName.Text, ex.StackTrace),
                                System.Diagnostics.EventLogEntryType.Error);

                        MessageBox.Show("Error Excel could not be loaded - " + ex.Message);

                    }
                    pleaseWait.Close();
                }
            }
            else
            {
                MessageBox.Show("Please enter the starting row value");
            }
            btnSubmit.Enabled = true;
            btnOpenExcel.Enabled = true;
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string errorText = String.Empty;
            int currentRow = 0;
            if (xlWorkSheet != null && int.TryParse(txtCurrentRow.Text, out currentRow))
            {
                foreach (DataGridViewRow row in dgvColumns.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[1].Value != null &&
                        !String.IsNullOrEmpty(row.Cells[0].Value.ToString()) && !String.IsNullOrEmpty(row.Cells[1].Value.ToString()))
                    {
                        try
                        {
                            string id = row.Cells[1].Value.ToString();
                            string excelColumn = row.Cells[0].Value.ToString();
                            object value =  string.Empty;

                            if (excelColumn.Contains(","))
                            {
                                string[] excelColumns = excelColumn.Split(',');
                                foreach(var c in excelColumns)
                                {
                                    int.TryParse(c, out int currentColumn);
                                    value +=  ", " + ((Excel.Range)xlWorkSheet.Cells[currentRow, currentColumn]).Value;
                                }
                                value = value.ToString().Substring(2);
                            }
                            else
                            {
                                int.TryParse(excelColumn, out int currentColumn);
                                value = ((Excel.Range)xlWorkSheet.Cells[currentRow, currentColumn]).Value;
                            }
                            IWebElement webElement = driver.FindElement(By.Id(row.Cells[1].Value.ToString()));
                            //string value = (string)(xlRange.Cells[currentRow, currentColumn] as Microsoft.Office.Interop.Excel.Range).Value2;
                            if(value != null)
                            {
                                string str = value.ToString().Replace("\n", " ").Replace("\t", " ");

                                js.ExecuteScript("arguments[0].value='" + str + "'", webElement);

                            }
                            //driver.FindElement(By.Id(row.Cells[1].Value.ToString())).SendKeys(value.ToString()); ;
                        }
                        catch (Exception ex)
                        {
                            errorText += String.Format("Message : {0}{1}Source: {2}{1}Field: {3}{1}Stack Trace: {4}",
                                ex.Message, Environment.NewLine, ex.Source, row.Cells[1].Value, ex.StackTrace);
                            MessageBox.Show("Unable to find the textbox " + row.Cells[1].Value + Environment.NewLine + ex.Message);
                        }
                    }
                }
                if(!string.IsNullOrEmpty(errorText))
                {
                    Logger.WriteToLog(errorText, System.Diagnostics.EventLogEntryType.Error);
                }
                txtCurrentRow.Text = (currentRow + 1).ToString();
            }
            else
            {
                MessageBox.Show("Invalid Excel or current row");
            }
        }
        //private void PopulateDataTable(string fileName)
        //{
        //    DataTable dt = new DataTable();

        //    using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
        //    {

        //        WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
        //        IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
        //        string relationshipId = sheets.First().Id.Value;
        //        WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
        //        Worksheet workSheet = worksheetPart.Worksheet;
        //        SheetData sheetData = workSheet.GetFirstChild<SheetData>();
        //        IEnumerable<Row> rows = sheetData.Descendants<Row>();

        //        foreach (Cell cell in rows.ElementAt(0))
        //        {
        //            dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
        //        }

        //        foreach (Row row in rows) //this will also include your header row...
        //        {
        //            DataRow tempRow = dt.NewRow();

        //            for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
        //            {
        //                tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i - 1));
        //            }

        //            dt.Rows.Add(tempRow);
        //        }
        //    }
        //    dt.Rows.RemoveAt(0); //...so i'm taking it out here.
        //}
        //public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        //{
        //    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
        //    string value = cell.CellValue.InnerXml;

        //    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //    {
        //        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        //    }
        //    else
        //    {
        //        return value;
        //    }
        //}
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {

            }
        }
    }
}
