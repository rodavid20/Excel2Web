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

namespace Excel2Web
{
    public partial class FrmTemplateCreation : Form
    {
        ChromeDriver driver;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range xlRange;
        DataTable dtFromExcel;

        public FrmTemplateCreation()
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
                        lblXmlName.Text = openXMLFile.FileName;
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
                        MessageBox.Show("XML Loaded");
                    }
                }
                catch (Exception ex)
                {
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

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtStartingRow.Text))
            {
                // Create Chrome instance using Selenium
                // Make sure the HttpWatch extension is enabled in the Selenium Chrome session by referencing the CRX file
                // e.g. C:\Program Files (x86)\HttpWatch\HttpWatchForChrome.crx
                // The HttpWatchCRXFile property returns the installed location of the CRX file
                ///var options = new ChromeOptions();
                //options.AddExtension(control.Chrome.HttpWatchCRXFile);
                // Start the Chrome browser session
                driver = new ChromeDriver();
                //// Goto blank start page so that HttpWatch recording can be started
                //driver.Navigate().GoToUrl("about:blank");
                //// Set a unique title on the first tab so that HttpWatch can attach to it
                //var uniqueTitle = Guid.NewGuid().ToString();
                //driver.ExecuteScript("document.title = '" + uniqueTitle + "'");

                //// Attach HttpWatch to the instance of Chrome created through Selenium
                //Plugin plugin = control.AttachByTitle(uniqueTitle);
                driver.Navigate().GoToUrl(txtLoginPage.Text);

            }
            else
            {
                MessageBox.Show("Please enter the submit button id");
            }
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openXMLFile = new OpenFileDialog
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
                try
                {
                    if (openXMLFile.ShowDialog() == DialogResult.OK)
                    {
                        lblExcelName.Text = openXMLFile.FileName;
                        xlApp = new Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(lblExcelName.Text, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Excel could not be loaded - " + ex.Message);
                }
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
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
                            int.TryParse(row.Cells[0].Value.ToString(), out int currentColumn);
                            object value = ((Excel.Range)xlWorkSheet.Cells[currentRow, currentColumn]).Value;
                            //string value = (string)(xlRange.Cells[currentRow, currentColumn] as Microsoft.Office.Interop.Excel.Range).Value2;
                            driver.FindElement(By.Id(row.Cells[1].Value.ToString())).SendKeys(value.ToString()); ;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Unable to find the textbox " + row.Cells[1].Value);
                        }
                    }
                }
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
