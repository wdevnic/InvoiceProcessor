using System;
using excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace InvoiceProcessor
{
    /// <summary>
    /// Helper class for main class
    /// </summary>
    public static class InvoiceProcessorHelper
    {
        /// <summary>
        /// Creates new excel spreadsheet
        /// </summary>
        /// <param name="fileName">spreadsheet file name</param>
        public static void CreateNewSummaryFile(string fileName)
        {
            const string worksheetName = "Invoice Summary";
            const string maxAreaRange = "A1:AA150";
            const int fontSize = 8;

            // declare excel variables
            excel.Application application = new excel.Application();
            excel.Workbook workBook = null;
            excel.Worksheet workSheet = null;

            application.Visible = false; // hide the excel GUI 

            // create new excel worksheet
            workBook = application.Workbooks.Add(excel.XlWBATemplate.xlWBATWorksheet); // workbook template
            workSheet = workBook.Worksheets[1];
            workSheet.Name = worksheetName;     // rename worksheet 

            excel.Range workArea = workSheet.Range[maxAreaRange];     // gets estimated range of the max used area
            workArea.Font.Size = fontSize; // set font size of area

            workBook.SaveAs(fileName);  // save work book
            workBook.Close();

            application.Quit();     // terminate appication
        }

        /// <summary>
        /// Checks for financial year start date in exisiting summary file
        /// </summary>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>start date of the summary sheet</returns>
        public static DateTime CheckSummaryFileStartDate(excel.Worksheet workSheet)
        {
            DateTime startDate = new DateTime();

            try
            {
                startDate = GetStartDate(workSheet);    //  tries to get the date value from the selected spreadsheet               
            }
            catch (NullReferenceException) //    handles the case where the date cell is blank
            {
                MessageBox.Show("The \"Financial Year\" field is blank in the selected Summary file. Please ensure that you have selected the correct file",
                                "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (FormatException) //   handles date in incorrect format
            {
                MessageBox.Show("The \"Financial Year\" field has an invalid date, please ensure you have selected the correct file",
                                "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex) //  handles any error that wasnt covered previously
            {
                MessageBox.Show(ex.Message, "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            
            return startDate;
        }

        /// <summary>
        /// Gets start date of existing summary file
        /// </summary>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>start date datetime object</returns>
        public static DateTime GetStartDate(excel.Worksheet workSheet)
        {
            // declare variables
            string yearStart;
            DateTime startDate;
            string financialYearRange;
            excel.Range usedRange = workSheet.UsedRange;

            // finds the date title field
            financialYearRange = usedRange.Find("Financial").Next.Value;

            // extracts date from the string
            yearStart = financialYearRange.Substring(0, financialYearRange.IndexOf(":"));
            startDate = Convert.ToDateTime(yearStart);

            return startDate;
        }

        /// <summary>
        /// Check if worksheet exists
        /// </summary>
        /// <param name="workbook">workbook used</param>
        /// <param name="worksheetName">worksheet name to check for</param>
        /// <returns>returns bool</returns>
        public static bool CheckWorkSheetExists(excel.Workbook workbook, string worksheetName)
        {
            // iterate through  sheets
            for (int i = 1; i <= workbook.Sheets.Count; i++)
            {
                excel.Worksheet sheet = workbook.Sheets[i];

                // checks if the sheet name exists
                if (sheet.Name == worksheetName)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Parses invoices from word to invoice objects
        /// </summary>
        /// <param name="files">file paths to word documents</param>
        /// <param name="startDate">financial year start date</param>
        /// <param name="endDate">financial year end date</param>
        /// <param name="workArea">Summary sheet area</param>
        /// <param name="progress">Progress object</param>
        /// <returns></returns>
        public static async Task<List<Invoice>> ProcessInvoices(string[] files, DateTime startDate, DateTime endDate, excel.Range workArea, IProgress<ProgressReport> progress)
        {
            List<Invoice> validInvoices = new List<Invoice>();
            Invoice currentInvoice;
            ProgressReport report = new ProgressReport();
            int totalFiles = files.Length;

            word.Application wordApp = new word.Application() { Visible = false };

            await Task.Run(() => // run asyn task
            { 

            Parallel.ForEach(files, fileName => // run invoice parsing to object in parrallel
            {
                try
                {
                    currentInvoice = InputProcessor.CreateInvoice(fileName, wordApp); // attempt to create invoice object from file
                                                                                      //   Console.WriteLine(currentInvoice.ToString());

                    // checks if the invoice date falls within the financial year
                    if (currentInvoice.InvoiceDate >= startDate && currentInvoice.InvoiceDate <= endDate)
                    {
                        // checks if the invoice is a duplicate
                        if ((workArea.Find(currentInvoice.InvoiceNumber) == null) && (!validInvoices.Contains(currentInvoice)))
                        {
                            validInvoices.Add(currentInvoice); // adds invoice to the arraylist

                            report.PercentComplete = validInvoices.Count * 100 / totalFiles;
                            Console.WriteLine(report.PercentComplete);
                            progress.Report(report);
                        }
                        else
                        {

                            totalFiles--; // updates total files if invalid invoice found, to progress bar will complete
                            // duplicate invoice error
                            MessageBox.Show("Invoice " + currentInvoice.InvoiceNumber + " has already been entered and will be skipped.", "Duplicate Invoice"
                                            , MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        totalFiles--; // updates total files if invalid invoice found, to progress bar will complete
                        // invoice not within financial year date error
                        MessageBox.Show("Invoice No. " + currentInvoice.InvoiceNumber + " is not part of the " + string.Format("{0:d-MMM-yy}", startDate) + "-" + string.Format("{0:d-MMM-yy}", endDate) + " financial year and will be skipped."
                                        , "Invoice Date Invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    } // string.Format("{0:d-MMM-yy}", startDate) + ":" + string.Format("{0:d-MMM-yy}", endOfFinancialYear                
                }
                catch (Exception ex)
                {
                    totalFiles--; // updates total files if invalid invoice found, to progress bar will complete
                    // invoice catch all exceptions
                    MessageBox.Show("Invoice " + fileName + " is not a valid invoice. Error Message: " + ex.Message,
                                     "Invalid Invoice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                
            });

            });

            wordApp.Quit();
            return validInvoices;
        }

        /// <summary>
        /// Gets end date
        /// </summary>
        /// <param name="startDate">start date datetime</param>
        /// <returns>last day of fincial year</returns>
        public static DateTime GetEndDate(DateTime startDate)
        {
            DateTime endDate;
            return endDate = startDate.AddYears(1).AddDays(-1);
        }
    }
}
