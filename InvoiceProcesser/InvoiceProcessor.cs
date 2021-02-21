using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;

namespace InvoiceProcessor
{
    public partial class mainWindow : Form
    {
        // declaration  of variables
        DateTime startDate;
        bool validSummaryFile = false;
        bool fileSelected = false;
        string spreadsheetFilePath = null;
        FinancialYearStartPicker datePicker = new FinancialYearStartPicker();
        
        
        // inialize main window
        public mainWindow()
        {
            InitializeComponent();
            MaximizeBox = false;
            fileUpdatedLabel.Visible = false; // hide label that indicates excel file updated/processed
            progressBar.Visible = false;
        }

        // launches new summary file creation process
        private void createFileButton_Click(object sender, EventArgs e)
        {

            datePicker.ShowDialog();        // launches financial year calender form

            SaveFileDialog saveFileDialog = new SaveFileDialog();       // creates save dialog instance
            saveFileDialog.Filter = "Excel|*.xlsx";     // filters save dialog to only see excel files

            // launches save dialog once ok is clicked
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                validSummaryFile = true; 
                startDate = datePicker.getDate();   // gets date from calender

                spreadsheetFilePath = saveFileDialog.FileName;   // gets the file path of the newly created file 
                string safeFileName = Path.GetFileName(saveFileDialog.FileName); // get the name of the file created
                fileSelectedLabel.Text = safeFileName;      // sets label to the filename
                fileSelected = true;

                InvoiceProcessorHelper.CreateNewSummaryFile(saveFileDialog.FileName);              
            }

        }

        // launches the select existing file process
        private void selectFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();   // create new open file dialog
            openFileDialog.Filter = "Excel|*.xlsx";  // only show excel files
            openFileDialog.RestoreDirectory = true; // allow restore directory

            // launches  open file dialog
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileSelectedLabel.Text = openFileDialog.SafeFileName;   //  gets file path
                spreadsheetFilePath = openFileDialog.FileName;      // get file name
                fileSelected = true;
            }
        }

        // sets drag enter behaviour
        private void invoiceDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        // initiates drag and drop process
        private async void invoiceDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            fileUpdatedLabel.Visible = false;
            progressBar.Value = 0;
            progressBar.Visible = false;
            invoiceDragAreaLabel.Text = "Processing, please wait...";   // change status of drag and drop box 

            // check if a file was selected (whether new of existing)
            if (fileSelected)  
            {
                //  create a new excel instance 
                excel.Application excelApp = new excel.Application() { Visible = false };
                excel.Workbook workbook = excelApp.Workbooks.Open(spreadsheetFilePath);
                string sheetName = "Invoice Summary";
                

                // check if the correct sheet name exists
                if (InvoiceProcessorHelper.CheckWorkSheetExists(workbook, sheetName))
                {                 
                    excel.Worksheet workSheet = workbook.Worksheets[sheetName] as excel.Worksheet; // open worksheet
                    excel.Range workArea = workSheet.UsedRange;     // get used area of the sheet at thsi point in time
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);     // save all the file name of the files dragged to a string array
                    DateTime endDate;

                   // checks if an exisitng summary file was selected
                    if (!validSummaryFile)
                    {
                        try
                        {
                            startDate = InvoiceProcessorHelper.GetStartDate(workSheet);    //  tries to get the date value from the selected spreadsheet
                            validSummaryFile = true;
                        }
                        catch (NullReferenceException) //    handles the case where the date cell is blank
                        {
                            MessageBox.Show("The \"Financial Year\" field is blank in the selected Summary file. Please ensure that you have selected the correct file",
                                            "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            invoiceDragAreaLabel.Text = "Drag invoices here..";
                        }
                        catch (FormatException) //   handles date in incorrect format
                        {
                            MessageBox.Show("The \"Financial Year\" field has an invalid date, please ensure you have selected the correct file",
                                            "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            invoiceDragAreaLabel.Text = "Drag invoices here..";
                        }
                        catch (Exception ex) //  handles any error that wasnt covered previously
                        {
                            MessageBox.Show(ex.Message, "Invalid Financial Year", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            invoiceDragAreaLabel.Text = "Drag invoices here..";
                        }

                    }

                    if (validSummaryFile)
                    {
                        endDate = InvoiceProcessorHelper.GetEndDate(startDate);

                        progressBar.Visible = true;
                        Progress<ProgressReport> progress = new Progress<ProgressReport>(); // create a progress object to report status of invoice processing
                        progress.ProgressChanged += ReportProgress; // link event to handler

                        List<Invoice> validInvoices = await InvoiceProcessorHelper.ProcessInvoices(files, startDate, endDate, workArea, progress); // process invoices

                        excel.Range startingRangeCell = workSheet.Range[workSheet.Cells[9, 3], workSheet.Cells[9, 3]]; // starting cell range

                        // invoice titles to be used
                        string[,] invoiceTitles = {{"Invoice Number"}, 
                                              { "Location"},
                                              { "Invoice Date"},
                                              { "Main Revenue"},
                                              { "Other Services"},
                                              { "Other Services2"},
                                              { "GST Collected"},
                                              { "PST Collected"},
                                              { "Product Purchases"},
                                              { "Admin Fees"},
                                              { "GST of Product Purchases"},
                                              { "GST on Admin Fees"},
                                              { "Equipment Rental"},
                                              { "Benefits"},
                                              { "Net Paid"}};


                        // subheading titles to be used
                        string[,] invoiceSubtitlesTitles = { { "Invoice Date Range"},
                                                          { "Total Main Revenue"},
                                                          { "Total Other Services"},
                                                          { "Total Other Services2" },
                                                          { "Total PST Collected"},
                                                          { "PST Commission"},
                                                          { "Difference"}};

                        // arrange summary sheet grid
                        OutputProcessor.ArrangeGrid(excelApp, validInvoices, startDate, workSheet, startingRangeCell, invoiceTitles, invoiceSubtitlesTitles, 8); 

                        workbook.Save();  // save workbook
                        excelApp.Quit(); // end excel

                        // set file updates label to visible
                        fileUpdatedLabel.Visible = true;
                        progressBar.Value = 100;
                        invoiceDragAreaLabel.Text = "Drag invoices here.."; // change drag box label back after processing invoices
                    }
                   
                }
                else
                {
                    // invalid worksheet error 
                    MessageBox.Show("Invoice Summary worksheet does not exist.", "Invalid Worksheet", MessageBoxButtons.OK, MessageBoxIcon.Error);             
                    excelApp.Quit();
                }
            }
            else
            {
                // no file selected error
                MessageBox.Show("No Summary file selected. Please select an existing Summary spreadsheet or create a new one.", "No File Selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        // progress report event handler
        private void ReportProgress(object sender, ProgressReport e)
        {
            progressBar.Value = e.PercentComplete; // updates progress bar 
        }

        // opens summary file after program is closed
        private void mainWindow_FormClosing_1(object sender, FormClosingEventArgs e)
        {

            // check if file path is valid
            if (spreadsheetFilePath != null)
            {
                // reopens excel file as visible
                excel.Application excelReopenApp = new excel.Application();
                excelReopenApp.Visible = true;
                excel.Workbook workbookReopen = excelReopenApp.Workbooks.Open(spreadsheetFilePath);

            }
        }
        
      
    }
}

