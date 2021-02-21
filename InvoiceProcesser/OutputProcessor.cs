using System;
using System.Collections.Generic;
using excel = Microsoft.Office.Interop.Excel;


namespace InvoiceProcessor
{

    /// <summary>
    /// Class that handles writing the Invoices to Excel
    /// </summary>
    static class OutputProcessor
    {
        const string accountingFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* \"" + "-\"" + "??_);_(@_)";
        const int numberOfQuarters = 4;

        /// <summary>
        /// Writes the invoice data to the spreadsheet
        /// </summary>
        /// <param name="invoice">Invoice object</param>
        /// <param name="columnToUpdate">Range object</param>
        public static void WriteInvoiceToSpreadSheet(Invoice invoice, excel.Range columnToUpdate, excel.Worksheet workSheet)
        {

            // string which sets the cells in the range to the accounting format 
            columnToUpdate.NumberFormat = accountingFormat;

            // write invoice data to the range
            columnToUpdate[1, 1] = invoice.InvoiceNumber;
            columnToUpdate[2, 1] = invoice.Location;
            columnToUpdate[3, 1] = string.Format("{0:dd-MMM-yy}", invoice.InvoiceDate); // formats the date 
            columnToUpdate[4, 1] = invoice.MainRevenue * -1;
            columnToUpdate[5, 1] = invoice.OtherServices * -1;
            columnToUpdate[6, 1] = invoice.OtherServices2 * -1;
            columnToUpdate[7, 1] = invoice.GSTCollected * -1;
            columnToUpdate[8, 1] = invoice.PSTCollected * -1;
            columnToUpdate[9, 1] = invoice.ProductPurchases; ;
            columnToUpdate[10, 1] = invoice.AdminFees; ;
            columnToUpdate[11, 1] = invoice.ProductPurchasesGST;
            columnToUpdate[12, 1] = invoice.AdminGST;
            columnToUpdate[13, 1] = invoice.EquipmentRental;
            columnToUpdate[14, 1] = invoice.Benefits;

            CalculateInvoiceTotal(columnToUpdate, workSheet);

            workSheet.Columns.AutoFit();   // autofit the columns

        }

     
        /// <summary>
        /// Writes header data
        /// </summary>
        /// <param name="invoice">Invoice object</param>
        /// <param name="startDate">Financial year range</param>
        /// <param name="workSheet">Worksheet object</param>
        public static void WriteHeaderData(Invoice invoice, DateTime startDate, excel.Worksheet workSheet)
        {

            const string headerTitleRange = "A1:A3";
            const string headerDataRange = "B1:B2";

            // 2D array header titles
            string[,] header = { { "Company" },
                                { "Financial Year" },
                                { "Summary of ASLA Invoices" } };

            // fixed title range
            excel.Range headerRange = workSheet.Range[headerTitleRange];
            headerRange.Value = header; // sets header titles to range
            headerRange.Columns.AutoFit();
            headerRange.Font.Bold = true;

            // fixed header data value
            excel.Range headerValues = workSheet.Range[headerDataRange];

            DateTime endOfFinancialYear = startDate.AddMonths(12).AddDays(-1); // get financial year range

            headerValues[1, 1] = invoice.Company; // write company to range
            headerValues[2, 1] = string.Format("{0:d-MMM-yy}", startDate) + ":" + string.Format("{0:d-MMM-yy}", endOfFinancialYear);   // writes financial year period        

        }


        /// <summary>
        /// Write totals data for each quarter
        /// </summary>
        /// <param name="usedInvoiceRange">Range used by invoices for this quarter</param>
        /// <param name="rowOffset">Offset is the number of additional rows required for the totals</param>
        /// <param name="workSheet">Worksheet object</param>
        public static void WriteTitleData(excel.Range subheadingDataRange, excel.Range usedInvoiceRange, excel.Worksheet workSheet)
        {

            excel.Range invoiceDateRange = subheadingDataRange.Columns[1].Rows[1];   //usedInvoiceRange.Columns[1].Rows[fullRowCount -5]; //  gets the range contains invoice dates    
            excel.Range totalMainRange = subheadingDataRange.Columns[1].Rows[2];  // get the range that contains revenue values
            excel.Range otherServiceRange = subheadingDataRange.Columns[1].Rows[3];  // get the range that contains revenue values
            excel.Range otherServiceRange2 = subheadingDataRange.Columns[1].Rows[4];
            excel.Range totalPSTRange = subheadingDataRange.Columns[1].Rows[5];  // gets the range that contains PST values      
            excel.Range pstCommission = subheadingDataRange.Columns[1].Rows[6]; // gets the range that contains commission values
            excel.Range differenceRange = subheadingDataRange.Columns[1].Rows[7];   // gets difference       

            // sets cell format to accounting
            pstCommission.NumberFormat = accountingFormat;
            differenceRange.NumberFormat = accountingFormat;

            // gets the date in the first column and last column
            DateTime invoiceStartDate = Convert.ToDateTime(usedInvoiceRange.Columns[1].Rows[3].Value);
            DateTime invoiceEndDate = Convert.ToDateTime(usedInvoiceRange.Columns[usedInvoiceRange.Columns.Count].Rows[3].Value);

            // format date
            string dateRange = string.Format("{0:dd-MMM-yy}", invoiceStartDate) + ":" + string.Format("{0:dd-MMM-yy}", invoiceEndDate);
            invoiceDateRange.Value = dateRange;

            // calculate the totals
            totalMainRange.Formula = "=SUM(" + usedInvoiceRange.Rows[4].Address[false, false] + ")";
            otherServiceRange.Formula = "=SUM(" + usedInvoiceRange.Rows[5].Address[false, false] + ")";
            otherServiceRange2.Formula = "=SUM(" + usedInvoiceRange.Rows[6].Address[false, false] + ")";
            totalPSTRange.Formula = "=SUM(" + usedInvoiceRange.Rows[8].Address[false, false] + ")";
            differenceRange.Formula = "=" + -1 + "*" + totalPSTRange.Address[false, false];
        }


        /// <summary>
        /// Setups up Summary Sheet
        /// </summary>
        /// <param name="excelApp">Excel Application</param>
        /// <param name="invoices">list of invocies</param>
        /// <param name="startFinancialYear">start date</param>
        /// <param name="workSheet">worsheet used</param>
        /// <param name="startingCell">Top left corner cell address of sheet working area</param>
        /// <param name="invoiceTitles">2D array of invoice field titles></param>
        /// <param name="invoiceSubtitlesTitles">2D array of subheading titles</param>
        /// <param name="interval">number of empty rows between quarters</param>
        public static void ArrangeGrid(excel.Application excelApp, List<Invoice> invoices, DateTime startFinancialYear, excel.Worksheet workSheet, excel.Range startingCell, string[,] invoiceTitles, string[,] invoiceSubtitlesTitles, int interval)
        {
            int invoiceHeadingsCount = invoiceTitles.GetLength(0);
            int subheadingCount = invoiceSubtitlesTitles.GetLength(0);

           
            excel.Range[] startingInvoiceColumns = FirstInvoiceRanges(startingCell, workSheet, invoiceHeadingsCount, subheadingCount, interval); //create an array of the first column range for each quarter
            excel.Range[] subHeadingRanges = SubHeadingRanges(startingCell, workSheet, invoiceHeadingsCount, subheadingCount, interval);// create an array of ranges for subheadings

            excel.Range firstColumn = startingInvoiceColumns[0];
            excel.Range secondColumn = startingInvoiceColumns[1];
            excel.Range thirdColumn = startingInvoiceColumns[2];
            excel.Range fourthColumn = startingInvoiceColumns[3];

            //  checks if no invoices have been written to the spreadsheet
            if (GetCurrentlyUsedRange(excelApp, firstColumn, workSheet) == null &&
                GetCurrentlyUsedRange(excelApp, secondColumn, workSheet) == null &&
                GetCurrentlyUsedRange(excelApp, thirdColumn, workSheet) == null &&
                GetCurrentlyUsedRange(excelApp, fourthColumn, workSheet) == null)
            {
                WriteHeaderData(invoices[0], startFinancialYear, workSheet); // writes header data data
            }

            // gets the next empty invoice range for each quarter
            excel.Range firstQuarterNextEmptyColumn = InitialNextEmptyRange(excelApp, firstColumn, workSheet);
            excel.Range secondQuarterNextEmptyColumn = InitialNextEmptyRange(excelApp, secondColumn, workSheet);
            excel.Range thirdQuarterNextEmptyColumn = InitialNextEmptyRange(excelApp, thirdColumn, workSheet);
            excel.Range fourthQuarterNextEmptyColumn = InitialNextEmptyRange(excelApp, fourthColumn, workSheet);

            // iterate through all invoices in arraylist
            foreach (Invoice invoice in invoices)
            {
              
                // check if invoice date falls within first quarter
                if ((invoice.InvoiceDate >= startFinancialYear) && (invoice.InvoiceDate <= startFinancialYear.AddMonths(3)))
                {
                    firstQuarterNextEmptyColumn = AddInvoice(firstQuarterNextEmptyColumn, workSheet, invoice);
                }

                // check if invoice date falls within second quarter
                if ((invoice.InvoiceDate > startFinancialYear.AddMonths(3)) && (invoice.InvoiceDate <= startFinancialYear.AddMonths(6)))
                {
                    secondQuarterNextEmptyColumn = AddInvoice(secondQuarterNextEmptyColumn, workSheet, invoice);
                }

                // check if invoice date falls within third quarter
                if ((invoice.InvoiceDate > startFinancialYear.AddMonths(6)) && (invoice.InvoiceDate <= startFinancialYear.AddMonths(9)))
                {
                    thirdQuarterNextEmptyColumn = AddInvoice(thirdQuarterNextEmptyColumn, workSheet, invoice);
                }

                // check if invoice date falls within fourth quarter
                if ((invoice.InvoiceDate > startFinancialYear.AddMonths(9)) && (invoice.InvoiceDate <= startFinancialYear.AddMonths(12)))
                {
                    fourthQuarterNextEmptyColumn = AddInvoice(fourthQuarterNextEmptyColumn, workSheet, invoice);
                }
            }

            excel.Range[] currentUsedInvoiceRange = GetUsedInvoiceRanges(excelApp, startingInvoiceColumns, workSheet); // gets the currently used invoice ranges for each quarter

            ProcessTitles(excelApp, workSheet, startingInvoiceColumns, invoiceTitles); // checks where invoice titles are needed and writes them to spreadsheet
            SortUsedRanges(currentUsedInvoiceRange); // sorts invoices by date and location

            ProcessSubHeadings(subHeadingRanges, currentUsedInvoiceRange, workSheet); // writes in subheading data where needed
            ProcessTitles(excelApp, workSheet, subHeadingRanges, invoiceSubtitlesTitles); // writes in subheading titles where needed
        }


        /// <summary>
        /// Writes subheading data for each quarter
        /// </summary>
        /// <param name="subheadingDataRange">Subheading ranges</param>
        /// <param name="usedInvoiceRange">Invocie ranges used</param>
        /// <param name="workSheet">worksheet used</param>
        public static void ProcessSubHeadings(excel.Range[] subheadingDataRange, excel.Range[] usedInvoiceRange, excel.Worksheet workSheet)
        {
            for (int i = 0; i < numberOfQuarters; i++)
            {
                if (usedInvoiceRange[i] != null) // checks if there are no invoices for a quarter
                {
                    WriteTitleData(subheadingDataRange[i], usedInvoiceRange[i], workSheet); // writes title data
                }

            }
        }


        /// <summary>
        /// Gets used invoice ranges for each quarter
        /// </summary>
        /// <param name="excelApp">excell application</param>
        /// <param name="firstColumnRanges">starting invoice range for each quarter</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>Array of used invoice ranges </returns>
        public static excel.Range[] GetUsedInvoiceRanges(excel.Application excelApp, excel.Range[] firstColumnRanges, excel.Worksheet workSheet)
        {
            excel.Range[] usedInvoiceRanges = new excel.Range[4];

            for (int i = 0; i < numberOfQuarters; i++)
            {
                usedInvoiceRanges[i] = GetCurrentlyUsedRange(excelApp, firstColumnRanges[i], workSheet); // used invoice range for each quarter
            }

            return usedInvoiceRanges;
        }


        /// <summary>
        /// Sorts invoices in each quarter by date and office
        /// </summary>
        /// <param name="currentlyUsedRange"></param>
        public static void SortUsedRanges(excel.Range[] currentlyUsedRange)
        {
            for (int i = 0; i < numberOfQuarters; i++)
            {
                OutputProcessorHelper.SortInvoices(currentlyUsedRange[i]);             
            }

        }


        /// <summary>
        /// Writes titles for invoice fields where needed
        /// </summary>
        /// <param name="excelApp">exel application</param>
        /// <param name="workSheet">worksheet used</param>
        /// <param name="startingColumns">first invoice ranges for each quarter</param>
        /// <param name="invoiceTitles">array of required titles</param>
        public static void ProcessTitles(excel.Application excelApp, excel.Worksheet workSheet, excel.Range[] startingColumns, string[,] invoiceTitles)
        {
            for (int i = 0; i < numberOfQuarters; i++)
            {
                excel.Range currentTitleRange = (OutputProcessorHelper.GetTitleRange(startingColumns[i], workSheet)); // gets title range

                if (OutputProcessorHelper.TitlesNeeded(excelApp, currentTitleRange, workSheet)) // checks if titles are needed (if there are invocies for this quarter)
                {
                    OutputProcessorHelper.WriteInvoiceTitles(currentTitleRange, invoiceTitles, workSheet); // writes the titles
                }
            }
        }


        /// <summary>
        /// Adds invoice to worksheet
        /// </summary>
        /// <param name="nextEmptyColumn">range to add invoice</param>
        /// <param name="workSheet">worksheet to use</param>
        /// <param name="invoice">invocie to write</param>
        /// <returns>next empty invoice range</returns>
        public static excel.Range AddInvoice(excel.Range nextEmptyColumn, excel.Worksheet workSheet, Invoice invoice)
        {
            WriteInvoiceToSpreadSheet(invoice, nextEmptyColumn, workSheet);    // write invoice to range
            nextEmptyColumn = OutputProcessorHelper.GetNextEmptyRange(nextEmptyColumn, workSheet); // get next empty invoice range

            return nextEmptyColumn; 
        }

        
        /// <summary>
        /// Gets current used range for a quarter
        /// </summary>
        /// <param name="excelApp">excel application used</param>
        /// <param name="startingRange">first invoice range in the column</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>used range</returns>
        public static excel.Range GetCurrentlyUsedRange(excel.Application excelApp, excel.Range startingRange, excel.Worksheet workSheet)
        {
            // calculate row number and column number for first and last cells in the range
            excel.Range currentColumnRange = startingRange;
            int firstCellRow = startingRange.Row;
            int firstCellColumn = startingRange.Column;
            int secondCellRow = startingRange.Row + (startingRange.Rows.Count - 1);
            int secondCellColumn = startingRange.Column + (startingRange.Columns.Count - 1);

            if (excelApp.WorksheetFunction.CountA(startingRange) == 0) // check if first invoice range is empty
            {
                currentColumnRange = null;
            }
            else
            {
                while (excelApp.WorksheetFunction.CountA(currentColumnRange) != 0) // iterate through all invoice ranges for a quarter
                {
                    currentColumnRange = workSheet.Range[workSheet.Cells[firstCellRow, ++firstCellColumn], workSheet.Cells[secondCellRow, ++secondCellColumn]];
                }

                currentColumnRange = workSheet.Range[workSheet.Cells[firstCellRow, startingRange.Column], workSheet.Cells[secondCellRow, secondCellColumn - 1]]; // get used range
            }

            return currentColumnRange;
        }

       
        /// <summary>
        /// Get used range for each quarter on an existing summary sheet
        /// </summary>
        /// <param name="excelApp">excel app used</param>
        /// <param name="startingColumnRange">first invocie range</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>next empty invocie range</returns>
        public static excel.Range InitialNextEmptyRange(excel.Application excelApp, excel.Range startingColumnRange, excel.Worksheet workSheet)
        {
            excel.Range rangeUsed = GetCurrentlyUsedRange(excelApp, startingColumnRange, workSheet);

            if (rangeUsed == null)
            {
                return startingColumnRange;
            }
            else
            {
                return OutputProcessorHelper.GetNextEmptyRange(rangeUsed, workSheet);
            }
        }


        /// <summary>
        /// Calculates invoice totals
        /// </summary>
        /// <param name="startingRange">range of the invoice</param>
        /// <param name="workSheet"></param>
        public static void CalculateInvoiceTotal(excel.Range startingRange, excel.Worksheet workSheet)
        {

            excel.Range currentColumnRange = startingRange;

            // get range numbers
            int firstCellRow = startingRange.Row;
            int firstCellColumn = startingRange.Column;
            int secondCellRow = startingRange.Row + (startingRange.Rows.Count - 1);
            int secondCellColumn = startingRange.Column + (startingRange.Columns.Count - 1);

            currentColumnRange = workSheet.Range[workSheet.Cells[firstCellRow + 3, firstCellColumn], workSheet.Cells[secondCellRow - 1, secondCellColumn]];
            excel.Range totalRange = workSheet.Range[workSheet.Cells[secondCellRow, secondCellColumn], workSheet.Cells[secondCellRow, secondCellColumn]]; // get total range
            totalRange.Formula = "=SUM(" + currentColumnRange.Address[false, false] + ")"; // assign totals formula to range

        }
        

        /// <summary>
        /// Gets an array of the first invoice ranges for each quarter
        /// </summary>
        /// <param name="startingCell">cell from which all other ranges are calculated</param>
        /// <param name="workSheet">worksheet used </param>
        /// <param name="amountOfRows">number of invoice headings</param>
        /// <param name="amountofSubheading">number of subheading fields</param>
        /// <param name="freeSpace">number of rows between quarters</param>
        /// <returns>Array of first invoice ranges for each quarter</returns>
        public static excel.Range[] FirstInvoiceRanges(excel.Range startingCell, excel.Worksheet workSheet, int amountOfRows, int amountofSubheading, int freeSpace)
        {
            excel.Range[] invoiceStartingRanges = new excel.Range[4];

            for (int i = 0; i < numberOfQuarters; i++)
            {
                invoiceStartingRanges[i] = OutputProcessorHelper.GetFirstInvoiceColumn(startingCell, workSheet, amountOfRows, amountofSubheading, freeSpace, i); // get range and add to array
            }

            return invoiceStartingRanges;
        }


        /// <summary>
        /// Gets an array of ranges used for subheadings
        /// </summary>
        /// <param name="startingCell">cell from which all other ranges are calculated</param>
        /// <param name="workSheet">worksheet used </param>
        /// <param name="amountOfRows">number of invoice headings</param>
        /// <param name="amountofSubheading">number of subheading fields</param>
        /// <param name="freeSpace">number of rows between quarters</param>
        /// <returns>Array of subheading ranges for each quarter</returns>/param>
        public static excel.Range[] SubHeadingRanges(excel.Range startingCell, excel.Worksheet workSheet, int amountOfRows, int amountofSubheading, int freeSpace)
        {
            excel.Range[] subHeadingRanges = new excel.Range[4];

            for (int i = 0; i < numberOfQuarters; i++)
            {
                subHeadingRanges[i] = OutputProcessorHelper.GetSubheadingColumn(startingCell, workSheet, amountOfRows, amountofSubheading, freeSpace, i);
            }

            return subHeadingRanges;

        }
    }
}
