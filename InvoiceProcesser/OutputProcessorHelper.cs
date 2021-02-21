using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;

namespace InvoiceProcessor
{
    /// <summary>
    /// Helper class for Output processor class
    /// </summary>
    public class OutputProcessorHelper
    {
        const int numberOfQuarters = 4;

        /// <summary>
        /// Sort a range of cells by 2 rows
        /// </summary>
        /// <param name="invoiceRange">range of invoices</param>
        public static void SortInvoices(excel.Range invoiceRange)
        {
            if (invoiceRange != null)
            {
                excel.Range firstRowFourthQuarter = invoiceRange.Rows[2]; // get range of locations
                excel.Range secondRowFourthQuarter = invoiceRange.Rows[3]; // get range of invoice dates

                // sort by location then invoice date
                invoiceRange.Sort(secondRowFourthQuarter, excel.XlSortOrder.xlAscending, firstRowFourthQuarter, Type.Missing, excel.XlSortOrder.xlAscending);
            }
        }


        /// <summary>
        /// Checks if invoice titles are required 
        /// </summary>
        /// <param name="excelApp">excel application</param>
        /// <param name="startingRange">first invoice range</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>bool</returns>
        public static bool TitlesNeeded(excel.Application excelApp, excel.Range startingRange, excel.Worksheet workSheet)
        {
            excel.Range currentColumnRange = startingRange;
            int firstCellRow = startingRange.Row;
            int firstCellColumn = startingRange.Column;
            int secondCellRow = startingRange.Row + (startingRange.Rows.Count - 1);
            int secondCellColumn = startingRange.Column + (startingRange.Columns.Count - 1);

            // checks if title range is blank while invoices are present for that quarter
            if ((excelApp.WorksheetFunction.CountA(startingRange) == 0) && (excelApp.WorksheetFunction.CountA(workSheet.Range[workSheet.Cells[firstCellRow, firstCellColumn + 1], workSheet.Cells[secondCellRow, secondCellColumn + 1]]) != 0))
            {
                return true;
            }

            return false;
        }


        /// <summary>
        /// Gets the title range for a quarter
        /// </summary>
        /// <param name="startingRange">starting invoice range</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns></returns>
        public static excel.Range GetTitleRange(excel.Range startingRange, excel.Worksheet workSheet)
        {
            excel.Range currentColumnRange = startingRange;
            int firstCellRow = startingRange.Row;
            int firstCellColumn = startingRange.Column;
            int secondCellRow = startingRange.Row + (startingRange.Rows.Count - 1);
            int secondCellColumn = startingRange.Column + (startingRange.Columns.Count - 1);

            // calculates title range for quarter
            currentColumnRange = workSheet.Range[workSheet.Cells[firstCellRow, secondCellColumn - 1], workSheet.Cells[secondCellRow, secondCellColumn - 1]];

            return currentColumnRange;
        }


        /// <summary>
        /// Writes titles for invoices
        /// </summary>
        /// <param name="titlesRange">title range to write titles</param>
        /// <param name="invoiceTitles">array of titles</param>
        /// <param name="workSheet">worksheet used</param>
        public static void WriteInvoiceTitles(excel.Range titlesRange, string[,] invoiceTitles, excel.Worksheet workSheet)
        {
            titlesRange.Value = invoiceTitles; // set the titles to range
            titlesRange.Columns.AutoFit(); // autofit titles
            titlesRange.Font.Bold = true; // bold all titles
        }


        /// <summary>
        /// Gets first invoice range for a quarter
        /// </summary>
        /// <param name="startingCell">range of cell from which all other ranges are calculated</param>
        /// <param name="workSheet">worksheet used</param>
        /// <param name="amountOfRows">amount of invoice headings</param>
        /// <param name="amountofSubheading">amount of subheadings</param>
        /// <param name="freeSpace">amount of rows between quarters</param>
        /// <param name="quarterNumber">int representing quarter</param>
        /// <returns>range of first invoice</returns>   
        public static excel.Range GetFirstInvoiceColumn(excel.Range startingCell, excel.Worksheet workSheet, int amountOfRows, int amountofSubheading, int freeSpace, int quarterNumber)
        {
            excel.Range quarterRange = null;

            int firstRow = startingCell.Row + ((amountOfRows + amountofSubheading + freeSpace) * quarterNumber);

            // calculates range of first invoice for specific quarter
            quarterRange = workSheet.Range[workSheet.Cells[firstRow, startingCell.Column + 1], workSheet.Cells[(firstRow + amountOfRows) - 1, startingCell.Column + 1]];

            return quarterRange;
        }


        /// <summary>
        /// Gets subheading range for a quarter
        /// </summary>
        /// <param name="startingCell">range of cell from which all other ranges are calculated</param>
        /// <param name="workSheet">worksheet used</param>
        /// <param name="amountOfRows">amount of invoice headings</param>
        /// <param name="amountofSubheading">amount of subheadings</param>
        /// <param name="freeSpace">amount of rows between quarters</param>
        /// <param name="quarterNumber">int representing quarter</param>
        /// <returns>returns range of subheading</returns>   
        public static excel.Range GetSubheadingColumn(excel.Range startingCell, excel.Worksheet workSheet, int amountOfRows, int amountofSubheading, int freeSpace, int quarterNumber)
        {
            excel.Range quarterRange = null;

            int firstRow = startingCell.Row + amountOfRows + 1 + ((amountOfRows + amountofSubheading + freeSpace) * quarterNumber);

            quarterRange = workSheet.Range[workSheet.Cells[firstRow, startingCell.Column + 1], workSheet.Cells[(firstRow + amountofSubheading) - 1, startingCell.Column + 1]];

            return quarterRange;
        }


        /// <summary>
        /// Gets range to write next empty invoice 
        /// </summary>
        /// <param name="startingRange">first invoice</param>
        /// <param name="workSheet">worksheet used</param>
        /// <returns>range of next empty invoice range</returns>
        public static excel.Range GetNextEmptyRange(excel.Range startingRange, excel.Worksheet workSheet)
        {
            excel.Range currentColumnRange = startingRange;
            int firstCellRow = startingRange.Row;
            int firstCellColumn = startingRange.Column;
            int secondCellRow = startingRange.Row + (startingRange.Rows.Count - 1);
            int secondCellColumn = startingRange.Column + (startingRange.Columns.Count - 1);

            // calculates next empty range
            currentColumnRange = workSheet.Range[workSheet.Cells[firstCellRow, secondCellColumn + 1], workSheet.Cells[secondCellRow, secondCellColumn + 1]];

            return currentColumnRange;
        }
    }
}
