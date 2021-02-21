using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;


namespace InvoiceProcessor
{
    /// <summary>
    /// Class responsible for parsing data from invoices
    /// </summary>
    static class InputProcessor
    {
        /// <summary>
        /// Parsing data from word document invoices
        /// </summary>
        /// <param name="wordInvoiceFilePath">path the word document</param>
        /// <param name="wordApp">word application used</param>
        /// <returns>Invoice object</returns>
        public static Invoice CreateInvoice(string wordInvoiceFilePath, word.Application wordApp)
        {
            word.Document wordDoc = wordApp.Documents.Open(wordInvoiceFilePath); // open document
            word.Tables tables = wordDoc.Tables; // get table references

            Invoice invoice = new Invoice(); // create invoice objects

            // populate invoice properties via setters
            invoice.InvoiceDate = DateTime.Parse(Trim(tables[1].Cell(1, 3).Range.Text));       
            invoice.InvoiceNumber = Trim(tables[1].Cell(2, 3).Range.Text);

            invoice.Company = Trim(tables[2].Cell(1, 2).Range.Text);


            string fullName = Trim(tables[2].Cell(2, 2).Range.Text);
            string[] names = fullName.Split(' ');
            invoice.FirstName = names[0];
            invoice.LastName = names[names.Length - 1];

            string timePeriod = Trim(tables[2].Cell(3, 2).Range.Text);
            string[] tempPeriod = timePeriod.Split(' ');
            string[] periodDays = tempPeriod[1].Split('-');
            int startingDay = Convert.ToInt32(periodDays[0]);
            invoice.InvoicePeriodStart = new DateTime(invoice.InvoiceDate.Year, invoice.InvoiceDate.Month, startingDay);



            string tempLocation = Trim(tables[2].Cell(4, 2).Range.Text);
            invoice.Location = tempLocation.Substring(tempLocation.IndexOf("- ") + 1);

            invoice.MainRevenue = Convert.ToDouble(Trim(tables[2].Cell(6, 3).Range.Text));
            invoice.OtherServices = Convert.ToDouble(Trim(tables[2].Cell(7, 3).Range.Text));
            invoice.OtherServices2 = Convert.ToDouble(Trim(tables[2].Cell(8, 3).Range.Text));

            string gst = Trim(tables[2].Cell(9, 3).Range.Text);
            invoice.GSTCollected = ParseNegatives(gst);

            string pst = Trim(tables[2].Cell(10, 3).Range.Text);
            invoice.PSTCollected = ParseNegatives(pst);
         
            invoice.ProductPurchases = Convert.ToDouble(Trim(tables[2].Cell(11, 3).Range.Text));
            invoice.AdminFees = Convert.ToDouble(Trim(tables[2].Cell(12, 3).Range.Text));
            invoice.ProductPurchasesGST = Convert.ToDouble(Trim(tables[2].Cell(13, 3).Range.Text));
            invoice.AdminGST = Convert.ToDouble(Trim(tables[2].Cell(14, 3).Range.Text));
            invoice.EquipmentRental = Convert.ToDouble(Trim(tables[2].Cell(14, 3).Range.Text));

            wordDoc.Close(); 

            return invoice;
        }

        /// <summary>
        /// Trims unwanted characters from strings from tables
        /// </summary>
        /// <param name="word">string to trim</param>
        /// <returns>returns trimmed string</returns>
        public static string Trim(string word)
        {
            return word.Substring(0, word.Length - 2);
        }

        /// <summary>
        /// Checks if string contains brackets and parses them accordingly
        /// </summary>
        /// <param name="tempCurrencyData">string to processes</param>
        /// <returns>double value</returns>
        public static double ParseNegatives(string tempCurrencyData)
        {
            double currencyData = 0; 

            if (tempCurrencyData[0] == '(')
            {
                currencyData = Convert.ToDouble(tempCurrencyData.Replace("(", string.Empty).Replace(")", string.Empty)) * -1;
            }
            else
            {
                currencyData = Convert.ToDouble(tempCurrencyData);
            }

            return currencyData;
        }

    }

    
}
