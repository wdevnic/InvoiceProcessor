using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceProcessor
{
    public partial class FinancialYearStartPicker : Form
    {
        // constructor
        public FinancialYearStartPicker()
        {           
            InitializeComponent();
            MaximizeBox = false;
        }

        // get date from calender
        public DateTime getDate()
        {
            return dateTimePicker.SelectionStart.Date ;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
