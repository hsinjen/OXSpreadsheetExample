using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OXSpreadsheetExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn新增列_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "xlsx|*.xlsx";
            if (of.ShowDialog() == DialogResult.OK)
            {
                OXSpreadsheet spreadsheet = new OXSpreadsheet();
                if (spreadsheet.Open(of.FileName, true) == false)
                {
                    return;
                }
                spreadsheet.InsertRow("工作表1", 2);
                spreadsheet.Close();
            }
        }
    }
}
