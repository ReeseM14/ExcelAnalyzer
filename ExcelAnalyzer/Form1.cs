using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAnalyzer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Filter = "Microsoft Excel Worksheet | *.xlsx";

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show(fbd.FileName);
            }
        }
    }
}
