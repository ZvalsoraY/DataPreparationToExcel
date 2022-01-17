using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataPreparationToExcelNS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog() 
            {
                Multiselect = true,
                Title = "Select files",
            };
            //Filter by output files
            dlg.Filter = "Output files|*.output";
            dlg.ShowDialog();

            if (dlg.FileName == String.Empty)
                return;
            string[] files_names = dlg.FileNames;
            listBox1.Items.Clear();
            listBox1.Items.AddRange(files_names);
            ConverterToExcel.list.Clear();
            ConverterToExcel.list.AddRange(files_names);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ConverterToExcel.createExcelForListFiles(ConverterToExcel.list);
            MessageBox.Show("Done");
        }
    }
}
