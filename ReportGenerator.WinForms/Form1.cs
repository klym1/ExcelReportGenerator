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
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraPrinting.Native;

namespace ReportGenerator.WinForms
{
    public partial class Form1 : XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void buttonBrowseFiles_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = true;
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = @"Excel files|*.xls;*.xlsx";

            openFileDialog1.ShowDialog();

            openFileDialog1.FileNames.ForEach(it => checkedListBoxControl1.Items.Add(it, CheckState.Checked));

        }

        private void simpleButtonSelectUnselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                checkedListBoxControl1.SetItemChecked(i, !checkedListBoxControl1.GetItemChecked(i));
            }
        }
    }
}
