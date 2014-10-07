using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.DocumentServices.ServiceModel.DataContracts.Xpf.Designer;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Utils;
using DevExpress.XtraPrinting.Native;
using ExcelReportGenerator;

namespace ReportGenerator.WinForms
{
    public partial class Form1 : XtraForm
    {
        public Form1()
        {
            InitializeComponent();

            monthModels = new Collection<MonthModel>();
        }
        
        private static Collection<MonthModel> monthModels { get; set; } 

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

        private static int j;

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            //layoutControlItem5.Visibility = LayoutVisibility.Always;

            var allTasks = checkedListBoxControl1.Items.Count;
            j = 0;

            var tasks = Enumerable.Range(0, checkedListBoxControl1.Items.Count).Select(it => new Task(() =>
            {
                try
                {
                    var g = new ExcelReader((string)checkedListBoxControl1.Items[it].Value);

                    var h = g.GetMonthModel();

                    monthModels.Add(h);

                    j++;

                    if (j == allTasks)
                    {
                        XtraMessageBox.Show("Done!");

                    }

                }
                catch (Exception ew)
                {
                    XtraMessageBox.Show("Error" + ew);
                }
            }));


            Parallel.ForEach(tasks, it => it.Start());


        }
    }
}
