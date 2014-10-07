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

            int totalIterations = 30;

            monthModels = new Collection<MonthModel>();

            var g = new ExcelReader(@"C:\Users\mykola.klymyuk\Desktop\3rd Quater 2014\Aug " + 19 + @" 2014.xls");

           var h = g.GetMonthModel();
            //var j = h.Records.Where(it => it.cst_nm < 0).OrderBy(it => it.cst_nam).ToList();

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
            layoutControlItem5.Visibility = LayoutVisibility.Always;

            var allTasks = 25;
            j = 0;

            var tasks = Enumerable.Range(0, checkedListBoxControl1.Items.Count).Select(it => new Task(() =>
            {
                try
                {
                    var g = new ExcelReader((string)checkedListBoxControl1.Items[it].Value);

                    var h = g.GetMonthModel();

                    monthModels.Add(h);
                    j++;

                    if (j > allTasks)
                    {
                        var h2 = monthModels;
                    }

                  //  progressBarControl1.Position += 4;
                  //  progressBarControl1.Refresh();

                    Debug.WriteLine(it + " --- " + h.Records.Count);
                }
                catch (Exception )
                {
                    Debug.WriteLine(it + " Exception");
                }
            }));


            Parallel.ForEach(tasks, it => it.Start());


        }
    }
}
