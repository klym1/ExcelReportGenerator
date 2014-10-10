using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting.Native;
using ExcelReportGenerator;
using ExcelReportGenerator.Models;

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

            checkedListBoxControl1.Items.Clear();
            openFileDialog1.FileNames.ForEach(it => checkedListBoxControl1.Items.Add(it, CheckState.Checked));
        }

        private void simpleButtonSelectUnselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                checkedListBoxControl1.SetItemChecked(i, !checkedListBoxControl1.GetItemChecked(i));
            }
        }

        private List<string> CheckedItems
        {
            get { return checkedListBoxControl1.Items.GetCheckedValues().Cast<string>().ToList(); }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            var allTasks = CheckedItems.Count;
            
            if(allTasks == 0) return;
                
            EnableAll(false);

            var  monthModels = new Collection<MonthModel>();
            
            BeginInvoke(new Action(() =>
            {
                for (int it = 0; it < allTasks; it++)
                {
                    try
                    {
                        var excelReader = new ExcelReader(CheckedItems[it]);

                        monthModels.Add(excelReader.GetMonthModel());
                    }
                    catch (Exception ew)
                    {
                        Invoke(new Action(() =>
                        {
                            EnableAll(true);
                            XtraMessageBox.Show(ew.Message);
                        }));

                        return;
                    }
                }

                var dataProcessor = new DataProcessor(monthModels);

                var modelsProcessor = new Generator(dataProcessor);

                modelsProcessor.GenerateExcelResult();

                Invoke(new Action(() => EnableAll(true)));
            }));
        }

        private void EnableAll(bool state)
        {
            simpleButtonGenerate.Enabled = state;
        }
    }
}
