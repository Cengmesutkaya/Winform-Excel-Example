using ExcelDataReader;
using MetroFramework.Forms;
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
using Winform_Excel_Import.models;

namespace Winform_Excel_Import
{
    public partial class frmMain : MetroForm
    { 
        DataTableCollection dataTableCollection;
        public frmMain()
        {
            InitializeComponent();
        }
        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            LoadExcel();
        }
        public void FillGrid()
        {
            DataTable table = dataTableCollection[cboSheet.SelectedItem.ToString()];
            dataGrid.DataSource = table;
            List<Customer> customers = new List<Customer>(); // Api tarafından gelen model
            /// Veri veri tabanına eklenmek istenirse bu kod bloğu kullanılabilir.
            //for (int i = 0; i < dataGrid.Rows.Count; i++)          
            //{
            //    Customer customer = new Customer();
            //    customer.Code = table.Rows[i]["Code"].ToString(); 
            //    customer.Name = table.Rows[i]["Name"].ToString();
            //    customers.Add(customer);
            //}
        }
        public void LoadExcel()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        txtFileName.Text = openFileDialog.FileName;
                        using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                });
                                dataTableCollection = result.Tables;
                                cboSheet.Items.Clear();
                                foreach (DataTable table in dataTableCollection)
                                    cboSheet.Items.Add(table.TableName);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Yüklemek istedğiniz excel dosyası açıktır. Lütfen kapatıp öyle yükleyin...");
                    }

                }
            }
        }
    }
}
