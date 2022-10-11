using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using ExcelDataReader;

namespace OBS_Date_Check
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTableCollection tableCollection;

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                comboBox1.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
            for (int b = 0; b < dataGridView1.Columns.Count; b++)
            {
                dataGridView1.Columns[b].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            int columnUpdIdx = 0;
            int columnStsIdx = 0;
            int rowIdx = dataGridView1.Rows.Count - 1;
            string columnUpdName = dt.Columns[columnUpdIdx].ColumnName;
            string columnStsName = dt.Columns[columnStsIdx].ColumnName;

            while (!string.Equals(columnStsName, "Status"))
            {
                columnStsIdx++;
                columnStsName = dt.Columns[columnStsIdx].ColumnName;
            }

            while (!string.Equals(columnUpdName, "Updates"))
            {
                columnUpdIdx++;
                columnUpdName = dt.Columns[columnUpdIdx].ColumnName;
            }


            for (int i = 0; i < rowIdx; i++)
            {
                if (!string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[columnUpdIdx].Value.ToString())) 
                {
                    string siString = dataGridView1.Rows[i].Cells[columnUpdIdx].Value.ToString();
                    string siDate = siString.Substring(0, 11);
                    DateTime obsDate1 = Convert.ToDateTime(siDate);
                    DateTime obsDate2 = Convert.ToDateTime(DateTime.Now.ToString("MMM dd yyyy", CultureInfo.CreateSpecificCulture("en-GB")));
                    TimeSpan obsDays = obsDate2 - obsDate1;
                    int x = int.Parse(obsDays.Days.ToString());
                    if ((string.Equals(dataGridView1.Rows[i].Cells[columnStsIdx].Value.ToString(), "Open") && x >= 7))
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    }
                }

                else
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
            }
        }


    }
}