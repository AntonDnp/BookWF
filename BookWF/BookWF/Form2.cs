using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BookWF
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {

                
                //openFileDialog1.DefaultExt = "*.xls;";
                //openFileDialog1.Filter= "Excel 2003(*.xls)";   ошибка
                //openFileDialog1.Title = "Выбирите файл";

                 if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //MessageBox.Show(openFileDialog1.FileName);
                    // строка подключения
                    var connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @openFileDialog1.FileName + ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                    OleDbConnection con = new OleDbConnection(connectstring);

                    con.Open();// открыли подключение   ошибка

                    DataSet ds = new DataSet();

                    DataTable schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    List<string> excelSheets = (from DataRow row in schemaTable.Rows select row["TABLE NAME"].ToString()).ToList();

                    foreach (string str in excelSheets)
                    {
                        MessageBox.Show(str);
                    }
                    //string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                    //string select = String.Format("SELECT * FROM [{0}]", sheet1);

                }

            }

            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            
        }
    }
}
