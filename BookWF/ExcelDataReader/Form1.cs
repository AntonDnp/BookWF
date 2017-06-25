using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Novacode;
using System.Diagnostics;


namespace ExcelDataReader
{
    public partial class Form1 : Form
    {
        DataSet ds;
        string Author, Title;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                FileStream stream =File.Open(ofd.FileName,FileMode.Open,FileAccess.Read);

                Excel.IExcelDataReader IEDR;

                int fileformat = ofd.SafeFileName.IndexOf(".xlsx");

                if (fileformat > -1)
                {
                    //2007 format *.xlsx
                    IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                else
                {
                    //97-2003 format *.xls
                    IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                }

                //Если данное значение установлено в true
                //то первая строка используется в качестве 
                //заголовков для колонок
                IEDR.IsFirstRowAsColumnNames = true;

                 ds = IEDR.AsDataSet();

                // выводим названия всех таблиц
                //foreach(DataTable dt in ds.Tables)
                //{
                //    MessageBox.Show(dt.TableName);
                //}


                //Устанавливаем в качестве источника данных dataset 
                //с указанием номера таблицы. Номер таблицы указавает 
                //на соответствующий лист в файле нумерация листов 
                //начинается с нуля.
                dataGridView1.DataSource = ds.Tables[0];
                IEDR.Close();

                // выбираем из DataSet только столбец Автор
                IEnumerable<string> query = from book in ds.Tables[0].AsEnumerable()
                            select book.Field<string>("Автор");

                // При помощи метода Distinct() удаляем повторяющиеся значения из запроса
                IEnumerable<string> distinctNames = query.Distinct();

                // Занесли значения в combobox1
                comboBox1.Items.AddRange(distinctNames.ToArray());

                // выбираем из DataSet только столбец Год издания
                IEnumerable<double> query1 = from book in ds.Tables[0].AsEnumerable()
                                            select book.Field<double>("Год издания");

                // При помощи метода Distinct() удаляем повторяющиеся значения из запроса
                IEnumerable<double> distinctY = query1.Distinct();

                // Занесли значения в combobox3
                //comboBox3.Items.AddRange(distinctY.ToArray());

                foreach(double year in distinctY)
                {
                    comboBox3.Items.Add(year);
                }

                //foreach (string name in distinctNames)
                //{
                //    MessageBox.Show(name);
                //}



            }
            else
            {
                MessageBox.Show("Вы не выбрали файл для открытия",
                 "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }




        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Title=comboBox2.SelectedItem.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName = @"D:\Фриланс\DocXExample1.docx";
            var doc = DocX.Create(fileName);
            //doc.InsertParagraph(Author);
            //doc.InsertParagraph(Title);

        // Выбираем номер стеллажа и полки на которых стоит выбранная книга
           var query = from book in ds.Tables[0].AsEnumerable()
                                        where book.Field<string>("Автор") == comboBox1.SelectedItem.ToString()
                                        where book.Field<string>("Название") == comboBox2.SelectedItem.ToString()
                                        select new {Stellag= book.Field<double>("Номер стеллажа"), Polka= book.Field<double>("Номер полки") };
            // заносим в файл word
            foreach (var book in query)
            {
                doc.InsertParagraph(book.Stellag.ToString());
                doc.InsertParagraph(book.Polka.ToString());
            }

            doc.Save();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int count = 0;
            string fileName = @"D:\Фриланс\DocXExample2.docx";
            var doc = DocX.Create(fileName);

            // отбираем из таблицы только книги года выбранного в combobox3
            IEnumerable<string> query = from book in ds.Tables[0].AsEnumerable()
                                        where book.Field<double>("Год издания") == (double)comboBox3.SelectedItem
                                        select book.Field<string>("Название");

            count = query.Count();
            doc.InsertParagraph("Количество книг:"+count);
            doc.Save();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            comboBox2.Items.Clear();
            string SelectedName= comboBox1.SelectedItem.ToString();
            Author = comboBox1.SelectedItem.ToString();

            // отбираем из таблицы только книги автора выбранного в comboBox1
            IEnumerable<string> query = from book in ds.Tables[0].AsEnumerable()
                                        where book.Field<string>("Автор")== SelectedName
                                        select book.Field<string>("Название");

            comboBox2.Items.AddRange(query.ToArray());

        }
    }
}
