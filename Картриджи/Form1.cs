using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Data;

using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;

namespace Картриджи
{

    public partial class Form1 : Form
    {
        private Word.Application wordapp;
        private Word.Documents worddocuments;
        private Word.Document worddocument;
        private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;
        public string pathToFile = @"C:\Users\8\Desktop\data.xlsx";
        public List<string> UseKart = new List<string>();
        public List<string> UseCode = new List<string>();
        public List<string> UsePlace = new List<string>();
        public List<string> UsePrint = new List<string>();
        SqlConnection cann;
        SqlDataReader d2,d3;
        SqlCommand comand;
        DataTable dt,dt1;
        string connectionString = @"Data Source=WS-0687058;Initial Catalog=kart;Integrated Security=True";
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                cann = new SqlConnection(connectionString);
                cann.Open();
                string sql1 = "SELECT  NameKart as 'Название картриджа',NamePrint as 'Название принтера', countkart as 'Количество', Code as 'Код номенклатуры', place as 'Место установки' FROM kart";
                comand = new SqlCommand(sql1, cann);
                d2 = comand.ExecuteReader();
                dt = new DataTable();
                dt.Load(d2);
                bindingSource1.DataSource = dt;
                dataGridView1.DataSource = bindingSource1;
                //q.Clear();
                // Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                //Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathToFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                //Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                //ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                // Указываем номер столбца (таблицы Excel) из которого будут считываться данные.
                //int numCol = 8;

                //Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
                //System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
                //string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
                //listBox1.Items.AddRange(strArray);
                // Выходим из программы Excel.
                //ObjExcel.Quit();
                cann.Close();
            }
            catch
            {
                MessageBox.Show("Похоже, что вы не подключили базу данных", "Ошибка!");
            }


        }

    private void button4_Click(object sender, EventArgs e)
        {
            //Zav a = new Zav();
            //a.ShowDialog();
            wordapp = new Word.Application();
            wordapp.Visible = true;
            Object template = Type.Missing;
            Object newTemplate = false;
            Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            //Меняем шаблон
            template = @"C:\проекты\Картриджи\Картриджи\bin\Debug\1123 заявка на картридж.docx";
            //Создаем документ 2 worddocument в данном случае создаваемый объект 
            worddocument = wordapp.Documents.Add(
             ref template, ref newTemplate, ref documentType, ref visible);
            for (int i = 1; i <= UseKart.Count; i++)
            {
                worddocument.Tables[2].Rows.Add();
                Word.Range worddocument1 = worddocument.Tables[2].Cell(i + 1, 1).Range;
                worddocument1.Text = i.ToString();
                worddocument1.Bold = 0;
                Word.Range worddocument2 = worddocument.Tables[2].Cell(i + 1, 2).Range;
                worddocument2.Text = UseKart[i-1];
                worddocument2.Bold = 0;
                Word.Range worddocument3 = worddocument.Tables[2].Cell(i + 1, 3).Range;
                worddocument3.Text =  "шт";
                worddocument3.Bold = 0;
                Word.Range worddocument4 = worddocument.Tables[2].Cell(i + 1, 4).Range;
                worddocument4.Text = Convert.ToString(1);
                worddocument4.Bold = 0;
                Word.Range worddocument5 = worddocument.Tables[2].Cell(i + 1, 5).Range;
                worddocument5.Text = UsePlace[i-1] + "/" + UsePrint[i - 1] + "/"+ UseCode[i-1] + "/ Взамен выработавших ресурс";
                worddocument5.Bold = 0;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            cann = new SqlConnection(connectionString);
            cann.Open();
            string sql1 = "UPDATE kart SET countkart = countkart - 1 WHERE NameKart = '" + dataGridView1.CurrentRow.Cells[0].Value.ToString() +"'and place='"+ dataGridView1.CurrentRow.Cells[4].Value.ToString() + "'";
            comand = new SqlCommand(sql1, cann);
            d2 = comand.ExecuteReader();
            cann.Close();
            cann.Open();
            string sql2 = "SELECT  NameKart as 'Название картриджа',NamePrint as 'Название принтера', countkart as 'Количество', Code as 'Код номенклатуры', place as 'Место установки' FROM kart";
            comand = new SqlCommand(sql2, cann);
            d2 = comand.ExecuteReader();
            dt = new DataTable();
            dt.Load(d2);
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1;
            cann.Close();
            UseKart.Add(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            UsePrint.Add(dataGridView1.CurrentRow.Cells[1].Value.ToString());
            UseCode.Add(dataGridView1.CurrentRow.Cells[3].Value.ToString());
            UsePlace.Add(dataGridView1.CurrentRow.Cells[4].Value.ToString());
        }


        private void button3_Click(object sender, EventArgs e)
        {
            cann = new SqlConnection(connectionString);
            cann.Open();
            string sql1 = "DELETE FROM kart WHERE NameKart = '" + dataGridView1.CurrentCell.Value.ToString() + "'";
            comand = new SqlCommand(sql1, cann);
            d2 = comand.ExecuteReader();
            cann.Close();
            cann.Open();
            string sql2 = "SELECT  NameKart as 'Название картриджа',NamePrint as 'Название принтера', countkart as 'Количество', Code as 'Код номенклатуры', place as 'Место установки' FROM kart";
            comand = new SqlCommand(sql2, cann);
            d2 = comand.ExecuteReader();
            dt = new DataTable();
            dt.Load(d2);
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Add a = new Add();
            a.ShowDialog();
        }


        private void button7_Click(object sender, EventArgs e)
        {
            cann = new SqlConnection(connectionString);
            cann.Open();
            string sql1 = "UPDATE kart SET countkart = countkart + 1 WHERE NameKart = '" + dataGridView1.CurrentCell.Value.ToString() + "'";
            comand = new SqlCommand(sql1, cann);
            d2 = comand.ExecuteReader();
            cann.Close();
            cann.Open();
            string sql2 = "SELECT  NameKart as 'Название картриджа',NamePrint as 'Название принтера', countkart as 'Количество', Code as 'Код номенклатуры', place as 'Место установки' FROM kart";
            comand = new SqlCommand(sql2, cann);
            d2 = comand.ExecuteReader();
            dt = new DataTable();
            dt.Load(d2);
            bindingSource1.DataSource = dt;
            dataGridView1.DataSource = bindingSource1;
        }
    }
}
