using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Картриджи
{
    public partial class Add : Form
    {
        SqlConnection cann;
        SqlDataReader d2;
        SqlCommand comand;
        DataTable dt;
        string connectionString = @"Data Source=WS-0687058;Initial Catalog=kart;Integrated Security=True";
        public Add()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cann = new SqlConnection(connectionString);
            cann.Open();
            string sql1 = "INSERT INTO kart VALUES ('"+ textBox1.Text + "','" + textBox1.Text + "'," + textBox3.Text +")";
            comand = new SqlCommand(sql1, cann);
            d2 = comand.ExecuteReader();
            dt = new DataTable();
            dt.Load(d2);
            cann.Close();
            this.Close();
        }
    }
}
