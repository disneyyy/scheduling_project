using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace scheduling
{
    public partial class Form_worker : Form
    {
        public Form_worker()
        {
            InitializeComponent();
        }

        private void Form_worker_Load(object sender, EventArgs e)
        {
            refresh_task();
        }
        public string refresh_string = "SELECT Id FROM Worker ORDER BY Id ASC";
        public string connection_string = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\college\111-2\project\code\scheduling\scheduling2\scheduling\Database1.mdf;Integrated Security=True";
        private void refresh_task()
        {
            //Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\college\111-2\project\code\scheduling\scheduling\scheduling\Database1.mdf;Integrated Security=True
            //建立SqlConnection物件db

            SqlConnection db4 = new SqlConnection();
            db4.ConnectionString = connection_string;
            db4.Open();
            SqlDataAdapter da4 = new SqlDataAdapter(refresh_string, db4);

            DataSet ds4 = new DataSet();
            da4.Fill(ds4, "Worker");
            dataGridView1.DataSource = ds4;
            dataGridView1.DataMember = "Worker";
            db4.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {//,天數,檢測項目/設備名稱,分析方法,數量,課別,案件負責人
                SqlConnection db = new SqlConnection();
                db.ConnectionString = connection_string;
                db.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = db;
                //cmd.CommandText = "SELECT MAX(Id) FROM 專案";
                cmd.CommandText = "INSERT INTO Worker(Id)VALUES(N'" +
                    textBox1.Text + "')";
                cmd.ExecuteNonQuery();
                db.Close();
                //Form_main_Load(sender, e);
                refresh_task();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlConnection db = new SqlConnection();
                db.ConnectionString = connection_string;
                db.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = db;
                cmd.CommandText = "DELETE FROM Worker WHERE Id = N'" + textBox1.Text + "'";
                cmd.ExecuteNonQuery();
                db.Close();
                refresh_task();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
