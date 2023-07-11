using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;

namespace scheduling
{
    public partial class Form_main : Form
    {

        public string refresh_string = "SELECT 專案編號,分析方法,數量,人員,日期 FROM 測試 ORDER BY 人員,分析方法,日期 ASC";
        //public string connection_string = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\college\111-2\project\code\scheduling\scheduling2\scheduling\Database1.mdf;Integrated Security=True";
        public string connection_string;
        public string excel_path;
        public Form_main()
        {
            InitializeComponent();
        }

        private void Form_main_Load(object sender, EventArgs e)
        {

            string filePath = @"..\..\Database1.mdf";
            string path = System.IO.Path.GetFullPath(filePath);
            //path = @"D:\college\111-2\project\code\scheduling\scheduling2\scheduling\Database1.mdf";
            connection_string = $"Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename={path};Integrated Security=True";

            //path = @"..\2022年分析組工作排程檢記進度表.xlsx";
            //excel_path = System.IO.Path.GetFullPath(path);
            path = @"excel_path_record.txt"; // Replace with the actual file path
            try
            {
                // Create a StreamReader to read the file
                using (StreamReader reader = new StreamReader(path))
                {
                    string line;

                    // Read and display each line of the file
                    while ((line = reader.ReadLine()) != null)
                    {
                        excel_path = line;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            
            refresh_worker();
            refresh_task();

        }

        private void refresh_task()
        {
            //Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\college\111-2\project\code\scheduling\scheduling\scheduling\Database1.mdf;Integrated Security=True
            //建立SqlConnection物件db

            SqlConnection db4 = new SqlConnection();
            db4.ConnectionString = connection_string;
            db4.Open();
            SqlDataAdapter da4 = new SqlDataAdapter(refresh_string, db4);

            DataSet ds4 = new DataSet();
            da4.Fill(ds4, "測試");
            dataGridView1.DataSource = ds4;
            dataGridView1.DataMember = "測試";
            db4.Close();

        }
        private void refresh_worker()
        {
            //Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\college\111-2\project\code\scheduling\scheduling\scheduling\Database1.mdf;Integrated Security=True
            //建立SqlConnection物件db

            SqlConnection db4 = new SqlConnection();
            string connectionString = connection_string;


            // Define your SQL query
            string query = "SELECT Id FROM Worker";

            // Create a SqlConnection object
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                // Open the connection
                connection.Open();

                // Create a SqlCommand object with the query and connection
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Create a SqlDataReader to read the data
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Create a DataTable to store the data
                        DataTable dataTable = new DataTable();

                        // Load the data from the SqlDataReader into the DataTable
                        dataTable.Load(reader);

                        // Set the data source and display member of the ComboBox
                        comboBox1.DataSource = dataTable;
                        comboBox1.DisplayMember = "Id";

                        // Set the value member (optional)
                        //comboBox1.ValueMember = "Id";

                        // Bind the data
                        //comboBox1.DataBindings.;
                    }
                }
                connection.Close();
            }
        }
        static int GetExcelProcessId(Excel.Application excelApp)
        {
            int processId = 0;

            try
            {
                // Get the window handle of the Excel application
                IntPtr hwnd = new IntPtr(excelApp.Hwnd);

                // Get the process ID associated with the window handle
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses();
                foreach (System.Diagnostics.Process process in processes)
                {
                    if (process.MainWindowHandle == hwnd)
                    {
                        processId = process.Id;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return processId;
        }
        private void add_sql(string task_no, string method, string num, string worker, string date)
        {
            try
            {//,天數,檢測項目/設備名稱,分析方法,數量,課別,案件負責人
                SqlConnection db = new SqlConnection();
                db.ConnectionString = connection_string;
                db.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = db;
                //cmd.CommandText = "SELECT MAX(Id) FROM 專案";
                cmd.CommandText = "INSERT INTO 測試(專案編號,分析方法,數量,人員,日期)VALUES('" +
                    task_no + "',N'" +     //// 雙引號取代單引號
                    method + "','" +
                    num + "',N'" +
                    worker + "','" +
                    date + "')";
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
        private void delete_all()
        {
            try
            {
                SqlConnection db = new SqlConnection();
                db.ConnectionString = connection_string;
                db.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = db;
                cmd.CommandText = "DELETE FROM 測試";
                cmd.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            refresh_task();
        }
        private void ReadExcelFile(string filePath)
        {
            Process[] runingProcess = Process.GetProcesses();
            List<int> pid_array = new List<int>();
            for (int i = 0; i < runingProcess.Length; i++)
            {
                // compare equivalent process by their name                 
                if (runingProcess[i].ProcessName == "EXCEL")
                {
                    //keep running process                    
                    pid_array.Add(runingProcess[i].Id);
                }

            }
            Excel.Application excelApp = new Excel.Application();
            //Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            //dynamic sheet = null;
            delete_all();
            //int processId = GetExcelProcessId(excelApp);
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = workbook.Sheets[1]; // Assuming the data is in the first sheet

                range = worksheet.UsedRange;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;
                string task_no = "";
                string method = "";
                string num = "";
                string worker = "";
                string date = "";
                label6.Text = rowCount.ToString();
                for (int i = 7; i <= rowCount; i++)
                {
                    //label6.Text = "" + i;
                    //for (int j = 1; j <= colCount; j++)
                    //{
                    /*
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                    {
                        string cellValue = range.Cells[i, j].Value2.ToString();
                        Console.WriteLine("Cell({0},{1}): {2}", i, j, cellValue);
                    }
                    */
                    if (range.Cells[i, 15].Value2 == null && range.Cells[i, 6].Value2 == null && range.Cells[i, 16].Value2 == null && range.Cells[i, 11].Value2 == null && range.Cells[i, 12].Value2 == null)
                    {
                        break;
                    }
                    if (range.Cells[i, 15] != null && range.Cells[i, 15].Value2 != null)
                    {
                        //人員
                        if (range.Cells[i, 15].Value2.ToString() == comboBox1.Text)
                        {
                            //label1.Text = range.Cells[i, 15].Value2.ToString();
                        }
                        else
                        {
                            ///continue;
                        }
                        if((bool)range.Cells[i, 15].Font.Strikethrough == true)
                        {
                            continue;
                        }
                        //label1.Text = range.Cells[i, 15].Value2.ToString();
                        worker = range.Cells[i, 15].Value2.ToString();
                    }

                    if (range.Cells[i, 16] != null && range.Cells[i, 16].Value2 != null)
                    {
                        //日期
                        //get date
                        double date_temp = (double)range.Cells[i, 16].Value2;
                        DateTime temp = DateTime.FromOADate(date_temp);
                        int day_compare_value = DateTime.Compare(temp, dateTimePicker1.Value.AddDays(-1));
                        int day_compare_value2 = DateTime.Compare(temp, dateTimePicker2.Value);
                        if (day_compare_value >= 0)
                        {
                            if (day_compare_value2 < 0)
                            {
                                //label2.Text = temp.ToString();
                                date = temp.ToString("yyyy-MM-dd");
                            }
                            else
                                continue;
                        }
                        else
                        {
                            continue;
                        }
                        if ((bool)range.Cells[i, 16].Font.Strikethrough == true)
                        {
                            continue;
                        }
                    }
                    if (range.Cells[i, 6] != null && range.Cells[i, 6].Value2 != null)
                    {
                        //專案編號
                        if ((bool)range.Cells[i, 6].Font.Strikethrough == true)
                        {
                            continue;
                        }
                        task_no = range.Cells[i, 6].Value2.ToString();
                        //label3.Text = range.Cells[i, 6].Value2.ToString();
                    }
                    if (range.Cells[i, 11] != null && range.Cells[i, 11].Value2 != null)
                    {
                        //方法
                        if ((bool)range.Cells[i, 11].Font.Strikethrough == true)
                        {
                            continue;
                        }
                        method = range.Cells[i, 11].Value2.ToString();
                        //label4.Text = range.Cells[i, 11].Value2.ToString();
                    }
                    if (range.Cells[i, 12] != null && range.Cells[i, 12].Value2 != null)
                    {
                        //數量
                        if ((bool)range.Cells[i, 12].Font.Strikethrough == true)
                        {
                            continue;
                        }
                        num = range.Cells[i, 12].Value2.ToString();
                        //label5.Text = range.Cells[i, 12].Value2.ToString();
                    }

                    add_sql(task_no, method, num, worker, date);
                    //}
                }

                //label1.Text = range.Cells[7, 16].Value2.ToString();
                /*
                //get date
                double date_temp = (double)range.Cells[7, 16].Value2;
                DateTime temp = DateTime.FromOADate(date_temp);
                int day_compare_value = DateTime.Compare(temp, dateTimePicker1.Value);
                if(day_compare_value < 0)
                {
                   label2.Text = "early";
                }
                else
                {
                   label2.Text = "late";
                }
                */
                //workbook.Close();
                //excelApp.Quit();
            }
            catch
            {
                //excelApp.Quit();
            }
            finally
            {
                //注意: Excel是Unmanaged程式，要妥善結束才能乾淨不留痕跡
                //否則，很容易留下一堆excel.exe在記憶體中
                //所有用過的COM+物件都要使用Marshal.FinalReleaseComObject清掉
                //COM+物件的Reference Counter，以利結束物件回收記憶體
                /*
                if (range != null)
                {
                    Marshal.FinalReleaseComObject(range);
                }
                if (worksheet != null)
                {
                    Marshal.FinalReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    workbook.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
                    Marshal.FinalReleaseComObject(workbook);
                }
                */
                if (excelApp != null)
                {
                    excelApp.Workbooks.Close();
                    excelApp.Quit();
                    /*
                    try
                    {
                        Process process = Process.GetProcessById(processId);
                        process.Kill();
                        Console.WriteLine("Process with ID " + processId + " has been terminated.");
                    }
                    catch (ArgumentException)
                    {
                        Console.WriteLine("Process with ID " + processId + " does not exist.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                    }
                    */
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                    bool keep_excel = false;
                    Process[] runingProcess2 = Process.GetProcesses();
                    for (int i = 0; i < runingProcess2.Length; i++)
                    {
                        // compare equivalent process by their name 
                        keep_excel = false;
                        if (runingProcess2[i].ProcessName == "EXCEL")
                        {
                            //kill  running process      
                            foreach(int pid in pid_array){
                                if(runingProcess2[i].Id == pid)
                                {
                                    keep_excel = true;
                                    break;
                                }
                            }
                            if(!keep_excel)
                                runingProcess2[i].Kill();
                            //break;
                        }
                    }
                    

                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ReadExcelFile(excel_path);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            refresh_worker();
            refresh_task();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value.AddDays(6);
        }

        private void button_method_Click(object sender, EventArgs e)
        {
            Form_method method = new Form_method();
            method.Show(this);
            refresh_task();
        }

        private void button_worker_Click(object sender, EventArgs e)
        {
            Form_worker worker = new Form_worker();
            worker.Show(this);
            refresh_task();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string connectionString = connection_string;
                string query = "SELECT 專案編號, 分析方法, 數量, 人員, 日期 FROM 測試 WHERE 人員 = N'" + comboBox1.Text + "' ORDER BY 人員,分析方法,日期 ASC";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Create a DataTable to store the query result
                        DataTable dataTable = new DataTable();

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            // Fill the DataTable with the query result
                            adapter.Fill(dataTable);
                        }

                        // Assign the DataTable as the data source for the DataGridView
                        dataGridView1.DataSource = dataTable;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button_file_select_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                excel_path = openFileDialog1.FileName;
                string ppath = @"excel_path_record.txt";
                using (StreamWriter sw = File.CreateText(ppath))
                {
                    sw.WriteLine(excel_path);
                }
            }
        }

        private void button_clear_process_Click(object sender, EventArgs e)
        {
            Process[] runingProcess = Process.GetProcesses();
            for (int i = 0; i < runingProcess.Length; i++)
            {
                // compare equivalent process by their name                 
                if (runingProcess[i].ProcessName == "EXCEL")
                {
                    //kill  running process                    
                    runingProcess[i].Kill();
                    //break;
                }
                /*
                try
                {
                    //Pass the filepath and filename to the StreamWriter Constructor
                    //Write a line of text
                    string ppath = @"Test.txt";
                    if (!File.Exists(ppath))
                    {
                        // Create a file to write to.
                        using (StreamWriter sw = File.CreateText(ppath))
                        {
                            //sw.WriteLine("Hello");
                            //sw.WriteLine("And");
                            //sw.WriteLine("Welcome");
                        }
                    }

                    // This text is always added, making the file longer over time
                    // if it is not deleted.
                    using (StreamWriter sw = File.AppendText(ppath))
                    {
                        sw.WriteLine(runingProcess[i].ProcessName);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }
                */
            }
        }
    }
}
