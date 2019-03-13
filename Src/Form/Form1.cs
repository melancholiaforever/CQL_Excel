using System;
using System.Windows.Forms;
using System.Data;
using DataTable = System.Data.DataTable;
using ClassLibrary7;
using System.Threading;

namespace ExcelAddIn1
{
    public partial class Query : Form
    {
        public Query()
        {
            InitializeComponent();
        }
        public event Action<string, string, string, string, string> ImportDataAction;
        private void Button1_Click(object sender, EventArgs e)
        {
            var selectedlist2 = listBox2.SelectedItem as DataRowView;
            var table = selectedlist2[0].ToString();
            var tbname = selectedlist2[1].ToString();
            var selectedlist1 = listBox1.SelectedItem as DataRowView;
            var dataset = selectedlist1[0].ToString();
            if (checkBox1.Checked == false)
            {
                textBox2.Text = "10000";
            }
            if (textBox2.Enabled == true && textBox2.Text == string.Empty)
            {
                textBox2.Text = "10000";
            }
            if (checkBox2.Checked == false)
            {
                textBox3.Text = "0";
            }
            if (textBox3.Enabled == true && textBox3.Text == string.Empty)
            {
                textBox3.Text = "0";
            }
            ImportDataAction(dataset, table, textBox2.Text, textBox3.Text, tbname);
            this.Close();
        }
        DataTable list1 = BLL.GetDataTableflex("select * from quandl_description limit 10000");

        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.DataSource = list1;
            listBox1.DisplayMember = "name_en";
            listBox1.ValueMember = "database";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox2.Text) > 100000)
            {
                textBox2.Text = "100000";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox2.Enabled = true;
            }
            else
            {
                textBox2.Enabled = false;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                textBox3.Enabled = true;
            }
            else
            {
                textBox3.Enabled = false;
            }

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(textBox3.Text) > 100000000)
            {
                textBox3.Text = "100000000";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selected = listBox1.SelectedItem as DataRowView;
            webBrowser2.DocumentText = selected[1].ToString();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var selected = listBox1.SelectedItem as DataRowView;
            table = selected[0].ToString();
            string countsql = "select count(*) from quandl_updateindex where databasecode = '" + table + "'";
            DataTable count = BLL.GetDataTableflex(countsql);
            total = System.Convert.ToDouble(count.Rows[0][0]);
            if (total >= 400000)
            {
                MessageBox.Show("The Index is to big to using Excel to query, please using SQL to query");
            }
            else
            {
                listBox1.Enabled = false;
                backgroundWorker1.RunWorkerAsync();
                //double total = System.Convert.ToDouble(count.Rows[0][0]);
                //double tf = total / 10000;
                //double totalfor = Math.Floor(tf);
                //string sql1 = "select quandlcode,datasetcode,dataname,description from quandl_updateindex where databasecode = '" + table + "' limit 10000";
                //DataTable list2 = BLL.GetDataTableflex(sql1);
                //if (totalfor >= 1)
                //{
                //    for (int i = 1; i <= totalfor; i++)
                //    {
                //        int offsetnumber = i * 10000;
                //        string sql = "select quandlcode,datasetcode,dataname,description from quandl_updateindex where databasecode = '" + table + "' limit 10000 offset " + offsetnumber.ToString() + "";
                //        DataTable temp = BLL.GetDataTableflex(sql);
                //        int rows = temp.Rows.Count;
                //        for (int j = 0; j < rows; j++)
                //        {
                //            list2.ImportRow(temp.Rows[j]);
                //        }
                //    }
                //}
                //listBox2.DataSource = list2;
                //listBox2.DisplayMember = "dataname";
                //listBox2.ValueMember = "quandlcode";
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selected = listBox2.SelectedItem as DataRowView;
            webBrowser1.DocumentText = selected[3].ToString();
            var selectedlist2 = listBox2.SelectedItem as DataRowView;
            var table = selectedlist2[0].ToString();
            var selectedlist1 = listBox1.SelectedItem as DataRowView;
            var dataset = selectedlist1[0].ToString();
            string sql = "select * from quandl_" + dataset + " where quandlcode = '" + table + "'";
            DataTable dt = BLL.GetDataTableflex(sql);
            dt = RemoveNULLColumns(dt);
            dataGridView1.DataSource = dt;
        }

        //Form2 waitingbar = new Form2();

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            //waitingbar.ShowDialog();
            //double total = System.Convert.ToDouble(count.Rows[0][0]);
            double tf = total / 10000;
            double totalfor = Math.Floor(tf);
            string sql1 = "select quandlcode,datasetcode,dataname,description from quandl_updateindex where databasecode = '" + table + "' limit 10000";
            list2 = BLL.GetDataTableflex(sql1);
            if (totalfor >= 1)
            {
                for (int i = 1; i <= totalfor; i++)
                {
                    int offsetnumber = i * 10000;
                    string sql = "select quandlcode,datasetcode,dataname,description from quandl_updateindex where databasecode = '" + table + "' limit 10000 offset " + offsetnumber.ToString() + "";
                    DataTable temp = BLL.GetDataTableflex(sql);
                    int rows = temp.Rows.Count;
                    int report = (int)Math.Round(i / totalfor) * 100;
                    for (int j = 0; j < rows; j++)
                    {
                        list2.ImportRow(temp.Rows[j]);
                    }

                    backgroundWorker1.ReportProgress(report);
                    Thread.Sleep(1000);
                }
            }
            //waitingbar.Close();
        }



        public double total { get; set; }
        public DataTable list2 { get; set; }
        public string table { get; set; }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            //修改进度条的显示。
            //progressBar1.Value = (e.ProgressPercentage);

            //如果有更多的信息需要传递，可以使用 e.UserState 传递一个自定义的类型。
            //这是一个 object 类型的对象，您可以通过它传递任何类型。
            //我们仅把当前 sum 的值通过 e.UserState 传回，并通过显示在窗口上。
            //string message = e.UserState.ToString();
            //this.labelSum.Text = message;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            listBox2.DataSource = list2;
            listBox2.DisplayMember = "dataname";
            listBox2.ValueMember = "quandlcode";
            MessageBox.Show("Secondry index Completed！");
            //progressBar1.Value = 0;
            backgroundWorker1.Dispose();
            listBox1.Enabled = true;
        }



        private DataTable RemoveNULLColumns(DataTable data)//删除空列
        {
            DataTable dt = data;
            for (int i = dt.Columns.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrEmpty(Convert.ToString(dt.Rows[0][i])))
                {
                    dt.Columns.RemoveAt(i);
                }
            }
            return dt;
        }

        //public DataTable list3 { get; set; }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            //DataTable list3 = new DataTable();
            //string importtext = textBox1.Text;
            //list3 = list2.Clone();
            //DataRow[] dr = list2.Select("dataname like '%importtext%'");
            //foreach (DataRow row in dr)  // 将查询的结果添加到dt中； 
            //{
            //    list3.Rows.Add(row.ItemArray);
            //}
            //listBox2.DataSource = list3;
            //listBox2.DisplayMember = "dataname";
            //listBox2.ValueMember = "quandlcode";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable list3 = new DataTable();
            string importtext = textBox1.Text;
            list3 = list2.Clone();
            DataRow[] dr = list2.Select("dataname like '%" + importtext + "%'");
            for (int i = 0; i < dr.Length; i++)
            {
                list3.ImportRow(dr[i]);
            }
            listBox2.DataSource = list3;
            listBox2.DisplayMember = "dataname";
            listBox2.ValueMember = "quandlcode";
        }

        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        ////public event Action<string> ImportDataAction2;
        //public void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        //{
        //    textBox3.Text = openFileDialog1.FileName;
        //    ImportDataAction2(textBox3.Text);
        //}

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    openFileDialog1.ShowDialog();
        //    string path = openFileDialog1.FileName;
        //}

        //private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (!(((e.KeyChar >= '0') && (e.KeyChar <= '9')) || e.KeyChar <= 31))
        //    {
        //        if (e.KeyChar == '.')
        //        {
        //            e.Handled = true;
        //        }
        //        else
        //            e.Handled = true;
        //    }
        //    else
        //    {
        //        if (e.KeyChar <= 31)
        //        {
        //            e.Handled = false;
        //        }
        //        else if ((e.KeyChar >= '0') && (e.KeyChar <= '9'))
        //        {
        //            if (((TextBox)sender).Text.ToString() != "")
        //            {
        //                if (Convert.ToDouble(((TextBox)sender).Text) == 0)
        //                {
        //                    if (((TextBox)sender).Text.Trim().IndexOf('.') > -1)
        //                    {
        //                        e.Handled = false;
        //                    }
        //                    else
        //                    {
        //                        e.Handled = true;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                e.Handled = false;
        //            }
        //        }
        //    }
        //}
    }
}
