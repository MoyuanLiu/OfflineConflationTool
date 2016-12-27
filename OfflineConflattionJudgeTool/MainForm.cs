using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Utils;
using System.Text.RegularExpressions;

namespace OfflineConflattionJudgeTool
{
    public partial class MainForm : Form
    {
        Job job { get; set; }
        List<DataTable> list { get; set; }
        DataTable dt { get; set; }
        public MainForm()
        {           
            InitializeComponent();
        }
        
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewFileForm newtask = new NewFileForm(this);
            newtask.ShowDialog();
        }
        public void GetJob(Job job)
        {
            //显示面板
            groupBox1.Visible = true;
            //清空下拉列表
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            //获取传过来的job
            this.job = job;
            //把当前index赋值给状态栏
            toolStripStatusLabel1.Text = job.CurrentIndex.ToString();
            //获取当前job对应的Excel的sheet,并转换成datatable
            list = ExcelDataTableUtil.ExceltoDataTable(job.JobPath);
            dt = list[0];
            for (int i = 1; i <= dt.Columns.Count; i++)
            {
                comboBox1.Items.Add(i);
                comboBox2.Items.Add(i);
            }
            dt.Columns.Add();
            dt.Columns.Add();
            this.job.CurrentIndex = 0;
            this.job.HasFinished = 0;
            comboBox1.SelectedItem = comboBox1.Items[0];
            comboBox2.SelectedItem = comboBox2.Items[0];
            
            ExcelDataTableUtil.WriteExcel(ExcelDataTableUtil.DataTabletoExcel(dt), job.JobPath);
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            DataRow dr = dt.Rows[job.CurrentIndex];
            textBox4.Text = dr[Convert.ToInt32(comboBox1.Text)].ToString();
            Match m1 = Regex.Match(dr[Convert.ToInt32(comboBox2.Text)].ToString(), @"http(s)?://");
            if (m1.Success)
            {
                webBrowser2.Url = new Uri(dr[Convert.ToInt32(comboBox2.Text)].ToString());
                if (webBrowser2.Url != null)
                {
                    textBox2.Text = webBrowser2.Url.ToString();
                }
            }
            else
            {
                webBrowser2.DocumentText = dr[Convert.ToInt32(comboBox2.Text)].ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
 
            if (job.CurrentIndex != job.Count)
            {
                Comment com = new Comment();
                com.ShowDialog();
                string txt = com.CommentText;
                XSSFWorkbook workbook = ExcelDataTableUtil.LoadExcel(job.JobPath);
                XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
                XSSFRow row = (XSSFRow)sheet.GetRow(job.CurrentIndex);
                row.Cells[dt.Columns.Count - 2].SetCellValue("Y");
                row.Cells[dt.Columns.Count - 1].SetCellValue(txt);
                ExcelDataTableUtil.WriteExcel(workbook, job.JobPath);
                job.CurrentIndex++;
                job.HasFinished++;
                Judge();
            }
            else
            {
                MessageBox.Show("已经是最后一条了");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (job.CurrentIndex != job.Count)
            {
                Comment com = new Comment();
                com.ShowDialog();
                string txt = com.CommentText;
                XSSFWorkbook workbook = ExcelDataTableUtil.LoadExcel(job.JobPath);
                XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
                XSSFRow row = (XSSFRow)sheet.GetRow(job.CurrentIndex);
                row.Cells[dt.Columns.Count - 2].SetCellValue("N");

                XSSFColor XlColour = new XSSFColor(button4.BackColor);
                XSSFCellStyle cellstyle = (XSSFCellStyle)workbook.CreateCellStyle();
                cellstyle.SetFillForegroundColor(XlColour);
                cellstyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                row.Cells[dt.Columns.Count - 2].CellStyle = cellstyle;
                row.Cells[dt.Columns.Count - 1].SetCellValue(txt);
                ExcelDataTableUtil.WriteExcel(workbook, job.JobPath);


                job.CurrentIndex++;
                job.HasFinished++;
                Judge();
            }
            else
            {
                MessageBox.Show("已经是最后一条了");
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (job.CurrentIndex != job.Count)
            {
                Comment com = new Comment();
                com.ShowDialog();
                string txt = com.CommentText;
                XSSFWorkbook workbook = ExcelDataTableUtil.LoadExcel(job.JobPath);
                XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);
                XSSFRow row = (XSSFRow)sheet.GetRow(job.CurrentIndex);
                row.Cells[dt.Columns.Count - 2].SetCellValue("U");

                XSSFColor XlColour = new XSSFColor(button5.BackColor);
                XSSFCellStyle cellstyle = (XSSFCellStyle)workbook.CreateCellStyle();
                cellstyle.SetFillForegroundColor(XlColour);
                cellstyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                row.Cells[dt.Columns.Count - 2].CellStyle = cellstyle;
                row.Cells[dt.Columns.Count - 1].SetCellValue(txt);
                ExcelDataTableUtil.WriteExcel(workbook, job.JobPath);
                job.CurrentIndex++;
                job.HasFinished++;
                Judge();
            }
            else
            {
                MessageBox.Show("已经是最后一条了");
            }
            
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            DataRow dr = dt.Rows[job.CurrentIndex];
            textBox3.Text = dr[Convert.ToInt32(comboBox1.Text)].ToString();
            Match m = Regex.Match(dr[Convert.ToInt32(comboBox1.Text)].ToString(), @"http(s)?://");
            if (m.Success)
            {
                webBrowser1.Url = new Uri(dr[Convert.ToInt32(comboBox1.Text)].ToString());
                if (webBrowser1.Url!=null)
                {
                    textBox1.Text = webBrowser1.Url.ToString();
                }
                
            }
            else
            {
                webBrowser1.DocumentText = dr[Convert.ToInt32(comboBox1.Text)].ToString();
            }
        }
        public void Judge()
        {
            if (job.HasFinished==job.Count)
            {
                toolStripStatusLabel1.Text = "You have finished the work,current index :"+job.CurrentIndex.ToString();
            }
            toolStripStatusLabel1.Text = "current index :"+job.CurrentIndex.ToString();
            
            DataRow dr = dt.Rows[job.CurrentIndex];
            textBox3.Text = dr[Convert.ToInt32(comboBox1.Text)].ToString();
            textBox4.Text = dr[Convert.ToInt32(comboBox2.Text)].ToString();
            if (checkBox1.Checked == true)
            {
                webBrowser1.Url = new Uri(@"http://knowledge.microsoft.com/Search?market=en-us&q=" + dr[Convert.ToInt32(comboBox1.Text)].ToString());
                if (webBrowser1.Url != null)
                {
                    textBox1.Text = webBrowser1.Url.ToString();
                }
            }
            else if (checkBox2.Checked == true)
            {
                webBrowser2.Url = new Uri(@"http://knowledge.microsoft.com/Search?market=en-us&q=" + dr[Convert.ToInt32(comboBox2.Text)].ToString());
                if (webBrowser2.Url != null)
                {
                    textBox2.Text = webBrowser2.Url.ToString();
                }
            }
            
            else
            {
                Match m = Regex.Match(dr[Convert.ToInt32(comboBox1.Text)].ToString(), @"http(s)?://");
                if (m.Success)
                {
                    webBrowser1.Url = new Uri(dr[Convert.ToInt32(comboBox1.Text)].ToString());
                    if (webBrowser1.Url != null)
                    {
                        textBox1.Text = webBrowser1.Url.ToString();
                    }
                }
                else
                {
                    webBrowser1.DocumentText = dr[Convert.ToInt32(comboBox1.Text)].ToString();
                }
                Match m1 = Regex.Match(dr[Convert.ToInt32(comboBox2.Text)].ToString(), @"http(s)?://");
                if (m1.Success)
                {
                    webBrowser2.Url = new Uri(dr[Convert.ToInt32(comboBox2.Text)].ToString());
                    if (webBrowser2.Url != null)
                    {
                        textBox2.Text = webBrowser2.Url.ToString();
                    }
                }
                else
                {
                    webBrowser2.DocumentText = dr[Convert.ToInt32(comboBox2.Text)].ToString();
                }
            }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (job.CurrentIndex != 1)
            {
                job.CurrentIndex--;
                Judge();
            }
            else
            {
                MessageBox.Show("已经是第一条了");
            }
            
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (job.CurrentIndex != job.Count)
            {
                job.CurrentIndex++;
                Judge();
            }
            else
            {
                MessageBox.Show("已经是最后一条了");
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (job == null)
            {
                groupBox1.Visible = false;
            }
            else
            {
                groupBox1.Visible = true;
            }
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            DataRow dr = dt.Rows[job.CurrentIndex];
            if (checkBox1.Checked==true)
            {
                webBrowser1.Url = new Uri(@"http://knowledge.microsoft.com/Search?market=en-us&q="+dr[Convert.ToInt32(comboBox1.Text)].ToString());
            }
        }

        private void checkBox2_CheckStateChanged(object sender, EventArgs e)
        {
            DataRow dr = dt.Rows[job.CurrentIndex];
            if (checkBox2.Checked == true)
            {
                webBrowser2.Url = new Uri(@"http://knowledge.microsoft.com/Search?market=en-us&q=" + dr[Convert.ToInt32(comboBox2.Text)].ToString());
            }
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (webBrowser1.Url != null)
            {
                textBox1.Text = webBrowser1.Url.ToString();
            }
        }

        private void webBrowser2_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (webBrowser2.Url != null)
            {
                textBox2.Text = webBrowser2.Url.ToString();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void webBrowser1_NewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void webBrowser2_NewWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
    }
    
}

