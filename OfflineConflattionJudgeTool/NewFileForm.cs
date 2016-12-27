using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Utils;

namespace OfflineConflattionJudgeTool
{
    public partial class NewFileForm : Form
    {
        MainForm mainform { get; set;}
        public NewFileForm(MainForm mainform)
        {
            InitializeComponent();
            this.mainform = mainform;
        }
        List<string> sheetnames = new List<string>();
        List<XSSFSheet> sheets; 
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (openFileDialog1.FileName.Contains(".xlsx"))
                {
                    txtPath.Text = openFileDialog1.FileName;
                    if (txtPath.Text.Trim() != null)
                    {
                        sheets = ExcelDataTableUtil.GetAllSheets(txtPath.Text.Trim());
                        foreach (var item in sheets)
                        {
                            sheetnames.Add(item.SheetName);
                        }
                        cboxSheet.Items.AddRange(sheetnames.ToArray());
                    }
                }
                else
                {
                    MessageBox.Show("文件格式错误！！");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboxSheet.Text.Trim()))
            {
                label7.Text = "请选择工作簿";
            }
            else if (string.IsNullOrEmpty(cboxSheet.Text.Trim()))
            {
                label4.Text = "请选择sheet";
            }
            else if (string.IsNullOrEmpty(txtFirstrow.Text.Trim()))
            {
                label5.Text = "请输入起始行号";
            }
            else if (string.IsNullOrEmpty(txtLastrow.Text.Trim()))
            {
                label6.Text = "请输入截至行号";
            }
            else if (!Regex.IsMatch(txtLastrow.Text.Trim(), @"\d+"))
            {
                label6.Text = "必须是整数";
            }
            else if (!Regex.IsMatch(txtFirstrow.Text.Trim(), @"\d+"))
            {
                label5.Text = "必须是整数";
            }
            else if (Convert.ToInt32(txtFirstrow.Text.Trim()) > Convert.ToInt32(txtLastrow.Text.Trim()))
            {
                label5.Text = "起始行号不能大于截至行号";
            }
            else
            {
                string target = Path.GetFileNameWithoutExtension(txtPath.Text.Trim()) + txtFirstrow.Text.Trim() + "_" + txtLastrow.Text.Trim();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "tasks", target + ".xlsx");
                XSSFSheet sheet = ExcelDataTableUtil.GetSheetbyName(txtPath.Text.Trim(), cboxSheet.Text.Trim());
                int lastnum = Convert.ToInt32(txtLastrow.Text.Trim());
                int firstnum = Convert.ToInt32(txtFirstrow.Text.Trim());
                DataTable dt = ExcelDataTableUtil.SheetToDataTable(sheet, firstnum-1, lastnum-1);
                XSSFWorkbook workbook = ExcelDataTableUtil.DataTabletoExcel(dt);
                ExcelDataTableUtil.WriteExcel(workbook, path);
                Job job = new Job(new Guid(), target, DateTime.Now.ToString(), path);
                
                job.JobPath = path;
                job.Count = (lastnum - firstnum + 1);
                mainform.GetJob(job);
                DialogResult result = MessageBox.Show("Success");
                if (result == DialogResult.OK)
                {
                    Close();
                }
            }   
        }

        private void NewFileForm_Load(object sender, EventArgs e)
        {

        }
    }
}
