using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Utils;

namespace OfflineConflattionJudgeTool
{
    public partial class Comment : Form
    {
        public string CommentText { get; set; }



        public Comment()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommentText = textBox1.Text.Trim();
            MessageBox.Show("Already judge");
            Close();
        }
        
    }
}
