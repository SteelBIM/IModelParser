using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IModelParser
{
    public partial class MatchForm : Form
    {
        public MatchForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.filepath.Text = fbd.SelectedPath + "\\";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filepath = this.filepath.Text;
            string target = this.target.Text;
            if(filepath == null || filepath == "")
            {
                MessageBox.Show("文件路径为空！");
            }
            else if(target == null || target == "")
            {
                MessageBox.Show("请输入匹配字符串");
            }
            FileOutput fo = new FileOutput();
            fo.regex(filepath, target);
        }
    }
}
