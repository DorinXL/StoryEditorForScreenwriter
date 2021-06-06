using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelWinForm
{
    public partial class Form1 : Form
    {
        CheckExcel checkfm = new CheckExcel();
        TransForm transfm = new TransForm();
        AddFrom addfm = new AddFrom();
        public Form1()
        {
            InitializeComponent();
            OpenForm(addfm,this.panel1);
            OpenForm(transfm, this.flowLayoutPanel1);
            this.panel2.Visible = false;
            OpenForm(checkfm, this.panel2);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void OpenForm(Form objFrm,Panel parent)
        {
            objFrm.TopLevel = false; //将子窗体设置成非最高层，非顶级控件
            //objFrm.WindowState = FormWindowState.Maximized;//将当前窗口设置成最大化
            objFrm.FormBorderStyle = FormBorderStyle.None;//去掉窗体边框
            objFrm.Parent = parent;//指定子窗体显示的容器
            objFrm.Show();
        }

        private void 查看excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = false;
            this.flowLayoutPanel1.Visible = false;
            this.panel2.Visible = true;
        }

        private void 修改ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = true;
            this.flowLayoutPanel1.Visible = true;
            this.panel2.Visible = false;
        }

        public void disposeTrans()
        {
            transfm.Dispose();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            ExcelUtil.Instance.CloseExcelApplication();
        }
    }
}
