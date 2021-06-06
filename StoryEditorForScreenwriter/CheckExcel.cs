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
    public partial class CheckExcel : Form
    {
        public static CheckExcel instance = null;

        public CheckExcel()
        {
            InitializeComponent();

            instance = this;

            //1.列表头创建
            this.listView1.Columns.Add("行数", 40, HorizontalAlignment.Left);
            this.listView1.Columns.Add("指令类型", 120, HorizontalAlignment.Left); //一步添加
            this.listView1.Columns.Add("指令方向", 120, HorizontalAlignment.Left);
            this.listView1.Columns.Add("选取对象", 120, HorizontalAlignment.Left);
            this.listView1.Columns.Add("对象位置", 200, HorizontalAlignment.Left);
            this.listView1.Columns.Add("交谈内容", 200, HorizontalAlignment.Left);


            InitPaint();
            ////3.显示项
            //foreach (ListViewItem item in this.listView1.Items)
            //{
            //    for (int i = 0; i < item.SubItems.Count; i++)
            //    {
            //        MessageBox.Show(item.SubItems[i].Text);
            //    }
            //}

            ////4.移除某项
            //foreach (ListViewItem lvi in listView1.SelectedItems)  //选中项遍历  
            //{
            //    listView1.Items.RemoveAt(lvi.Index); // 按索引移除  
            //                                         //listView1.Items.Remove(lvi);   //按项移除  
            //}

            ////5.行高设置
            //ImageList imgList = new ImageList();

            //imgList.ImageSize = new Size(10, 20);// 设置行高 20 //分别是宽和高  

            //listView1.SmallImageList = imgList; //这里设置listView的SmallImageList ,用imgList将其撑大  

            ////6.清空
            //this.listView1.Clear();  //从控件中移除所有项和列（包括列表头）。  

            //this.listView1.Items.Clear();  //只移除所有的项。  
        }

        void InitPaint()
        {
            this.listView1.Items.Clear();
            //2.添加数据项
            ExcelUtil tes = ExcelUtil.Instance;
            tes.getExcel();
            string colvalue = ((char)('A' + tes.colValue - 1)) + tes.rowValue.ToString();
            tes.getValue("A1", colvalue);

            this.listView1.BeginUpdate();   //数据更新，UI暂时挂起，直到EndUpdate绘制控件，可以有效避免闪烁并大大提高加载速度  


            for (int row = 1; row <= tes.dataEditor.Count; row++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.SubItems[0].Text = row.ToString();
                for (int col = 1; col < tes.dataEditor[row - 1].Length; col++)
                {
                    lvi.SubItems.Add(tes.dataEditor[row - 1][col - 1]);
                }
                this.listView1.Items.Add(lvi);
            }

            this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。  
        }

        public void PaintUpdate()
        {
            this.listView1.Items.Clear();
            ExcelUtil tes = ExcelUtil.Instance;
            this.listView1.BeginUpdate();
            for (int row = 1; row <= tes.dataEditor.Count; row++)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.SubItems[0].Text = row.ToString();
                for (int col = 1; col < tes.dataEditor[row - 1].Length; col++)
                {
                    lvi.SubItems.Add(tes.dataEditor[row - 1][col - 1]);
                }
                this.listView1.Items.Add(lvi);
            }
            this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。  
        }
    }
}
