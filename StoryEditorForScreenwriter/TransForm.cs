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
    public partial class TransForm : Form
    {
        ExcelUtil util = ExcelUtil.Instance;

        public static TransForm instance;
        //public static TransForm Instance
        //{
        //    get
        //    {
        //        if (instance == null)
        //        {
        //            instance = ;
        //        }
        //        return instance;
        //    }
        //}

        public TransForm()
        {
            InitializeComponent();
            instance = this;
            this.listView1.Columns.Add("行数", 40, HorizontalAlignment.Left);
            this.listView1.Columns.Add("内容", 480, HorizontalAlignment.Left); //一步添加

            piant();
        }

        public void piant()
        {
            this.listView1.Items.Clear();

            this.listView1.BeginUpdate();

            string cellValue;
            for (int row = 1; row <= util.dataEditor.Count; row++)
            {
                ListViewItem lvi = new ListViewItem();
                List<string> content = new List<string>();
                lvi.Text = row.ToString();
                for (int col = 1; col <= util.colValue; col++)
                {
                    cellValue = util.dataEditor[row - 1][col - 1] == null ? string.Empty : util.dataEditor[row - 1][col - 1].ToString();
                    //lvi.SubItems.Add(cellValue);
                    content.Add(cellValue);
                }
                string trans = transform(content);
                lvi.SubItems.Add(trans);
                this.listView1.Items.Add(lvi);
            }

            this.listView1.EndUpdate();
        }

        public void DataUpdate() { 
        }

        public void cle()
        {
            this.listView1.Items.Clear();
            piant();
        }

        string transform(List<string> tmp)
        {
            string trans = "";
            switch (tmp[0])
            {
                case "#":
                    trans = "这是一句注释: " + tmp[1];
                    break;
                case "Command":
                    switch (tmp[1])
                    {
                        case "SetBackground":
                            trans = "设置 背景 ：" + tmp[2];
                            break;
                        case "SetBgm":
                            trans = "设置 BGM ：" + tmp[2];
                            break;
                        case "DelBackground":
                            trans = "删除 背景 ：" + tmp[2];
                            break;
                        case "DelCharacter":
                            trans = "删除 单个人物 ：" + tmp[2];
                            break;
                        case "DelAll":
                            trans = "删除 所有人物";
                            break;
                        case "SetCharacter":
                            trans = "设置 人物：" + tmp[2] + " 在 " + tmp[3];
                            break;
                        case "UnlockCG":
                            trans = "解锁 CG：" + tmp[2];
                            break;
                        case "ChangeEmojy":
                            trans = "设置 人物：" + tmp[2] + " 在 " + tmp[3] + " 表情为 " + tmp[4];
                            break;
                    }
                    break;
                case "Text":
                    switch (tmp[1])
                    {
                        case "Monologue":
                            trans = "【独白】" + tmp[2] + ":" + tmp[4];
                            break;
                        case "Dialogue":
                            trans = "【对话】" + "位于 " + tmp[3] + " 的 " + tmp[2] + " 说：" + tmp[4]; 
                            break;
                        default:
                            break;
                    }
                    break;
                case "Stop":
                    trans = "结束游戏";
                    break;
                default:
                    trans = "此行为空 或 指令未定义";
                    break;

            }
            return trans;
        }
    }
}
