using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelWinForm
{
    public partial class AddFrom : Form
    {
        List<string> commandType = new List<string>();
        List<string> command = new List<string>();
        List<string> text = new List<string>();
        List<string> pos = new List<string>();
        ExcelUtil ce = ExcelUtil.Instance;
        public AddFrom()
        {
            InitializeComponent();
            controlVisible("#");
            listCreate();
            this.comboBox1.DataSource = commandType;

        }


        private void button1_Click(object sender, EventArgs e) 
        {

            //ce.dataEditor.Add(tmp);
            if(textBox1.Text != string.Empty)
            {
                if(int.Parse(textBox1.Text) > 0)
                {
                    getValue(int.Parse(textBox1.Text) - 1);
                }
            }
        }

        void getValue(int line)
        {
            string[] emp = new string[5];
            if(line + 1 > ce.dataEditor.Count)
            {
                line = ce.dataEditor.Count + 1 - 1;
                textBox1.Text = (ce.dataEditor.Count + 1).ToString();
                ce.dataEditor.Add(emp);
            }
            switch (comboBox1.SelectedItem.ToString())
            {
                case "#":
                    emp[0] = "#";
                    emp[1] = richTextBox1.Text;
                    ce.dataEditor[line] = emp;
                    break;
                case "Command":
                    emp[0] = "Command";
                    emp[1] = comboBox2.SelectedItem.ToString();
                    switch (emp[1]) 
                    {
                        case "DelAll":
                            break;
                        case "SetCharacter":
                            emp[2] = textBox2.Text;
                            emp[3] = comboBox3.Text;
                            break;
                        case "ChangeEmojy":
                            emp[2] = textBox2.Text;
                            emp[3] = comboBox3.Text;
                            emp[4] = textBox3.Text;
                            break;
                        default:
                            emp[2] = textBox2.Text;
                            break;
                    }
                    ce.dataEditor[line] = emp;
                    break;
                case "Text":
                    emp[0] = "Text";
                    emp[1] = comboBox2.SelectedItem.ToString();
                    switch (emp[1])
                    {
                        case "Monologue":
                            if (textBox2.Text == string.Empty) emp[2] = "*";
                            else emp[2] = textBox2.Text;
                            emp[3] = "/";
                            if (textBox3.Text == string.Empty) emp[4] = "*";
                            else emp[4] = textBox3.Text;
                            break;
                        case "Dialogue":
                            if (textBox2.Text == string.Empty) emp[2] = "*";
                            else emp[2] = textBox2.Text;
                            emp[3] = comboBox3.SelectedItem.ToString();
                            if (textBox3.Text == string.Empty) emp[4] = "*";
                            else emp[4] = textBox3.Text;
                            break;
                        default:
                            break;
                    }
                    ce.dataEditor[line] = emp;
                    break;
                case "Stop":
                    emp[0] = "Stop";
                    ce.dataEditor[line] = emp;
                    break;
                case "插一行":
                    ce.dataEditor.Insert(line, emp);
                    break;
                case "删除该行":
                    ce.dataEditor.RemoveAt(line);
                    break;
                default:
                    break;
            }
            TransForm.instance.cle();
            textBox1.Text = (int.Parse(textBox1.Text) + 1).ToString();
        }

        //第一个comboBox切换数据的时候
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            controlVisible(comboBox1.SelectedItem.ToString());
            controlContentChange(comboBox1.SelectedItem.ToString());
        }

        //行数的输入
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //如果输入的不是退格和数字，则屏蔽输入
            if (!(e.KeyChar == 8 || (e.KeyChar >= 48 && e.KeyChar <= 57)))
            {
                e.Handled = true;
            }
        }

        //各种控件的可视性
        void controlVisible(string control)
        {
            switch (control)
            {
                case "#":
                    richTextBox1.Visible = true;
                    comboBox2.Visible = false;
                    textBox2.Visible = false;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;

                    break;
                case "Command":
                    richTextBox1.Visible = false;
                    comboBox2.Visible = true;
                    textBox2.Visible = true;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;

                    break;
                case "Text":
                    richTextBox1.Visible = false;
                    comboBox2.Visible = true;
                    textBox2.Visible = true;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;

                    break;
                //case "Stop":
                //    richTextBox1.Visible = false;
                //    comboBox2.Visible = false;
                //    textBox2.Visible = false;
                //    comboBox3.Visible = false;
                //    textBox3.Visible = false;

                //    break;
                default:
                    richTextBox1.Visible = false;
                    comboBox2.Visible = false;
                    textBox2.Visible = false;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;

                    break;
            }
        }

        //切换第二个框的内容
        void controlContentChange(string control)
        {
            switch (control)
            {
                case "Command":
                    comboBox2.DataSource = command;
                    textBox3.Visible = false;
                    break;
                case "Text":
                    comboBox2.DataSource = text;
                    comboBox3.Visible = true;
                    textBox3.Visible = true;
                    break;
                default:
                    break;
            }
        }

        //第二个comboBox切换数据的时候
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedItem.ToString())
            {
                case "DelAll":
                    textBox2.Visible = false;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;
                    break;
                case "SetCharacter":
                    textBox2.Visible = true;
                    comboBox3.Visible = true;
                    comboBox3.DataSource = pos;
                    textBox3.Visible = false;

                    break;
                case "ChangeEmojy":
                    textBox2.Visible = true;
                    comboBox3.Visible = true;
                    comboBox3.DataSource = pos;
                    textBox3.Visible = true;

                    break;
                case "Monologue":
                    textBox2.Visible = true;
                    comboBox3.Visible = true;
                    comboBox3.DataSource = pos;
                    textBox3.Visible = true;
                    break;
                case "Dialogue":
                    textBox2.Visible = true;
                    comboBox3.Visible = true;
                    comboBox3.DataSource = pos;
                    textBox3.Visible = true;
                    break;
                default:
                    textBox2.Visible = true;
                    comboBox3.Visible = false;
                    textBox3.Visible = false;
                    break;
            }
        }

        void listCreate()
        {
            #region command添加
            command.Add("SetBackground");
            command.Add("SetBgm");
            command.Add("DelBackground");
            command.Add("DelCharacter");
            command.Add("DelAll");
            command.Add("SetCharacter");
            command.Add("UnlockCG");
            command.Add("ChangeEmojy");

            #endregion


            #region Text添加
            text.Add("Monologue");
            text.Add("Dialogue");

            #endregion


            #region 位置添加
            pos.Add("LEFT");
            pos.Add("RIGHT");
            pos.Add("*");
            #endregion


            #region type添加
            commandType.Add("#");
            commandType.Add("Command");
            commandType.Add("Text");
            commandType.Add("Stop");
            commandType.Add("插一行");
            commandType.Add("删除该行");
            #endregion

        }

        //第三个comboBox切换数据的时候
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //写入excel
        private void button2_Click(object sender, EventArgs e)
        {
            object[,] obj = new object[ce.dataEditor.Count + 1, ce.colValue + 1];
            for (int i = 0; i < ce.dataEditor.Count; i++)
            {
                for (int j = 0; j < ce.colValue; j++)
                {
                    obj[i, j] = ce.dataEditor[i][j] == null ? string.Empty : ce.dataEditor[i][j];
                }
            }
            string end = ((char)('A' + ce.colValue - 1)) + ce.dataEditor.Count.ToString();
            ce.excelRange = ce.Worksheet.get_Range("A1", end);
            ce.excelRange.Value = obj;
            //string path = "E:\\Project\\C#\\ExcelWinForm\\ExcelWinForm\\Item2.xlsx";
            string path = Application.StartupPath + "\\Item.xlsx";
            ce.Workbook.SaveAs(path);
            TransForm.instance.cle();
            CheckExcel.instance.PaintUpdate();
        }
    }
}
