using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWinForm
{
    class ExcelUtil
    {
        private Excel.Application excelApplication = null;
        private Excel.Workbooks excelWorkbooks;
        private Excel.Workbook excelWorkbook;
        private Excel.Worksheet excelWorksheet;
        public Excel.Range excelRange;
        private int activeSheetIndex;
        private string saveFilePath = string.Empty;
        private string openFilePath = string.Empty;

        private int Row, Col;
        private object[,] dataValueRange;
        public List<string[]> dataEditor = new List<string[]>();
        public int rowValue
        {
            get { return Row; }
        }
        public int colValue
        {
            get { return Col; }
        }
        private ExcelRange ex = new ExcelRange();


        private static ExcelUtil instance = null;
        public static ExcelUtil Instance
        {
            get
            {
                if(instance == null)
                {
                    instance = new ExcelUtil();
                }
                return instance;
            }
        }

        /// <summary>
        /// 获取或设置当前有效活动Sheet索引
        /// </summary>
        public int ActiveSheetIndex
        {
            get { return activeSheetIndex; }
            set { activeSheetIndex = value; }
        }

        /// <summary>
        /// 获取或设置
        /// </summary>
        public Excel.Workbook Workbook
        {
            get { return excelWorkbook; }
            set { excelWorkbook = value; }
        }

        public Excel.Worksheet Worksheet
        {
            get { return excelWorksheet; }
            set { excelWorksheet = value; }
        }

        /// <summary>
        /// 获取设置当前Excel含有的Worksheet数
        /// </summary>
        public int WorksheetsCount
        {
            get
            {
                if (excelWorkbook == null) return 0;
                if (excelWorkbook.Worksheets == null) return 0;
                return excelWorkbook.Worksheets.Count;
            }
        }

        //传入Excel文件路径打开一个Excel文件
        public bool OpenExcelApplication(string path)
        {
            if (excelApplication != null) CloseExcelApplication();
            if (string.IsNullOrEmpty(path)) throw new Exception("请选择一个文件！");

            if (!File.Exists(path))
                throw new Exception(path + "文件不存在！");
            else
            {
                try
                {
                    //点击引用到的第三方组件然后属性中将Embed Interop Types置为False, ActiveSheet.UsedRange.Rows.Count
                    excelApplication = new Excel.ApplicationClass();
                    excelWorkbooks = excelApplication.Workbooks;
                    excelWorkbook = excelWorkbooks.Open(path) as Excel.Workbook;
                    excelWorksheet = excelWorkbook.Worksheets[1] as Excel.Worksheet;
                    excelApplication.Visible = false;

                    return true;
                }
                catch (Exception ex)
                {
                    CloseExcelApplication();
                    throw new Exception(string.Format("（1）程序中没有安装Excel程序。（2）或没有安装Excel所需要支持的.NetFramework\n详细信息：{0}", ex.Message));
                }
            }
        }

        /// <summary>
        /// 关闭Excel程序
        /// </summary>
        /// 每次操作完Excel都要对Excel中所使用到的对象进项释放资源的操作如下
        public void CloseExcelApplication()
        {
            try
            {
                //Save();
                excelWorksheet = null;
                excelWorkbook = null;
                excelWorkbooks = null;
                excelRange = null;
                if (excelApplication != null)
                {
                    excelApplication.Workbooks.Close();
                    excelApplication.Quit();
                    excelApplication = null;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        //读取excel
        public void getExcel()
        {
            string path = Application.StartupPath + "\\Item.xlsx";
            //string path = "E:\\Project\\C#\\ExcelWinForm\\ExcelWinForm\\Item.xlsx";
            OpenExcelApplication(path);
            ex = GetCurrentRegion("A1", Worksheet);
            //Col = ex.Column;
            Col = 5;
            Row = ex.Row;
        }

        //获取到表中数据并保存在dataValueRange中
        //并传给字符串队列dataEditor
        public void getValue(string start, string end)
        {
            excelRange = Worksheet.get_Range(start, end);
            dataValueRange = (object[,])excelRange.Value;
            if(dataValueRange == null)
            {
                return;
            }
            for (int row = 1; row <= Row; row++)
            {
                string[] tmp = new string[Col + 1];
                for (int col = 1; col <= Col; col++)
                {
                    tmp[col - 1] = dataValueRange[row, col] == null ? string.Empty : dataValueRange[row, col].ToString();
                }
                dataEditor.Add(tmp);
            }
        }



        /*
    假设该表的左上角在单元格 a1,并且注意该表中间没有空行和空列,则: 
    sheets['sheet1'].range['a1'].CurrentRegion.Rows.Count
    sheets['sheet1'].range['a1'].CurrentRegion.Columns.Count
    返回该表的行数,Columns返回该表的列数 
*/
        public ExcelRange GetCurrentRegion(string startRange, Excel.Worksheet activeWorksheet)
        {
            return new ExcelRange()
            {
                Row = activeWorksheet.Range[startRange.ToUpper()].CurrentRegion.Rows.Count,
                Column = activeWorksheet.Range[startRange.ToUpper()].CurrentRegion.Columns.Count
            };
        }

    }

    public class ExcelRange
    {
        #region 构造方法

        public ExcelRange()
        {
            this.Row = 0;
            this.Column = 0;
        }

        #endregion

        #region 公共属性

        public int Row { get; set; }

        public int Column { get; set; }

        #endregion
    }
}
