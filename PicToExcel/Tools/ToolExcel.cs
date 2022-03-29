using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace PicToExcel
{
    public static partial class Tools
    {
        #region ----------读取、导出Excel----------

        public static Excel.Application m_xlApp = null;

        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <returns></returns>
        public static System.Data.DataTable ExcelToDataTable()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel(*.xlsx;*.xls)|*.xlsx;*.xls";
            openFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFile.Multiselect = false;
            if (openFile.ShowDialog() == DialogResult.Cancel)
            {
                return null;
            }
            string ImportExcelPath = openFile.FileName;
            System.Data.DataTable dtExcel = new System.Data.DataTable();
            Excel.Range range = null;

            try
            {
                //创建Excel对象
                m_xlApp = new Excel.Application();
                m_xlApp.DisplayAlerts = false;      //不显示更改提示
                m_xlApp.Visible = false;            //不显示界面
                m_xlApp.ScreenUpdating = false;     //不显示屏幕刷新
                //打开Excel
                object missing = System.Reflection.Missing.Value;
                Excel.Workbook workbook = m_xlApp.Workbooks.Open(ImportExcelPath);

                //获取数据Sheel页,工作薄从1开始，不是0
                Excel.Worksheet worksheet = workbook.Worksheets[1];

                //获取有效行数和字段数
                int clmCount = worksheet.UsedRange.Columns.Count;
                int rowCount = worksheet.UsedRange.Rows.Count;
                //Excel.Range rangeTitle = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, clmCount]];
                //Excel.Range rangeData = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[rowCount - 1, clmCount]];
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, clmCount]];
                dtExcel = ConvertToDataTable(range.Value2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出异常：" + ex.Message);
            }
            finally
            {
                KillSpecialExcel();
                //EndReport();
            }
            return dtExcel;
        }

        /// <summary>  
        /// 反一个M行N列的二维数组转换为DataTable  
        /// </summary>  
        /// <param name="Arrays">M行N列的二维数组</param>  
        /// <returns>返回DataTable</returns>
        public static System.Data.DataTable ConvertToDataTable(object[,] Arrays)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            int rowCount = Arrays.GetLength(0);
            int clmCount = Arrays.GetLength(1);
            //添加列
            for (int i = 1; i <= clmCount; i++)
            {
                dt.Columns.Add(Arrays[1, i].ToString(), typeof(string));
            }

            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    System.Data.DataRow dr = dt.NewRow();
                    for (int j = 1; j <= clmCount; j++)
                    {
                        dr[j - 1] = Arrays[i, j] == null ? "" : Arrays[i, j].ToString();
                    }
                    dt.Rows.Add(dr);
                }
            }
            catch (Exception)
            {
                throw;
            }
            return dt;
        }

        /// <summary>
        /// 将DataTable数据导出到Excel表
        /// 支持自动sheel分页
        /// </summary>
        /// <param name="tmpDataTable">要导出的DataTable</param>
        /// <param name="saveDataPath">Excel的保存路径及名称</param>
        public static void DataTableToExcel(string saveDataPath, System.Data.DataTable tmpDataTable)
        {
            //进度显示
            System.Windows.Forms.Application.DoEvents();

            long rowNum = tmpDataTable.Rows.Count;//行数
            int columnNum = tmpDataTable.Columns.Count;//列数
            m_xlApp = new Excel.Application();
            m_xlApp.DisplayAlerts = false;      //不显示更改提示
            m_xlApp.ScreenUpdating = false;     //不显示屏幕刷新
            m_xlApp.Visible = false;            //不显示界面

            Excel.Workbooks workbooks = m_xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1

            try
            {
                if (rowNum > 65536)//单张Excel表格最大行数
                {
                    long pageRows = 65535;//定义每页显示的行数,行数必须小于65536
                    int scount = (int)(rowNum / pageRows);//导出数据生成的表单数
                    if (scount * pageRows < rowNum)//当总行数不被pageRows整除时，经过四舍五入可能页数不准
                    {
                        scount = scount + 1;
                    }
                    for (int sc = 1; sc <= scount; sc++)
                    {
                        if (sc > 1)
                        {
                            object missing = System.Reflection.Missing.Value;
                            worksheet = (Excel.Worksheet)workbook.Worksheets.Add(
                                        missing, missing, missing, missing);//添加一个sheet
                        }
                        else
                        {
                            worksheet = (Excel.Worksheet)workbook.Worksheets[sc];//取得sheet1
                        }
                        string[,] datas = new string[pageRows + 1, columnNum];

                        for (int i = 0; i < columnNum; i++) //写入字段
                        {
                            datas[0, i] = tmpDataTable.Columns[i].Caption;//表头信息
                        }
                        //Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]);
                        Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
                        range.Interior.ColorIndex = 15;//15代表灰色
                        range.Font.Bold = true;
                        range.Font.Size = 9;

                        int init = int.Parse(((sc - 1) * pageRows).ToString());
                        int r = 0;
                        int index = 0;
                        int result;
                        if (pageRows * sc >= rowNum)
                        {
                            result = (int)rowNum;
                        }
                        else
                        {
                            result = int.Parse((pageRows * sc).ToString());
                        }

                        for (r = init; r < result; r++)
                        {
                            index = index + 1;
                            for (int i = 0; i < columnNum; i++)
                            {
                                object obj = tmpDataTable.Rows[r][tmpDataTable.Columns[i].ToString()];
                                datas[index, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式
                            }
                            System.Windows.Forms.Application.DoEvents();
                            //添加进度条
                        }

                        Excel.Range fchR = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[index + 1, columnNum]];
                        fchR.Value2 = datas;
                        worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。
                        m_xlApp.WindowState = Excel.XlWindowState.xlMaximized;//Sheet表最大化
                        range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[index + 1, columnNum]];
                        //range.Interior.ColorIndex = 15;//15代表灰色
                        range.Font.Size = 9;
                        range.RowHeight = 14.25;
                        range.Borders.LineStyle = 1;
                        range.HorizontalAlignment = 1;
                    }
                }
                else
                {
                    string[,] datas = new string[rowNum + 1, columnNum];
                    for (int i = 0; i < columnNum; i++) //写入字段
                    {
                        datas[0, i] = tmpDataTable.Columns[i].Caption;
                    }
                    Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
                    range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Bold = true;
                    range.Font.Size = 9;

                    int r = 0;
                    for (r = 0; r < rowNum; r++)
                    {
                        for (int i = 0; i < columnNum; i++)
                        {
                            object obj = tmpDataTable.Rows[r][tmpDataTable.Columns[i].ToString()];
                            datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式
                        }
                        System.Windows.Forms.Application.DoEvents();
                        //添加进度条
                    }
                    Excel.Range fchR = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                    fchR.Value2 = datas;

                    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。
                    m_xlApp.WindowState = Excel.XlWindowState.xlMaximized;

                    range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                    //range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Size = 9;
                    range.RowHeight = 14.25;
                    range.Borders.LineStyle = 1;
                    range.HorizontalAlignment = 1;
                }
                workbook.Saved = true;
                workbook.SaveCopyAs(saveDataPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出异常：" + ex.Message);
            }
            finally
            {
                KillSpecialExcel();
                //EndReport();
            }
        }

        /// <summary>
        /// 杀死Excel残留进程的另一种方法
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        public static void KillSpecialExcel()
        {
            try
            {
                if (m_xlApp != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(m_xlApp.Hwnd), out lpdwProcessId);

                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }

        /// <summary>
        /// 退出报表时关闭Excel和清理垃圾Excel进程
        /// </summary>
        private static void EndReport()
        {
            object missing = System.Reflection.Missing.Value;
            try
            {
                m_xlApp.Workbooks.Close();
                m_xlApp.Workbooks.Application.Quit();
                m_xlApp.Application.Quit();
                m_xlApp.Quit();
            }
            catch { }
            finally
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Application);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp);
                    m_xlApp = null;
                }
                catch { }
                try
                {
                    //清理垃圾进程
                    killProcessThread();
                }
                catch { }
                GC.Collect();
            }
        }
        /// <summary>
        /// 杀掉不死进程
        /// </summary>
        private static void killProcessThread()
        {
            ArrayList myProcess = new ArrayList();
            for (int i = 0; i < myProcess.Count; i++)
            {
                try
                {
                    System.Diagnostics.Process.GetProcessById(int.Parse((string)myProcess[i])).Kill();
                }
                catch { }
            }
        }
        #endregion

        #region ----------内存回收----------

        [DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);
        public static void FlushMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
        }

        #endregion
    }
}
