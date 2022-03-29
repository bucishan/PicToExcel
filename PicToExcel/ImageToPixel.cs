using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Drawing.Imaging;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace PicToExcel
{
    public partial class ImageToPixel : Form
    {
        private Timer timer;                    //实时清理缓存计时器

        public ImageToPixel()
        {
            InitializeComponent();
            SetTimer();
        }

        private void SetTimer()
        {
            timer = new Timer();
            timer.Tick -= new EventHandler(Timer_Tick);
            timer.Interval = 1000;
            timer.Tick += new EventHandler(Timer_Tick);
            timer.Start();
        }

        /// <summary>
        /// 定时销毁不用的内存
        /// </summary>
        private void Timer_Tick(object sender, EventArgs e)
        {
            Tools.FlushMemory();
        }

        private void ImageToPixel_Load(object sender, EventArgs e)
        {
        }

        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        public bool isCancel = false;           //是否退出
        Bitmap bmpGrap;                         //画布
        Graphics grap;
        /// <summary>
        /// 打开图片
        /// </summary>
        private void btnOpenImg_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "图像文件|*.jpg;*.jpeg;*.bmp;*.png";
            if (openfile.ShowDialog() == DialogResult.OK && openfile.FileName != "")
            {
                Image img = Image.FromFile(openfile.FileName);
                pOriginal.Image = img;
                //pOriginal.ImageLocation = openfile.FileName;
                txtImgPath.Text = openfile.FileName;
            }
            openfile.Dispose();
        }
        /// <summary>
        /// 打开Excel
        /// </summary>
        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "Excel文件|*.xlsx;*.xls";
            if (openfile.ShowDialog() == DialogResult.OK && openfile.FileName != "")
            {
                //pOriginal.ImageLocation = openfile.FileName;
                txtExcelPath.Text = openfile.FileName;
            }
            openfile.Dispose();
        }

        /// <summary>
        /// 中断并保存
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            isCancel = true;
        }

        /// <summary>
        /// 转换成像素
        /// </summary>
        private void btnConvert_Click(object sender, EventArgs e)
        {
            //GetColorDic();
            hebingExcel();
        }

        private void hebingExcel()
        {
            Excel.Application app = new Excel.Application();
            Excel._Workbook result = app.Workbooks.Add();
            Excel._Workbook wb1 = app.Workbooks.Open(Path.GetFullPath("F:\\1.xlsx"));
            Excel._Workbook wb2 = app.Workbooks.Open(Path.GetFullPath("F:\\2.xlsx"));

            Excel._Worksheet sheet = wb1.Sheets[1];


            foreach (Excel._Worksheet each in wb1.Sheets)
            {
                each.Copy(result.Worksheets[1]);
            }
            foreach (Excel._Worksheet each in wb2.Sheets)
            {
                each.Copy(result.Worksheets[1]);
            }
            wb1.Close();
            wb2.Close();
            result.SaveAs("F:\\result.xlsx");
            app.Quit();
        }


        private Dictionary<string, Color> GetColorDic()
        {
            //Excel.Application app = new Excel.Application();
            //Excel._Workbook Workbook = app.Workbooks.Add();
            //Excel._Worksheet Worksheet = Workbook.Sheets.Add();


            Dictionary<string, Color> dicColor = new Dictionary<string, Color>();
             
            Image img = pOriginal.Image;
            Bitmap bitmap = new Bitmap(img);
            int height = bitmap.Height;                         //位图高度
            int width = bitmap.Width;                           //位图宽度
            byte R, G, B;                                       //RGB颜色变量
            int forX = 0, forY = 0;

            BitmapData dataOut = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            int offset = dataOut.Stride - dataOut.Width * 3;

            unsafe
            {
                byte* pOut = (byte*)(dataOut.Scan0.ToPointer());
                //pOut = pOut += (3 * width * forY) + (offset * forY);

                for (int Y = forY; Y < height; Y++)
                {
                    for (int X = 0; X < width; X++)
                    {

                        B = pOut[0];
                        G = pOut[1];
                        R = pOut[2];
                        pOut += 3;
                        Color color = Color.FromArgb(255, R, G, B);//设置颜色
                        //dicColor.Add(string.Format("{0}-{1}", X, Y), color);
                        string colorHtml = ColorTranslator.ToHtml(color);
                    }
                }
                bitmap.UnlockBits(dataOut);
            }
            return dicColor;
        }
        /// <summary>
        /// 保存像素文件
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            GetPixs();
        }

        public void GetPixs()
        {
            string filename = txtImgPath.Text.Substring(0, txtImgPath.Text.LastIndexOf('.')) + ".xlsx";
            int forX = 0, forY = 0;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;      //不显示更改提示
            xlApp.ScreenUpdating = false;     //不显示屏幕刷新
            xlApp.Visible = false;            //不显示界面

            //Workbooks workbooks = xlApp.Workbooks;
            //xlWorkBook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];//取得sheet1

            if (txtExcelPath.Text != string.Empty)
                xlWorkBook = xlApp.Workbooks.Add(@txtExcelPath.Text);
            else
                xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = xlWorkBook.Sheets["Sheet1"];

            //断点继续
            if (txtExcelPath.Text != string.Empty)
            {
                filename = txtExcelPath.Text;
                string temp = ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1]).Text.ToString();
                if (temp != string.Empty)
                {
                    string[] str = temp.Split(';');
                    forX = int.Parse(str[0].Split(':')[1]);
                    forY = int.Parse(str[1].Split(':')[1]);
                }
            }

            Image img = pOriginal.Image;
            //img = ImageZoom.GenerateHighThumbnail(img, this.pOriginal.Width, this.pOriginal.Height);
            Bitmap bitmap = new Bitmap(img);

            int height = bitmap.Height;                         //位图高度
            int width = bitmap.Width;                           //位图宽度
            byte R, G, B;                                       //RGB颜色变量

            int lX = (pOriginal.Width - width) / 2;             //虚拟画布坐标X
            int lY = (pOriginal.Height - height) / 2;           //虚拟画布坐标Y


            this.Invoke(new System.Action(() =>
            {
                lblW.Text = width.ToString();
                lblH.Text = height.ToString();
            }));

            //1、在内存中建立一块“虚拟画布”：
            bmpGrap = new Bitmap(width, height);
            //2、获取这块内存画布的Graphics引用：
            grap = Graphics.FromImage(bmpGrap);

            BitmapData dataOut = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, width, height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            int offset = dataOut.Stride - dataOut.Width * 3;
            int barNum = 0;
            SetBar(barNum, width * height);
            try
            {
                Excel.Range rangeAll = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[height, width]];
                rangeAll.RowHeight = 1.00;
                rangeAll.ColumnWidth = 0.10;

                unsafe
                {
                    byte* pOut = (byte*)(dataOut.Scan0.ToPointer());
                    if (txtExcelPath.Text != string.Empty)
                    {
                        pOut = pOut += (3 * width * forY) + (offset * forY);
                    }
                    for (int Y = forY; Y < height; Y++)
                    {
                        for (int X = 0; X < width; X++)
                        {
                            if (isCancel)
                            {
                                xlWorkSheet.Cells[1, 1] = "X:" + X + ";Y:" + Y;
                                break;
                            }
                            B = pOut[0];
                            G = pOut[1];
                            R = pOut[2];
                            pOut += 3;
                            //if (B == 0 && G == 0 && R == 0)
                            //{
                            //    barNum++;
                            //    continue;
                            //}
                            this.Invoke(new System.Action(() =>
                            {
                                System.Windows.Forms.Application.DoEvents();
                                lblX.Text = X.ToString();
                                lblY.Text = Y.ToString();

                                ////3、在这块内存画布上绘图：
                                //Color color = Color.FromArgb(255, R, G, B);
                                //Brush brush = new SolidBrush(color);
                                //grap.FillRectangle(brush, X, Y, 1, 1);
                                ////4、将内存画布画到窗口中
                                //this.pCurrent.CreateGraphics().DrawImage(bmpGrap, lX, lY);
                            }));

                            Microsoft.Office.Interop.Excel.Range titleRange = xlWorkSheet.Range[(object)xlWorkSheet.Cells[Y + 1, X + 1], (object)xlWorkSheet.Cells[Y + 1, X + 1]];//选取单元格，选取一行或多行 
                            titleRange.Interior.Color = Color.FromArgb(255, R, G, B);//设置颜色

                            System.Windows.Forms.Application.DoEvents();
                            SetBar(barNum++, 0);          //进度条

                            //ThreadPool.QueueUserWorkItem(new WaitCallback(excelTitle), (object)dic);
                        }
                        pOut += offset;
                        if (isCancel)
                        {
                            break;
                        }
                    }
                    bitmap.UnlockBits(dataOut);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            isCancel = false;
            xlWorkBook.Saved = true;
            xlWorkBook.SaveCopyAs(txtExcelPath.Text);
            //xlWorkBook.SaveAs(filename, misValue, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close();
            //xlApp.Quit();
            KillSpecialExcelProcess();

            MessageBox.Show("File created !");
        }
        /// <summary>
        /// 杀死Excel残留进程的另一种方法
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="lpdwProcessId"></param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        public void KillSpecialExcelProcess()
        {
            try
            {
                if (xlApp != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out lpdwProcessId);

                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }
        /// <summary>
        /// 设置进度条
        /// </summary>
        public void SetBar(int value, int maxNum)
        {
            this.Invoke(new System.Action(() =>
            {
                if (maxNum > 0)
                    bar.Maximum = maxNum;
                bar.Value = value;
            }));
        }

        /// <summary>
        /// 测试用线程池处理颜色写入，解决不了Excel格式写入慢的问题
        /// </summary>
        public void excelTitle(object data)
        {
            Dictionary<string, int> dic = (Dictionary<string, int>)data;
            int R = dic["R"]; int G = dic["G"]; int B = dic["B"]; int X = dic["X"]; int Y = dic["Y"];

            this.Invoke(new System.Action(() =>
            {
                lblX.Text = X.ToString();
                lblY.Text = Y.ToString();

                //3、在这块内存画布上绘图：
                Color color = Color.FromArgb(255, R, G, B);
                Brush brush = new SolidBrush(color);
                grap.FillRectangle(brush, X, Y, 1, 1);
                //4、将内存画布画到窗口中
                this.pCurrent.CreateGraphics().DrawImage(bmpGrap, 0, 0);
            }));
            Microsoft.Office.Interop.Excel.Range titleRange = xlWorkSheet.get_Range((object)xlWorkSheet.Cells[Y + 1, X + 1], (object)xlWorkSheet.Cells[Y + 1, X + 1]);//选取单元格，选取一行或多行 
            titleRange.Interior.Color = Color.FromArgb(255, R, G, B);//设置颜色
            titleRange.RowHeight = 5.00;
            titleRange.ColumnWidth = 0.50;
        }

        public void WriteColor(Bitmap bmp, BitmapData bData, int x, int y)
        {
            Color pixelColor = bmp.GetPixel(x, y);
            byte alpha = pixelColor.A;                          //颜色的 Alpha 分量值
            byte red = pixelColor.R;                            //颜色的 RED 分量值
            byte green = pixelColor.G;                          //颜色的 GREEN 分量值
            byte blue = pixelColor.B;                           //颜色的 BLUE 分量值

            Microsoft.Office.Interop.Excel.Range titleRange = xlWorkSheet.get_Range((object)xlWorkSheet.Cells[y, x], (object)xlWorkSheet.Cells[y, x]);//选取单元格，选取一行或多行 
            titleRange.Interior.Color = Color.FromArgb(alpha, red, green, blue);//设置颜色
            titleRange.RowHeight = 5;
            titleRange.ColumnWidth = 0.5;
        }

        public void GraphicsImage()
        {
            //1、在内存中建立一块“虚拟画布”：
            Bitmap bmp = new Bitmap(600, 600);
            //2、获取这块内存画布的Graphics引用：
            Graphics g = Graphics.FromImage(bmp);
            //3、在这块内存画布上绘图
            Color color = Color.FromArgb(255, 1, 1, 1);
            Brush brush = new SolidBrush(color);
            g.FillRectangle(brush, 10, 10, 10, 10);
            g.FillRectangle(brush, 20, 10, 1, 1);
            g.FillRectangle(brush, 30, 10, 5, 5);
            g.FillRectangle(brush, 40, 10, 100, 100);
            //g.FillEllipse(brush, 10, 10, 10, 10);
            //4、将内存画布画到窗口中
            this.pCurrent.CreateGraphics().DrawImage(bmp, 0, 0);
        }

        public void TestExcelOpen()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(@txtExcelPath.Text);
            xlWorkSheet = xlWorkBook.Sheets["Sheet2"];

            //((object)xlWorkSheet.Cells[1, 1]).ToString();
            string temp = ((Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1]).Text.ToString();
            Dictionary<string, int> dic = new Dictionary<string, int>();
            if (temp != string.Empty)
            {
                string[] str = temp.Split(';');
                dic["X"] = int.Parse(str[0].Split(':')[1]);
                dic["Y"] = int.Parse(str[1].Split(':')[1]);
            }
            Microsoft.Office.Interop.Excel.Range titleRange = xlWorkSheet.get_Range((object)xlWorkSheet.Cells[dic["Y"], dic["X"]], (object)xlWorkSheet.Cells[dic["Y"], dic["X"]]);//选取单元格，选取一行或多行 
            titleRange.Interior.Color = Color.FromArgb(255, 12, 23, 34);//设置颜色
            //titleRange.RowHeight = 5.00;
            //titleRange.ColumnWidth = 0.50;
            //isCancel = false;
            xlWorkBook.SaveAs(@txtExcelPath.Text, misValue, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }



    }
}
