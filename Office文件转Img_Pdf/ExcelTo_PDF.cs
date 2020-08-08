using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office文件转Img_Pdf
{
    /// <summary>
    /// Excel转PDF
    /// </summary>
    class ExcelTo_PDF
    {
        /// <summary>
        /// Excel转PDF
        /// </summary>
        /// <param name="ExcelInputPath">读取Excel文件路径</param>
        /// <param name="PDFOutputPath">PDF输出路径</param>
        /// <param name="imageName">PDF名称</param>
        public static void ExcelToPDF(string ExcelInputPath, string PDFOutputPath, string imageName)
        {

            Microsoft.Office.Interop.Excel.Application lobjExcelApp = null;
            Microsoft.Office.Interop.Excel.Workbooks lobjExcelWorkBooks = null;
            Microsoft.Office.Interop.Excel.Workbook lobjExcelWorkBook = null;


            object lobjMissing = System.Reflection.Missing.Value;
            //PDF文件保存路径和PDF文件名称
            string PDFInputPath = PDFOutputPath + imageName + ".pdf";

            try
            {
                lobjExcelApp = new Microsoft.Office.Interop.Excel.Application();
                //是否显示Excel文件界面
                lobjExcelApp.Visible = false;

                lobjExcelWorkBooks = lobjExcelApp.Workbooks;
                //打开Excel文件进行读取
                lobjExcelWorkBook = lobjExcelWorkBooks.Open(ExcelInputPath, false, false, lobjMissing, lobjMissing, lobjMissing, false,
                  lobjMissing, lobjMissing, lobjMissing, lobjMissing, lobjMissing, false, lobjMissing, lobjMissing);


                //输出为PDF 第一个选项指定转出为PDF,还可以指定为XPS格式  
                lobjExcelWorkBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, PDFInputPath);//, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false, Type.Missing, Type.Missing, false, Type.Missing);


            }
            catch (Exception e)
            {
                //打印错误
                Console.WriteLine(e);

            }
            finally
            {
                Process[] localByNameApp = Process.GetProcessesByName(ExcelInputPath);//获取程序名的所有进程
                if (localByNameApp.Length > 0)
                {
                    foreach (var app in localByNameApp)
                    {
                        if (!app.HasExited)
                        {
                            #region
                            //设置禁止弹出保存和覆盖的询问提示框   
                            lobjExcelApp.DisplayAlerts = false;
                            lobjExcelApp.AlertBeforeOverwriting = false;
                            //确保Excel进程关闭   
                            lobjExcelApp.Quit();
                            lobjExcelApp = null;
                            #endregion
                            app.Kill();//关闭进程  
                        }
                    }
                }
                if (lobjExcelWorkBook != null)
                    lobjExcelWorkBook.Close(true, Type.Missing, Type.Missing);
                lobjExcelWorkBooks.Close();
                lobjExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lobjExcelApp);
                // 安全回收进程
                System.GC.GetGeneration(lobjExcelApp);
            }

        }
    }
}
