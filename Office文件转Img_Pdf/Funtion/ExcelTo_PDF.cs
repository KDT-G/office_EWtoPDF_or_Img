using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp;
using Aspose.Cells;
using System.IO;

namespace Office_ToDLL.Funtion
{
    /// <summary>
    /// Excel转PDF
    /// </summary>
    public class ExcelTo_PDF
    {

        #region Microsoft.Office.Interop.Excel-旧代码
        //public static void ExcelToPDF1(string ExcelInputPath, string PDFOutputPath, string FileOutType=null)
        //{

        //    Excel.Application lobjExcelApp = null;
        //    Excel.Workbooks lobjExcelWorkBooks = null;
        //    Excel.Workbook lobjExcelWorkBook = null;


        //    object lobjMissing = System.Reflection.Missing.Value;
        //    //PDF文件保存路径和PDF文件名称
        //    string PDFInputPath = FileOutType==null ? PDFOutputPath + ".pdf": PDFOutputPath + FileOutType;

        //    try
        //    {
        //        lobjExcelApp = new Excel.Application();
        //        //是否显示Excel文件界面
        //        lobjExcelApp.Visible = true;

        //        lobjExcelWorkBooks = lobjExcelApp.Workbooks;
        //        //打开Excel文件进行读取
        //        lobjExcelWorkBook = lobjExcelWorkBooks.Open(ExcelInputPath, false, false, lobjMissing, lobjMissing, lobjMissing, false,
        //          lobjMissing, lobjMissing, lobjMissing, lobjMissing, lobjMissing, false, lobjMissing, lobjMissing);


        //        //输出为PDF 第一个选项指定转出为PDF,还可以指定为XPS格式  
        //        lobjExcelWorkBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, PDFInputPath);//, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false, Type.Missing, Type.Missing, false, Type.Missing);


        //    }
        //    catch (Exception e)
        //    {
        //        //打印错误
        //        Console.WriteLine(e);

        //    }
        //    finally
        //    {
        //        Process[] localByNameApp = Process.GetProcessesByName(ExcelInputPath);//获取程序名的所有进程
        //        if (localByNameApp.Length > 0)
        //        {
        //            foreach (var app in localByNameApp)
        //            {
        //                if (!app.HasExited)
        //                {
        //                    #region
        //                    //设置禁止弹出保存和覆盖的询问提示框   
        //                    lobjExcelApp.DisplayAlerts = false;
        //                    lobjExcelApp.AlertBeforeOverwriting = false;
        //                    //确保Excel进程关闭   
        //                    lobjExcelApp.Quit();
        //                    lobjExcelApp = null;
        //                    #endregion
        //                    app.Kill();//关闭进程  
        //                }
        //            }
        //        }
        //        if (lobjExcelWorkBook != null)
        //            lobjExcelWorkBook.Close(true, Type.Missing, Type.Missing);
        //        lobjExcelWorkBooks.Close();
        //        lobjExcelApp.Quit();
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(lobjExcelApp);
        //        // 安全回收进程
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //        System.GC.GetGeneration(lobjExcelApp);
        //    }

        //}
        #endregion

        /// <summary>
        /// Excel转PDF
        /// </summary>
        /// <param name="ExcelInputPath">读取Excel文件路径</param>
        /// <param name="PDFOutputPath">PDF输出路径</param>
        /// <param name="FileOutType">文件后缀</param>
        public static void ExcelToPDF(string ExcelInputPath, string PDFOutputPath, string FileOutType = null)
        {
            //可以破解附带的Aspose.Cells 16.12.0.0版本
            LicenseHelper.ModifyInMemory.ActivateMemoryPatching();
            //Create a new Workbook
            //Open an Excel file
            Workbook wb = new Workbook(ExcelInputPath);

            string PDFInputPath = FileOutType == null ? PDFOutputPath + Path.GetFileNameWithoutExtension(ExcelInputPath) + ".pdf" : PDFOutputPath + Path.GetFileNameWithoutExtension(ExcelInputPath)+ FileOutType;
            //Save the excel file to PDF format
            wb.Save(PDFInputPath, SaveFormat.Pdf);
        }
    }
}
