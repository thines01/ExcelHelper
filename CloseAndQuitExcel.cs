using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static void CloseAndQuit(this _Application excel, Workbook wb)
      {
         wb.Close(XlSaveAction.xlSaveChanges, Missing.Value, Missing.Value);
         excel.Workbooks.Close();
         excel.Quit();
         Marshal.ReleaseComObject(wb);
         wb = null;
         Marshal.ReleaseComObject(excel);
         excel = null;
      }
   }
}