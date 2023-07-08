//using System;
using Microsoft.Office.Interop.Excel;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static System.Func<Workbook, string, string, string, object[,]> GetAreaAs2dArray = (wb, strWs, strStart, strLast) =>
         (object[,])((Range)((Worksheet)wb.Sheets[strWs]).Range[strStart + ":" + strLast]).Value2;
   }
}