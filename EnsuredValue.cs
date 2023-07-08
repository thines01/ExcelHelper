//using System;
using Microsoft.Office.Interop.Excel;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static System.Func<Range, string> EnsuredValue = (cellCurrent) =>
         SafeStringValue(cellCurrent, 1);
   }
}