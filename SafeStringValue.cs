//using System;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static string SafeStringValue(Range row, int intCol)
      {
         object obj = ((Range)row.Cells[Missing.Value, intCol]).Value2;
         return (null == obj) ? "" : obj.ToString();
      }

      public static string SafeStringValue(object cell, int v)
      {
         return (null == (Range)cell) ? "" : ((Range)cell).Value2.ToString();
      }
   }
}