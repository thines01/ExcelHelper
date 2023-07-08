using Microsoft.Office.Interop.Excel;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static string ActiveLastCell(this Application excel)
      {
         return excel.ActiveCell.SpecialCells(XlCellType.xlCellTypeLastCell).get_Address();
      }
   }
}