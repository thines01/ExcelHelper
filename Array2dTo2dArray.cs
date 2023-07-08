using System.Linq;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static object[][] Array2dTo2dArray(object[,] arr_int)
      {
         int intHeight = arr_int.GetUpperBound(0);
         int intWidth = arr_int.GetUpperBound(1);

         return
         (
            from i in Enumerable.Range(0, intHeight)
            select (arr_int.OfType<object>().ToList().GetRange(i * intWidth, intWidth)).ToArray()
         ).ToArray();
      }
   }
}