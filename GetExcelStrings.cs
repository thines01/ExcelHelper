using System.Collections.Generic;
using System.Linq;
//
namespace ExcelHelper
{
   public static partial class CExcelHelper
   {
      public static IEnumerable<string> GetExcelStrings()
      {
         return
         (
            from c1 in arr_strAlphaPlus
            from c2 in arr_strAlphaPlus
            where c1 == string.Empty || c2 != string.Empty // magic
            select c1 + c2
         ).Skip(1).Take(256);
      }
   }
}
