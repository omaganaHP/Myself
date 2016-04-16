using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HP.SDLCToolsBO.Model
{
    public abstract class ExcelHandler
    {
        
        public abstract string ExportToExcel(List<string> IdListsToExport, List<string> ColumnsNames,char separationChar);
        public abstract int CountListItems(String fileName);
    }
}
