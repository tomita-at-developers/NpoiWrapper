using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlWBATemplate  of Interop.Excel is shown below....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlwbatemplate
    //----------------------------------------------------------------------------------------------
    //public enum XlWBATemplate
    //{
    //    xlWBATChart = -4109,
    //    xlWBATExcel4IntlMacroSheet = 4,
    //    xlWBATExcel4MacroSheet = 3,
    //    xlWBATWorksheet = -4167
    //}

    /// <summary>
    /// ファイルテンプレート
    /// </summary>
    public enum XlWBATemplate
    {
        xlWBATChart = -4109,
        xlWBATExcel4IntlMacroSheet = 4,
        xlWBATExcel4MacroSheet = 3,
        xlWBATWorksheet = -4167
    }
}
