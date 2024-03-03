using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlPageBreak in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlpagebreak
    //----------------------------------------------------------------------------------------------
    //public enum XlPageBreak
    //{
    //    //
    //    // 概要:
    //    //     Excel will automatically add page breaks.
    //    xlPageBreakAutomatic = -4105,
    //    //
    //    // 概要:
    //    //     Page breaks are manually inserted.
    //    xlPageBreakManual = -4135,
    //    //
    //    // 概要:
    //    //     Page breaks are not inserted in the worksheet.
    //    xlPageBreakNone = -4142
    //}
    public enum XlPageBreak
    {
        //
        // 概要:
        //     Excel will automatically add page breaks.
        xlPageBreakAutomatic = -4105,
        //
        // 概要:
        //     Page breaks are manually inserted.
        xlPageBreakManual = -4135,
        //
        // 概要:
        //     Page breaks are not inserted in the worksheet.
        xlPageBreakNone = -4142
    }
}
