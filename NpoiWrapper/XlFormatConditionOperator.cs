using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// 評価条件。ValidationのOperationに指定できる。
    /// </summary>
    public enum XlFormatConditionOperator : int
    {
        xlBetween = 1,
        xlNotBetween = 2,
        xlEqual = 3,
        xlNotEqual = 4,
        xlGreater = 5,
        xlLess = 6,
        xlGreaterEqual = 7,
        xlLessEqual = 8
    }
    //----------------------------------------------------------------------------------------------
    //  XlFormatConditionOperator in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum XlFormatConditionOperator
    //{
    //    xlBetween = 1,
    //    xlNotBetween,
    //    xlEqual,
    //    xlNotEqual,
    //    xlGreater,
    //    xlLess,
    //    xlGreaterEqual,
    //    xlLessEqual
    //}
}

