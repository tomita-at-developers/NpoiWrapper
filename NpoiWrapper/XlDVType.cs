using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// データバリデーションの種類
    /// </summary>
    public enum XlDVType
    {
        xlValidateInputOnly,
        xlValidateWholeNumber,
        xlValidateDecimal,
        xlValidateList,
        xlValidateDate,
        xlValidateTime,
        xlValidateTextLength,
        xlValidateCustom
    }
    //----------------------------------------------------------------------------------------------
    //  XlDVType in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum XlDVType
    //{
    //    xlValidateInputOnly,
    //    xlValidateWholeNumber,
    //    xlValidateDecimal,
    //    xlValidateList,
    //    xlValidateDate,
    //    xlValidateTime,
    //    xlValidateTextLength,
    //    xlValidateCustom
    //}
}
