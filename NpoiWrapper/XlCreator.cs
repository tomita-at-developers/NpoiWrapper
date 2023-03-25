using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlCreator  of Interop.Excel is shown below.....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlcreator
    //  in hex 58 43 45 4C  = XCEL
    //----------------------------------------------------------------------------------------------
    //public enum XlCreator
    //{
    //    xlCreatorCode = 1480803660
    //}

    /// <summary>
    /// ダミークリエイター
    /// </summary>
    public enum XlCreator : int
    {
        xlCreatorCode = 0

    }
}
