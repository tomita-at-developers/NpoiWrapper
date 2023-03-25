using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlSaveAsAccessMode  of Interop.Excel is shown below....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlsaveasaccessmode
    //----------------------------------------------------------------------------------------------
    //public enum XlSaveAsAccessMode
    //{
    //    xlExclusive = 3,
    //    xlNoChange = 1,
    //    xlShared = 2
    //}

    /// <summary>
    /// [名前を付けて保存]のアクセス モード
    /// </summary>
    public enum XlSaveAsAccessMode
    {
        xlExclusive = 3,
        xlNoChange = 1,
        xlShared = 2
    }
}
