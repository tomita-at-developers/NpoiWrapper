using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// [名前を付けて保存]のアクセス モード
    /// </summary>
    public enum XlSaveAsAccessMode
    {
        xlExclusive = 3,
        xlNoChange = 1,
        xlShared = 2
    }
    //----------------------------------------------------------------------------------------------
    //XlSaveAsAccessMode  of Interop.Excel is shown below....
    //----------------------------------------------------------------------------------------------
    //public enum XlSaveAsAccessMode
    //{
    //    xlExclusive = 3,
    //    xlNoChange = 1,
    //    xlShared = 2
    //}
}
