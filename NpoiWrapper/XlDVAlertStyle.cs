using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// データバリデーションエラーのスタイル
    /// </summary>
    public enum XlDVAlertStyle : int
    {
        xlValidAlertStop = NPOI.SS.UserModel.ERRORSTYLE.STOP,
        xlValidAlertWarning = NPOI.SS.UserModel.ERRORSTYLE.WARNING,
        xlValidAlertInformation = NPOI.SS.UserModel.ERRORSTYLE.INFO
    }
    //----------------------------------------------------------------------------------------------
    //  XlDVAlertStyle in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum XlDVAlertStyle
    //{
    //    xlValidAlertStop = 1,
    //    xlValidAlertWarning,
    //    xlValidAlertInformation
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public static class ERRORSTYLE
    //{
    //    public const int STOP = 0;
    //    public const int WARNING = 1;
    //    public const int INFO = 2;
    //}
}
