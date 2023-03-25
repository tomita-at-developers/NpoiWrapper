using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlDVAlertStyle in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xldvalertstyle
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

    /// <summary>
    /// データバリデーションエラーのスタイル
    /// </summary>
    public enum XlDVAlertStyle : int
    {
        xlValidAlertStop = 1,
        xlValidAlertWarning = 2,
        xlValidAlertInformation = 3
    }
    /// <summary>
    /// XlDVAlertStyleとNPOI.SS.UserModel.ERRORSTYLEの相互変換
    /// </summary>
    internal static class XlDVAlertStyleParser
    {
        private static Dictionary<XlDVAlertStyle, int> _Map = new Dictionary<XlDVAlertStyle, int>()
        {
            { XlDVAlertStyle.xlValidAlertStop,          NPOI.SS.UserModel.ERRORSTYLE.STOP       },
            { XlDVAlertStyle.xlValidAlertWarning,        NPOI.SS.UserModel.ERRORSTYLE.WARNING    },
            { XlDVAlertStyle.xlValidAlertInformation,    NPOI.SS.UserModel.ERRORSTYLE.INFO       }
        };
        /// <summary>
        /// XlDVAlertStyle値を指定してERRORSTYLE値を取得。
        /// </summary>
        /// <param name="XlValue">XlDVAlertStyle値</param>
        /// <returns>ERRORSTYLE値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static int GetPoiValue(XlDVAlertStyle XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlDVAlertStyle.");
            }
        }
        /// <summary>
        /// ERRORSTYLE値を指定してXlDVAlertStyle値を取得。
        /// </summary>
        /// <param name="PoiValue">ERRORSTYLE値</param>
        /// <returns>XlDVAlertStyle値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlDVAlertStyle GetXlValue(int PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlDVAlertStyle, int> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of NPOI.SS.UserModel.ERRORSTYLE.");
            }
        }
    }
}
