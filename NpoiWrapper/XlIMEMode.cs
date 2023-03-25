using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlIMEMode in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlimemode?view=excel-pia
    //----------------------------------------------------------------------------------------------
    //public enum XlIMEMode
    //{
    //    xlIMEModeNoControl,
    //    xlIMEModeOn,
    //    xlIMEModeOff,
    //    xlIMEModeDisable,
    //    xlIMEModeHiragana,
    //    xlIMEModeKatakana,
    //    xlIMEModeKatakanaHalf,
    //    xlIMEModeAlphaFull,
    //    xlIMEModeAlpha,
    //    xlIMEModeHangulFull,
    //    xlIMEModeHangul
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum ST_DataValidationImeMode
    //{
    //    noControl,
    //    off,
    //    on,
    //    disabled,
    //    hiragana,
    //    fullKatakana,
    //    halfKatakana,
    //    fullAlpha,
    //    halfAlpha,
    //    fullHangul,
    //    halfHangul
    //}

    /// <summary>
    /// IMEモード
    /// </summary>
    public enum XlIMEMode : int
    {
        xlIMEModeNoControl = 0,
        xlIMEModeOn = 1,
        xlIMEModeOff = 2,
        xlIMEModeDisable = 3,
        xlIMEModeHiragana = 4,
        xlIMEModeKatakana = 5,
        xlIMEModeKatakanaHalf = 6,
        xlIMEModeAlphaFull = 7,
        xlIMEModeAlpha = 8,
        xlIMEModeHangulFull = 9,
        xlIMEModeHangul = 10
    }
    /// <summary>
    /// XlIMETypeとST_DataValidationImeModeの相互変換
    /// </summary>
    internal static class XlIMEModeParser
    {
        private static Dictionary<XlIMEMode, ST_DataValidationImeMode> _Map = new Dictionary<XlIMEMode, ST_DataValidationImeMode>()
        {
            { XlIMEMode.xlIMEModeNoControl,     ST_DataValidationImeMode.noControl      },
            { XlIMEMode.xlIMEModeOn,            ST_DataValidationImeMode.on             },
            { XlIMEMode.xlIMEModeOff,           ST_DataValidationImeMode.off            },
            { XlIMEMode.xlIMEModeDisable,       ST_DataValidationImeMode.disabled       },
            { XlIMEMode.xlIMEModeHiragana,      ST_DataValidationImeMode.hiragana       },
            { XlIMEMode.xlIMEModeKatakana,      ST_DataValidationImeMode.fullKatakana   },
            { XlIMEMode.xlIMEModeKatakanaHalf,  ST_DataValidationImeMode.halfKatakana   },
            { XlIMEMode.xlIMEModeAlphaFull,     ST_DataValidationImeMode.fullAlpha      },
            { XlIMEMode.xlIMEModeAlpha,         ST_DataValidationImeMode.halfAlpha      },
            { XlIMEMode.xlIMEModeHangulFull,    ST_DataValidationImeMode.fullHangul     },
            { XlIMEMode.xlIMEModeHangul,        ST_DataValidationImeMode.halfHangul     }
        };
        /// <summary>
        /// XlIMEMode値を指定してST_DataValidationImeMode値を取得。
        /// </summary>
        /// <param name="XlValue">XlIMEMode値</param>
        /// <returns>ST_DataValidationImeMode値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static ST_DataValidationImeMode GetPoiValue(XlIMEMode XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlIMEMode.");
            }
        }
        /// <summary>
        /// ST_DataValidationImeMode値を指定してXlIMEMode値を取得。
        /// </summary>
        /// <param name="PoiValue">ST_DataValidationImeMode値</param>
        /// <returns>XlIMEMode値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlIMEMode GetXlValue(ST_DataValidationImeMode PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlIMEMode, ST_DataValidationImeMode> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of NPOI.OpenXmlFormats.Spreadsheet.ST_DataValidationImeMode.");
            }
        }
    }
}
