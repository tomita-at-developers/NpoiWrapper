using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlVAlign in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlvalign
    //----------------------------------------------------------------------------------------------
    //public enum XlVAlign
    //{
    //    xlVAlignBottom = -4107,
    //    xlVAlignCenter = -4108,
    //    xlVAlignDistributed = -4117,
    //    xlVAlignJustify = -4130,
    //    xlVAlignTop = -4160
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum VerticalAlignment
    //{
    //    None = -1,
    //    Top,
    //    Center,
    //    Bottom,
    //    Justify,
    //    Distributed
    //}

    /// <summary>
    /// 文字配置(垂直方向)
    /// </summary>
    public enum XlVAlign : int
    {
        xlVAlignBottom = -4107,
        xlVAlignCenter = -4108,
        xlVAlignDistributed = -4117,
        xlVAlignJustify = -4130,
        xlVAlignTop = -4160
    }

    /// <summary>
    /// XlVAlignとVerticalAlignmentの相互変換
    /// </summary>
    internal static class XlVAlignParser
    {
        private static readonly Dictionary<XlVAlign, VerticalAlignment> _Map = new Dictionary<XlVAlign, VerticalAlignment>()
        {
            { XlVAlign.xlVAlignBottom,      VerticalAlignment.Bottom        },
            { XlVAlign.xlVAlignCenter,      VerticalAlignment.Center        },
            { XlVAlign.xlVAlignDistributed, VerticalAlignment.Distributed   },
            { XlVAlign.xlVAlignJustify,     VerticalAlignment.Justify       },
            { XlVAlign.xlVAlignTop,         VerticalAlignment.Top           }
        };
        /// <summary>
        /// XlVAlign値を指定してVerticalAlignment値を取得。
        /// </summary>
        /// <param name="XlValue">XlVAlign</param>
        /// <returns>VerticalAlignment値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static VerticalAlignment GetPoiValue(XlVAlign XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlVAlign.");
            }
        }
        /// <summary>
        /// VerticalAlignment値を指定してXlVAlign値を取得。
        /// </summary>
        /// <param name="PoiValue">VerticalAlignment値</param>
        /// <returns>XlVAlign値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlVAlign GetXlValue(VerticalAlignment PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlVAlign, VerticalAlignment> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of VerticalAlignment.");
            }
        }
        /// <summary>
        /// 正当なXlAlignメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="XlValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out XlVAlign XlValue)
        {
            bool RetVal = false;
            XlValue = (XlVAlign)0;
            if (TryValue is XlVAlign EnumValue)
            {
                XlValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(XlVAlign), IntValue))
                {
                    XlValue = (XlVAlign)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
        /// <summary>
        /// 正当なVerticalAlignmentメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="PoiValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out VerticalAlignment PoiValue)
        {
            bool RetVal = false;
            PoiValue = (VerticalAlignment)0;
            if (TryValue is VerticalAlignment EnumValue)
            {
                PoiValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(VerticalAlignment), IntValue))
                {
                    PoiValue = (VerticalAlignment)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
    }
}
