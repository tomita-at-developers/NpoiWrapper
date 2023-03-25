using NPOI.SS.UserModel;
using System.Collections.Generic;
using System;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlHAlign in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlhalign
    //----------------------------------------------------------------------------------------------
    //public enum XlHAlign
    //{
    //    xlHAlignCenter = -4108,
    //    xlHAlignCenterAcrossSelection = 7,
    //    xlHAlignDistributed = -4117,
    //    xlHAlignFill = 5,
    //    xlHAlignGeneral = 1,
    //    xlHAlignJustify = -4130,
    //    xlHAlignLeft = -4131,
    //    xlHAlignRight = -4152
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum HorizontalAlignment
    //{
    //    General = 0,
    //    Left = 1,
    //    Center = 2,
    //    Right = 3,
    //    Justify = 5,
    //    Fill = 4,
    //    CenterSelection = 6,
    //    Distributed = 7
    //}

    /// <summary>
    /// 文字配置(水平方向)
    /// </summary>
    public enum XlHAlign
    {
        xlHAlignCenter = HorizontalAlignment.Center,
        xlHAlignCenterAcrossSelection = HorizontalAlignment.CenterSelection,
        xlHAlignDistributed = HorizontalAlignment.Distributed,
        xlHAlignFill = HorizontalAlignment.Fill,
        xlHAlignGeneral = HorizontalAlignment.General,
        xlHAlignJustify = HorizontalAlignment.Justify,
        xlHAlignLeft = HorizontalAlignment.Left,
        xlHAlignRight = HorizontalAlignment.Right
    }

    /// <summary>
    /// XlVAlignとVerticalAlignmentの相互変換
    /// </summary>
    internal static class XlHAlignParser
    {
        private static Dictionary<XlHAlign, HorizontalAlignment> _Map = new Dictionary<XlHAlign, HorizontalAlignment>()
        {
            { XlHAlign.xlHAlignCenter,                  HorizontalAlignment.Center          },
            { XlHAlign.xlHAlignCenterAcrossSelection,   HorizontalAlignment.CenterSelection },
            { XlHAlign.xlHAlignDistributed,             HorizontalAlignment.Distributed     },
            { XlHAlign.xlHAlignFill,                    HorizontalAlignment.Fill            },
            { XlHAlign.xlHAlignGeneral,                 HorizontalAlignment.General         },
            { XlHAlign.xlHAlignJustify,                 HorizontalAlignment.Justify         },
            { XlHAlign.xlHAlignLeft,                    HorizontalAlignment.Left            },
            { XlHAlign.xlHAlignRight,                   HorizontalAlignment.Right           }
        };
        /// <summary>
        /// XlVAlign値を指定してHorizontalAlignment値を取得。
        /// </summary>
        /// <param name="XlValue">XlVAlign値</param>
        /// <returns>HorizontalAlignment値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static HorizontalAlignment GetPoiValue(XlHAlign XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlValue.");
            }
        }
        /// <summary>
        /// HorizontalAlignment値を指定してXlHAlign値を取得。
        /// </summary>
        /// <param name="PoiValue">HorizontalAlignment値</param>
        /// <returns>XlHAlign値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlHAlign GetXlValue(HorizontalAlignment PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlHAlign, HorizontalAlignment> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of HorizontalAlignment.");
            }
        }
        /// <summary>
        /// 正当なXlAlignメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="XlValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out XlHAlign XlValue)
        {
            bool RetVal = false;
            XlValue = (XlHAlign)0;
            if (TryValue is XlHAlign EnumValue)
            {
                XlValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(XlHAlign), IntValue))
                {
                    XlValue = (XlHAlign)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
        /// <summary>
        /// 正当なHorizontalAlignmentメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="PoiValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out HorizontalAlignment PoiValue)
        {
            bool RetVal = false;
            PoiValue = (HorizontalAlignment)0;
            if (TryValue is HorizontalAlignment EnumValue)
            {
                PoiValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(HorizontalAlignment), IntValue))
                {
                    PoiValue = (HorizontalAlignment)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
    }
}
