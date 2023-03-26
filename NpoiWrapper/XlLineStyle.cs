using System;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlLineStyle in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum XlLineStyle
    //{
    //    xlContinuous = 1,
    //    xlDash = -4115,
    //    xlDashDot = 4,
    //    xlDashDotDot = 5,
    //    xlDot = -4118,
    //    xlDouble = -4119,
    //    xlSlantDashDot = 13,
    //    xlLineStyleNone = -4142
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum BorderStyle : short
    //{
    //    None,
    //    Thin,
    //    Medium,
    //    Dashed,
    //    Dotted,
    //    Thick,
    //    Double,
    //    Hair,
    //    MediumDashed,
    //    DashDot,
    //    MediumDashDot,
    //    DashDotDot,
    //    MediumDashDotDot,
    //    SlantedDashDot
    //}
    //------------------------------------------------------------------------------------------------------------------------------------------
    //  The relation of XlLineStye and BorderStyle is Shown below... 
    //------------------------------------------------------------------------------------------------------------------------------------------
    // XlLineStyle                          BorderStyle
    //------------------------------------------------------------------------------------------------------------------------------------------
    //                                      -- xlThin           -- xlMedium                 xlMedium                        xlThick
    //------------------------------------------------------------------------------------------------------------------------------------------
    // XlLineStyle.xlContinuous = 1         BorderStyle.Hair    BorderStyle.Thin,           BorderStyle.Medium,             BorderStyle.Thick
    // XlLineStyle.xlDash = -4115                               BorderStyle.Dashed,         BorderStyle.MediumDashed
    // XlLineStyle.xlDashDot = 4                                BorderStyle.DashDot,        BorderStyle.MediumDashDot
    // XlLineStyle.xlDashDotDot = 5                             BorderStyle.DashDotDot,     BorderStyle.MediumDashDotDot
    // XlLineStyle.xlDot = -4118                                BorderStyle.Dotted, 
    // XlLineStyle.xlDouble = -4119                             BorderStyle.Double
    // XlLineStyle.xlSlantDashDot = 13                          BorderStyle.SlantedDashDot
    // XlLineStyle.xlLineStyleNone = -4142                      BorderStyle.None
    //------------------------------------------------------------------------------------------------------------------------------------------

    /// <summary>
    /// 罫線の種類
    /// </summary>
    public enum XlLineStyle : int
    {
        xlContinuous = 1,
        xlDash = -4115,
        xlDashDot = 4,
        xlDashDotDot = 5,
        xlDot = -4118,
        xlDouble = -4119,
        xlSlantDashDot = 13,
        xlLineStyleNone = -4142
    }

    /// <summary>
    /// XlLineStyleの解析
    /// LineStyleは単純な一対一対応ではないので、XlsBorderStyle、PoiBorderStyleで処理している。
    /// ここでは上位パラメータ妥当性のチェックのみ実装している。
    /// </summary>
    internal static class XlLineStyleParser
    {
        /// <summary>
        /// 正当なXlLineStyleメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="XlValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out XlLineStyle XlValue)
        {
            bool RetVal = false;
            XlValue = (XlLineStyle)0;
            if (TryValue is XlLineStyle EnumValue)
            {
                XlValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(XlLineStyle), IntValue))
                {
                    XlValue = (XlLineStyle)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
    }
}
