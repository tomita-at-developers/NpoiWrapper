using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlColorIndex of Interop.Excel is shown below....
    //----------------------------------------------------------------------------------------------
    //public enum XlColorIndex
    //{
    //    xlColorIndexAutomatic = -4105,
    //    xlColorIndexNone = -4142
    //}

    /// <summary>
    /// カラーパレット上のインデックス
    /// </summary>
    public enum XlColorIndex : int
    {
        xlColorIndexAutomatic = -4105,
        xlColorIndexNone = -4142
    }

    internal static class ColorIndexParser
    {
        internal static class Poi
        {
            /// <summary>
            /// 指定されたExcelのColorIndexがPoiのIndexedColorに変換できるか確認する
            /// </summary>
            /// <param name="TryValue">確認対象</param>
            /// <param name="PoiValue">確認できたメンバー</param>
            /// <returns>確認結果</returns>
            public static bool TryParse(object TryValue, out short PoiValue)
            {
                bool RetVal = false;
                PoiValue = IndexedColors.Automatic.Index;
                try
                {
                    int RawValue = Convert.ToInt32(TryValue);
                    //XlColorIndexにある値
                    if (Enum.IsDefined(typeof(XlColorIndex), RawValue))
                    {
                        //ExcelにはAutomaticとNoneの二つがあるがPOIにはAutomaticしかない。
                        //(POIのAutomaticの実装は恐らくExcelのNoneなのではないか､､､､)
                        PoiValue = IndexedColors.Automatic.Index;
                        RetVal = true;
                    }
                    //パレット範囲内であれば対応するIndexedColorに変換
                    else if (1 <= RawValue && RawValue <= 56)
                    {
                        PoiValue = (short)(RawValue + 7);
                        RetVal = true;
                    }
                    //その他はエラー
                    else
                    {
                        RetVal = false;
                    }
                }
                catch
                {
                    RetVal = false;
                }
                return RetVal;
            }
        }
        internal static class Xls
        {
            /// <summary>
            /// 指定されたPoiのIndexedColorがExcelのColorIndexに変換できる確認する
            /// </summary>
            /// <param name="TryValue">確認対象</param>
            /// <param name="XlValue">変換後のExcel値</param>
            /// <returns>確認結果</returns>
            public static bool TryParse(object TryValue, out int XlValue)
            {
                bool RetVal = false;
                XlValue = (int)XlColorIndex.xlColorIndexAutomatic;
                try
                {
                    int RawValue = Convert.ToInt16(TryValue);
                    //実験した結果、63以下、または66から81までは７ひいた値になっている
                    if (RawValue <= 63 || (66 <= RawValue && RawValue <= 81))
                    {
                        XlValue = (int)(RawValue - 7);
                    }
                    //上記以外はAutoになる
                    else
                    {
                        XlValue = (int)XlColorIndex.xlColorIndexAutomatic;
                    }
                    RetVal = true;
                }
                catch
                {
                    RetVal = false;
                }
                return RetVal;
            }
        }
    }
}