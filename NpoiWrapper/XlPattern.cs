using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlPattern  of Interop.Excel is shown below....     22 patterns.
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlpattern
    //----------------------------------------------------------------------------------------------
    //public enum XlPattern
    //{
    //    xlPatternAutomatic = -4105,
    //    xlPatternChecker = 9,
    //    xlPatternCrissCross = 0x10,
    //    xlPatternDown = -4121,
    //    xlPatternGray16 = 17,
    //    xlPatternGray25 = -4124,
    //    xlPatternGray50 = -4125,
    //    xlPatternGray75 = -4126,
    //    xlPatternGray8 = 18,
    //    xlPatternGrid = 0xF,
    //    xlPatternHorizontal = -4128,
    //    xlPatternLightDown = 13,
    //    xlPatternLightHorizontal = 11,
    //    xlPatternLightUp = 14,
    //    xlPatternLightVertical = 12,
    //    xlPatternNone = -4142,
    //    xlPatternSemiGray75 = 10,
    //    xlPatternSolid = 1,
    //    xlPatternUp = -4162,
    //    xlPatternVertical = -4166,
    //    xlPatternLinearGradient = 4000,
    //    xlPatternRectangularGradient = 4001
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...     19 patterns.
    //----------------------------------------------------------------------------------------------
    //public enum FillPattern : short
    //{
    //    NoFill,
    //    SolidForeground,
    //    FineDots,
    //    AltBars,
    //    SparseDots,
    //    ThickHorizontalBands,
    //    ThickVerticalBands,
    //    ThickBackwardDiagonals,
    //    ThickForwardDiagonals,
    //    BigSpots,
    //    Bricks,
    //    ThinHorizontalBands,
    //    ThinVerticalBands,
    //    ThinBackwardDiagonals,
    //    ThinForwardDiagonals,
    //    Squares,
    //    Diamonds,
    //    LessDots,
    //    LeastDots
    //}

    /// <summary>
    /// セルの塗りつぶしパターン
    /// Interop.Excel.XlPatternとNPOI.SS.UserModel.FillPatternとで、見た目の似たものを対応付けた。
    /// ただし４つの未サポート値がある。
    /// </summary>
    public enum XlPattern : short
    {
        xlPatternAutomatic = -4105,         // NpoiWrapperでは未サポート(xlPatternSolidとみなす)
        xlPatternChecker = 9,
        xlPatternCrissCross = 0x10,
        xlPatternDown = -4121,
        xlPatternGray16 = 17,
        xlPatternGray25 = -4124,
        xlPatternGray50 = -4125,
        xlPatternGray75 = -4126,            // NpoiWrapperでは未サポート(xlPatternSolidとみなす)
        xlPatternGray8 = 18,
        xlPatternGrid = 0xF,
        xlPatternHorizontal = -4128,
        xlPatternLightDown = 13,
        xlPatternLightHorizontal = 11,
        xlPatternLightUp = 14,
        xlPatternLightVertical = 12,
        xlPatternNone = -4142,
        xlPatternSemiGray75 = 10,
        xlPatternSolid = 1,
        xlPatternUp = -4162,
        xlPatternVertical = -4166,
        xlPatternLinearGradient = 4000,     // NpoiWrapperでは未サポート(xlPatternSolidとみなす)
        xlPatternRectangularGradient = 4001 // NpoiWrapperでは未サポート(xlPatternSolidとみなす)
    }

    /// <summary>
    /// XlPatternとFillPatternの相互変換
    /// </summary>
    internal static class XlPatternParser
    {
        private static readonly Dictionary<XlPattern, FillPattern> _Map = new Dictionary<XlPattern, FillPattern>()
        {
            { XlPattern.xlPatternAutomatic,             FillPattern.SolidForeground         },  //未サポート
            { XlPattern.xlPatternChecker,               FillPattern.BigSpots                },
            { XlPattern.xlPatternCrissCross,            FillPattern.Diamonds                },
            { XlPattern.xlPatternDown,                  FillPattern.ThickBackwardDiagonals  },
            { XlPattern.xlPatternGray16,                FillPattern.AltBars                 },
            { XlPattern.xlPatternGray25,                FillPattern.SparseDots              },
            { XlPattern.xlPatternGray50,                FillPattern.FineDots                },
            { XlPattern.xlPatternGray75,                FillPattern.SolidForeground         },  //未サポート
            { XlPattern.xlPatternGray8,                 FillPattern.LeastDots               },
            { XlPattern.xlPatternGrid,                  FillPattern.Squares                 },
            { XlPattern.xlPatternHorizontal,            FillPattern.ThickHorizontalBands    },
            { XlPattern.xlPatternLightDown,             FillPattern.ThinBackwardDiagonals   },
            { XlPattern.xlPatternLightHorizontal,       FillPattern.ThinHorizontalBands     },
            { XlPattern.xlPatternLightUp,               FillPattern.ThinForwardDiagonals    },
            { XlPattern.xlPatternLightVertical,         FillPattern.ThinVerticalBands       },
            { XlPattern.xlPatternNone,                  FillPattern.NoFill                  },
            { XlPattern.xlPatternSemiGray75,            FillPattern.Bricks                  },
            { XlPattern.xlPatternSolid,                 FillPattern.SolidForeground         },
            { XlPattern.xlPatternUp,                    FillPattern.ThickForwardDiagonals   },
            { XlPattern.xlPatternVertical,              FillPattern.ThickVerticalBands      },
            { XlPattern.xlPatternLinearGradient,        FillPattern.SolidForeground         },  //未サポート
            { XlPattern.xlPatternRectangularGradient,   FillPattern.SolidForeground         }   //未サポート
        };
        /// <summary>
        /// XlPattern値を指定してFillPattern値を取得。
        /// </summary>
        /// <param name="XlValue">XlPattern値</param>
        /// <returns>FillPattern値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static FillPattern GetPoiValue(XlPattern XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlPattern.");
            }
        }
        /// <summary>
        /// FillPattern値を指定してXlPattern値を取得。
        /// </summary>
        /// <param name="PoiValue">FillPattern値</param>
        /// <returns>XlHAlign値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlPattern GetXlValue(FillPattern PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlPattern, FillPattern> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of FillPattern.");
            }
        }
        /// <summary>
        /// 正当なXlPatternメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="XlValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out XlPattern XlValue)
        {
            bool RetVal = false;
            XlValue = (XlPattern)0;
            if (TryValue is XlPattern EnumValue)
            {
                XlValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(XlPattern), IntValue))
                {
                    XlValue = (XlPattern)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
        /// <summary>
        /// 正当なFillPatternメンバーかどうか確認する
        /// </summary>
        /// <param name="TryValue">確認対象</param>
        /// <param name="PoiValue">確認できたメンバー</param>
        /// <returns>確認結果</returns>
        public static bool Try(object TryValue, out FillPattern PoiValue)
        {
            bool RetVal = false;
            PoiValue = (FillPattern)0;
            if (TryValue is FillPattern EnumValue)
            {
                PoiValue = EnumValue;
                RetVal = true;
            }
            else if (TryValue is int IntValue)
            {
                if (Enum.IsDefined(typeof(FillPattern), IntValue))
                {
                    PoiValue = (FillPattern)IntValue;
                    RetVal = true;
                }
            }
            return RetVal;
        }
    }


}
