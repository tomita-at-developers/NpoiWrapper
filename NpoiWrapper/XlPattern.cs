using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// セルの塗りつぶしパターン
    /// Interop.Excel.XlPatternとNPOI.SS.UserModel.FillPatternとで、見た目の似たものを対応付けた。
    /// ただし４つの未サポート値がある。
    /// </summary>
    public enum XlPattern : short
    {
        //xlPatternAutomatic = -4105,
        xlPatternChecker = FillPattern.BigSpots,
        xlPatternCrissCross = FillPattern.Diamonds,
        xlPatternDown = FillPattern.ThickBackwardDiagonals,
        xlPatternGray16 = FillPattern.AltBars,
        xlPatternGray25 = FillPattern.SparseDots,
        xlPatternGray50 = FillPattern.FineDots,
        //xlPatternGray75 = -4126,
        xlPatternGray8 = FillPattern.LeastDots,
        xlPatternGrid = FillPattern.Squares,
        xlPatternHorizontal = FillPattern.ThickHorizontalBands,
        xlPatternLightDown = FillPattern.ThinBackwardDiagonals,
        xlPatternLightHorizontal = FillPattern.ThinHorizontalBands,
        xlPatternLightUp = FillPattern.ThinForwardDiagonals,
        xlPatternLightVertical = FillPattern.ThinVerticalBands,
        xlPatternNone = FillPattern.NoFill,
        xlPatternSemiGray75 = FillPattern.Bricks,
        xlPatternSolid = FillPattern.SolidForeground,
        xlPatternUp = FillPattern.ThickForwardDiagonals,
        xlPatternVertical = FillPattern.ThickVerticalBands
        //xlPatternLinearGradient = 4000,
        //xlPatternRectangularGradient = 4001
    }

    //----------------------------------------------------------------------------------------------
    //XlPattern  of Interop.Excel is shown below....     22 patterns.
    //----------------------------------------------------------------------------------------------
    //public enum XlPattern
    //{
    //    xlPatternAutomatic = -4105,           *** NpoiWrapperでは未サポート
    //    xlPatternChecker = 9,
    //    xlPatternCrissCross = 0x10,
    //    xlPatternDown = -4121,
    //    xlPatternGray16 = 17,
    //    xlPatternGray25 = -4124,
    //    xlPatternGray50 = -4125,
    //    xlPatternGray75 = -4126,              *** NpoiWrapperでは未サポート
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
    //    xlPatternLinearGradient = 4000,       *** NpoiWrapperでは未サポート
    //    xlPatternRectangularGradient = 4001   *** NpoiWrapperでは未サポート
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


}
