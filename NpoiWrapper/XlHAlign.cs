using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
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

    //----------------------------------------------------------------------------------------------
    //  XlHAlign in Interop.Excel is shown below...
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
}
