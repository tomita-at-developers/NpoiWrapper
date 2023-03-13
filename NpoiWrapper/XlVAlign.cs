using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// 文字配置(垂直方向)
    /// </summary>
    public enum XlVAlign
    {
        xlVAlignBottom = VerticalAlignment.Bottom,
        xlVAlignCenter = VerticalAlignment.Center,
        xlVAlignDistributed = VerticalAlignment.Distributed,
        xlVAlignJustify = VerticalAlignment.Justify,
        xlVAlignTop = VerticalAlignment.Top
    }

    //----------------------------------------------------------------------------------------------
    //  XlVAlign in Interop.Excel is shown below...
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
}
