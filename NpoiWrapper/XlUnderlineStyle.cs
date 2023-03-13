using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// フォント下線のスタイル
    /// </summary>
    public enum XlUnderlineStyle : byte
    {
        xlUnderlineStyleDouble = FontUnderlineType.Double,
        xlUnderlineStyleDoubleAccounting = FontUnderlineType.DoubleAccounting,
        xlUnderlineStyleNone = FontUnderlineType.None,
        xlUnderlineStyleSingle = FontUnderlineType.Single,
        xlUnderlineStyleSingleAccounting = FontUnderlineType.SingleAccounting
    }
    //----------------------------------------------------------------------------------------------
    //XlUnderlineStyle  of Interop.Excel is shown below....
    //----------------------------------------------------------------------------------------------
    //public enum XlUnderlineStyle
    //{
    //    xlUnderlineStyleDouble = -4119,
    //    xlUnderlineStyleDoubleAccounting = 5,
    //    xlUnderlineStyleNone = -4142,
    //    xlUnderlineStyleSingle = 2,
    //    xlUnderlineStyleSingleAccounting = 4
    //}
    //----------------------------------------------------------------------------------------------
    //Corresponding definition in NPOI is shown below....
    //----------------------------------------------------------------------------------------------
    //public enum FontUnderlineType : byte
    //{
    //    None = 0,
    //    Single = 1,
    //    Double = 2,
    //    SingleAccounting = 33,
    //    DoubleAccounting = 34
    //}
}