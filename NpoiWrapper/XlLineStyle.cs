using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// 罫線の種類(EXCEL名のPOI値)
    /// </summary>
    public enum XlLineStyle : short
    {
        xlContinuous = BorderStyle.Thin,
        xlDash = BorderStyle.Dashed,
        xlDashDot = BorderStyle.DashDot,
        xlDashDotDot = BorderStyle.DashDotDot,
        xlDot = BorderStyle.Dotted,
        xlDouble = BorderStyle.Double,
        xlSlantDashDot = BorderStyle.SlantedDashDot,
        xlLineStyleNone = BorderStyle.None
    }
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
}
