using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper.Styles.Properties
{
    internal class XlsBorderStyle
    {
        //----------------------------------------------------------------------------------------------------------------------------------------
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
        public XlLineStyle LineStyle { get; private set; } = XlLineStyle.xlLineStyleNone;
        public XlBorderWeight Weight { get; private set; } = XlBorderWeight.xlThin;

        public XlsBorderStyle(BorderStyle PoiBorderStyle)
        {
            LineStyle = XlLineStyle.xlLineStyleNone;
            Weight = XlBorderWeight.xlThin;
            switch (PoiBorderStyle)
            {
                case BorderStyle.Hair:
                    LineStyle = XlLineStyle.xlContinuous;
                    Weight = XlBorderWeight.xlHairline;
                    break;
                case BorderStyle.Thin:
                    LineStyle = XlLineStyle.xlContinuous;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.Medium:
                    LineStyle = XlLineStyle.xlContinuous;
                    Weight = XlBorderWeight.xlMedium;
                    break;
                case BorderStyle.Thick:
                    LineStyle = XlLineStyle.xlContinuous;
                    Weight = XlBorderWeight.xlThick;
                    break;
                case BorderStyle.Dashed:
                    LineStyle = XlLineStyle.xlDash;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.MediumDashed:
                    LineStyle = XlLineStyle.xlDash;
                    Weight = XlBorderWeight.xlMedium;
                    break;
                case BorderStyle.DashDot:
                    LineStyle = XlLineStyle.xlDashDot;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.MediumDashDot:
                    LineStyle = XlLineStyle.xlDashDot;
                    Weight = XlBorderWeight.xlMedium;
                    break;
                case BorderStyle.DashDotDot:
                    LineStyle = XlLineStyle.xlDashDotDot;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.MediumDashDotDot:
                    LineStyle = XlLineStyle.xlDashDotDot;
                    Weight = XlBorderWeight.xlMedium;
                    break;
                case BorderStyle.Dotted:
                    LineStyle = XlLineStyle.xlDot;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.Double:
                    LineStyle = XlLineStyle.xlDouble;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.SlantedDashDot:
                    LineStyle = XlLineStyle.xlSlantDashDot;
                    Weight = XlBorderWeight.xlThin;
                    break;
                case BorderStyle.None:
                    LineStyle = XlLineStyle.xlLineStyleNone;
                    Weight = XlBorderWeight.xlThin;
                    break;
                default:
                    LineStyle = XlLineStyle.xlLineStyleNone;
                    Weight = XlBorderWeight.xlThin;
                    break;
            }
        }
    }
}
