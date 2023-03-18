﻿using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper.Styles.Properties
{
    /// <summary>
    /// Poi名前のBorderStyleを取得する(ExcelのXlLineStyle, XlBorderWeightから)
    /// </summary>
    internal class PoiBorderStyle
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
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="LineStyle">XlLineStyle</param>
        /// <param name="Weight">XlBorderWeight</param>
        public PoiBorderStyle(XlLineStyle LineStyle, XlBorderWeight Weight)
        {
            switch (LineStyle)
            {
                case XlLineStyle.xlContinuous:
                    switch (Weight)
                    {
                        case XlBorderWeight.xlHairline:
                            BorderStyle = BorderStyle.Hair;
                            break;
                        case XlBorderWeight.xlThin:
                            BorderStyle = BorderStyle.Thin;
                            break;
                        case XlBorderWeight.xlMedium:
                            BorderStyle = BorderStyle.Medium;
                            break;
                        case XlBorderWeight.xlThick:
                            BorderStyle = BorderStyle.Thick;
                            break;
                        default:
                            BorderStyle = BorderStyle.Thin;
                            break;
                    }
                    break;
                case XlLineStyle.xlDash:
                    switch (Weight)
                    {
                        case XlBorderWeight.xlMedium:
                        case XlBorderWeight.xlThick:
                            BorderStyle = BorderStyle.MediumDashed;
                            break;
                        default:
                            BorderStyle = BorderStyle.Dashed;
                            break;
                    }
                    break;
                case XlLineStyle.xlDashDot:
                    switch (Weight)
                    {
                        case XlBorderWeight.xlMedium:
                        case XlBorderWeight.xlThick:
                            BorderStyle = BorderStyle.MediumDashDot;
                            break;
                        default:
                            BorderStyle = BorderStyle.DashDot;
                            break;
                    }
                    break;
                case XlLineStyle.xlDashDotDot:
                    switch (Weight)
                    {
                        case XlBorderWeight.xlMedium:
                        case XlBorderWeight.xlThick:
                            BorderStyle = BorderStyle.MediumDashDotDot;
                            break;
                        default:
                            BorderStyle = BorderStyle.DashDotDot;
                            break;
                    }
                    break;
                case XlLineStyle.xlDot:
                    BorderStyle = BorderStyle.Dotted;
                    break;
                case XlLineStyle.xlDouble:
                    BorderStyle = BorderStyle.Double;
                    break;
                case XlLineStyle.xlSlantDashDot:
                    BorderStyle = BorderStyle.SlantedDashDot;
                    break;
                case XlLineStyle.xlLineStyleNone:
                    BorderStyle = BorderStyle.None;
                    break;
                default:
                    BorderStyle = BorderStyle.None;
                    break;
            }
        }

        #endregion

        #region "properties"

        /// <summary>
        /// Poi名前のBorderStyle
        /// </summary>
        public BorderStyle BorderStyle { get; private set; } = BorderStyle.None;

        #endregion
    }
}
