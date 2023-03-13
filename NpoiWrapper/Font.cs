using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Styles.Properties;
using Developers.NpoiWrapper.Styles.Models;
using Developers.NpoiWrapper.Utils;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    /*
    --------------------------------------------------------------------------------------------   
    Font interface of Interop.Excel is shown below....
    --------------------------------------------------------------------------------------------
    public interface Font
    {
        Application Application { get; }                                - not implemented.
        XlCreator Creator       { get; }                                - not implemented.
        object Parent           { get; }                                - not implemented.
        object Background       { get; set; }   //XlBackground 
        object Bold             { get; set; }	//bool					IsBold
        object Color            { get; set; }	//double                ----------- later...
        object ColorIndex       { get; set; }	//short					Color
        object FontStyle        { get; set; }	//bool					- not implemented.
        object Italic           { get; set; }	//bool					IsItalic
        object Name             { get; set; }	//string				FontName
        object OutlineFont      { get; set; }	//bool					- not implemented. no effect on windows.
        object Shadow           { get; set; }	//bool					- not implemented. no effect on windows.
        object Size             { get; set; }	//double				FontHeight, FontHeightInPoints
        object Strikethrough    { get; set; }	//bool					IsStrikeout
        object Subscript        { get; set; }	//bool					FontSuperScript TypeOffset
        object Superscript      { get; set; }	//bool					FontSuperScript TypeOffset
        object Underline        { get; set; }	//XlUnderlineStyle		FontUnderlineType Underline
        object ThemeColor       { get; set; }	//int					- not implemented.
        object TintAndShade     { get; set; }	//single				- not implemented.
        XlThemeFont ThemeFont   { get; set; }	//XlThemeFont			- not implemented.
    }
    */

    /// <summary>
    /// FFontクラス
    /// Range.Fontとしての利用を想定
    /// </summary>
    public class Font
    {
        /// <summary>
        /// ISheetインスタンス
        /// </summary>
        private ISheet PoiSheet { get; set; }

        /// <summary>
        /// CellRangeAddressListインスタンス
        /// </summary>
        private CellRangeAddressList SafeRangeAddressList { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        public Font(ISheet PoiSheet, CellRangeAddressList SafeAddressList)
        {
            //親Range情報の保存
            this.PoiSheet = PoiSheet;
            this.SafeRangeAddressList = SafeAddressList;
        }

        public object Bold
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (bool?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Bold));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.Bold, SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Bold");
                }
            }
        }

        //object ColorIndex { get; set; }
        public object ColorIndex
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (short?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.ColorIndex));
            }
            set
            {
                if (value is short SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam >
                    {
                        new CellStyleParam(StyleName.Font.ColorIndex, SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.ColorIndex");
                }
            }
        }
        public object Italic
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                string PropName = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.IsItalic);
                return (bool?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Italic));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { 
                        new CellStyleParam(StyleName.Font.Italic, SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Italic");
                }
            }
        }
        public object Name
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (string)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Name));
            }
            set
            {
                if (value is string SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.Name, SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Name");
                }
            }
        }
        public object Size
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (double?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Size));
            }
            set
            {
                double SafeValue;
                if (value is double SafeDouble)
                {
                    SafeValue = SafeDouble;
                }
                else if (value is int SafeInt)
                {
                    SafeValue = (double)SafeInt;
                }
                else
                {
                    throw new ArgumentException("Font.Size");
                }
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                List<CellStyleParam> Params = new List<CellStyleParam>
                {
                    new CellStyleParam(StyleName.Font.Size, SafeValue)
                };
                StyleManger.UpdateProperties(Params);
            }
        }
        public object Strikethrough
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (bool?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Strikethrough));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.Strikethrough, SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Strikethrough");
                }
            }
        }
        public object Subscript
        { 
            get
            {
                object RetVal = null;
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                string PropName = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.TypeOffset);
                object CommonProp = StyleManger.GetCommonProperty(new CellStyleParam(PropName));
                if (CommonProp is FontSuperScript)
                {
                    if (CommonProp.Equals(FontSuperScript.Sub))
                    {
                        RetVal = true;
                    }
                    else
                    {
                        RetVal = false;
                    }
                }
                return RetVal;
            }
            set
            {
                if (value is bool SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    //変更初期値(true;FontSuperScript.Sub)
                    FontSuperScript PropValue = FontSuperScript.Sub;
                    //False指定の場合は現状によって判断
                    if (SafeValue == false)
                    {
                        object CommonProp = StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.TypeOffset));
                        //現状が特定できるとき
                        if (CommonProp is FontSuperScript)
                        {
                            //現状がFontSuperScript.SubならFalseはFontSuperScript.Noneとみなす
                            if (CommonProp.Equals(FontSuperScript.Sub))
                            {
                                PropValue = FontSuperScript.None;
                            }
                            //現状がFontSuperScript.SuperならFontSuperScript.Superとする
                            else if (CommonProp.Equals(FontSuperScript.Super))
                            {
                                PropValue = FontSuperScript.Super;
                            }
                            //現状がFontSuperScript.NoneならFontSuperScript.Noneとする
                            else
                            {
                                PropValue = FontSuperScript.None;
                            }
                        }
                        //現状が特定できないならFontSuperScript.Noneとみなす
                        else
                        {
                            PropValue = FontSuperScript.None;
                        }
                    }
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.TypeOffset, PropValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Subscript");
                }
            }
        }
        public object Superscript
        {
            get
            {
                object RetVal = DBNull.Value;
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                object CommonProp = StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.TypeOffset));
                if (CommonProp is FontSuperScript)
                {
                    if (CommonProp.Equals(FontSuperScript.Super))
                    {
                        RetVal = true;
                    }
                    else
                    {
                        RetVal = false;
                    }
                }
                return RetVal;
            }
            set
            {
                if (value is bool SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    //変更初期値(true:FontSuperScript.Super)
                    FontSuperScript PropValue = FontSuperScript.Super;
                    //False指定の場合は現状によって判断
                    if (SafeValue == false)
                    {
                        object CommonProp = StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.TypeOffset));
                        //現状が特定できるとき
                        if (CommonProp is FontSuperScript)
                        {
                            //現状がFontSuperScript.SubならFalseはFontSuperScript.Noneとみなす
                            if (CommonProp.Equals(FontSuperScript.Super))
                            {
                                PropValue = FontSuperScript.None;
                            }
                            //現状がFontSuperScript.SubrならFontSuperScript.Subとする
                            else if (CommonProp.Equals(FontSuperScript.Sub))
                            {
                                PropValue = FontSuperScript.Super;
                            }
                            //現状がFontSuperScript.NoneならFontSuperScript.Noneとする
                            else
                            {
                                PropValue = FontSuperScript.None;
                            }

                        }
                        //現状が特定できないならFontSuperScript.Noneとみなす
                        else
                        {
                            PropValue = FontSuperScript.None;
                        }
                    }
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.TypeOffset, PropValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Superscript");
                }
            }
        }
        public object Underline
        {
            get
            {
                RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                return (FontUnderlineType?)StyleManger.GetCommonProperty(new CellStyleParam(StyleName.Font.Underline));
            }
            set
            {
                if (value is XlUnderlineStyle SafeValue)
                {
                    RangeStyle StyleManger = new RangeStyle(this.PoiSheet, this.SafeRangeAddressList);
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    {
                        new CellStyleParam(StyleName.Font.Underline, (FontUnderlineType)SafeValue)
                    };
                    StyleManger.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentException("Font.Italic");
                }
            }
        }
    }
}
