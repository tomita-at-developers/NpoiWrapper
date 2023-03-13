using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Styles.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util.Collections;
using System;
using System.Collections.Generic;



namespace Developers.NpoiWrapper
{
    // Interior interface in Interop.Excel is shown below...
    //public interface Interior
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    object Color { get; set; }
    //    object ColorIndex { get; set; }
    //    object InvertIfNegative { get; set; }
    //    object Pattern { get; set; }
    //    object PatternColor { get; set; }
    //    object PatternColorIndex { get; set; }
    //    object ThemeColor { get; set; }
    //    object TintAndShade { get; set; }
    //    object PatternThemeColor { get; set; }
    //    object PatternTintAndShade { get; set; }
    //    object Gradient { get; }
    //}

    public class Interior
    {
        /// <summary>
        /// 親ISheet
        /// </summary>
        private ISheet PoiSheet { get; set; }

        /// <summary>
        /// 絶対表現(RonwIndex,ColumnIndexとして直接利用可能)されたアドレスリスト
        /// </summary>
        private CellRangeAddressList SafeRangeAddressList { get; set; }

        /// <summary>
        /// RangeStyleクラス
        /// </summary>
        private RangeStyle _RangeStyle;
        private RangeStyle RangeStyle
        { 
            get
            {
                if (_RangeStyle == null) { _RangeStyle = new RangeStyle(PoiSheet, SafeRangeAddressList); }
                return _RangeStyle;
            }
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        public Interior(ISheet PoiSheet, CellRangeAddressList SafeAddressList)
        {
            this.PoiSheet = PoiSheet;
            this.SafeRangeAddressList = SafeAddressList;
        }

        /// <summary>
        /// セル内部の色(FillBackgroundColor)
        /// </summary>
        public object ColorIndex
        {
            get
            {
                return (short?)RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.ColorIndex));
            }
            set
            {
                if (value is short SafeValue)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.ColorIndex, SafeValue) } };
                    RangeStyle.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("Range.Interior.ColorIndex");
                }
            }
        }

        /// <summary>
        /// セル内部の模様(FillPattern)
        /// </summary>
        public object Pattern
        {
            get
            {
                return (XlPattern?)RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.Pattern));
            }
            set
            {
                if (value is XlPattern SafeValue)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.Pattern, (FillPattern)SafeValue) } };
                    RangeStyle.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("Range.Interior.Pattern");
                }
            }
        }

        /// <summary>
        /// 模様の色(FillForegroundColor)
        /// </summary>
        public object PatternColorIndex
        {
            get
            {
                return (short?)RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.PatternColorIndex));
            }
            set
            {
                if (value is short SafeValue)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.PatternColorIndex, SafeValue) } };
                    RangeStyle.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("Range.Interior.ColorIndex");
                }
            }
        }
    }
}
