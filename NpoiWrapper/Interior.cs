using Developers.NpoiWrapper.Model;
using Developers.NpoiWrapper.Model.Param;
using Developers.NpoiWrapper.Model.Utils;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Interior interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
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
        #region "fields"

        /// <summary>
        /// RangeStyleクラス
        /// </summary>
        private RangeStyle _RangeStyle;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        public Interior(Range ParentRange)
        {
            this.Parent = ParentRange;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// セル内部の色(FillBackgroundColor)
        /// </summary>
        public object ColorIndex
        {
            get
            {
                int? RetVal = null;
                object RawVal = RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.ColorIndex));
                if (ColorIndexParser.Xls.TryParse(RawVal, out int XlValue))
                {
                    RetVal = XlValue;
                }
                return RetVal;
            }
            set
            {
                if (ColorIndexParser.Poi.TryParse(value, out short PoiValue))
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.ColorIndex, PoiValue) } };
                    //実体のある色が指定された場合
                    if (PoiValue != IndexedColors.Automatic.Index)
                    {
                        //FillPatternがNoFillの場合は自動的にSolidForegroundを設定する
                        object Pattern = RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.Pattern));
                        if((FillPattern)Pattern == FillPattern.NoFill)
                        {
                            Params.Add(new CellStyleParam(StyleName.Interior.Pattern, FillPattern.SolidForeground));
                        }
                    }
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
                XlPattern? RetVal = null;
                object RawVal = RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.Pattern));
                if (XlPatternParser.Try(RawVal, out FillPattern PoiValue))
                {
                    RetVal = XlPatternParser.GetXlValue(PoiValue);
                }
                return RetVal;
            }
            set
            {
                if (XlPatternParser.Try(value, out XlPattern XlValue))
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.Pattern, XlPatternParser.GetPoiValue(XlValue)) } };
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
                int? RetVal = null;
                object RawVal = RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.PatternColorIndex));
                if (ColorIndexParser.Xls.TryParse(RawVal, out int XlValue))
                {
                    RetVal = XlValue;
                }
                return RetVal;
            }
            set
            {
                if (ColorIndexParser.Poi.TryParse(value, out short PoiValue))
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.PatternColorIndex, PoiValue) } };
                    RangeStyle.UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("Range.Interior.PatternColorIndex");
                }
            }
        }

        #endregion

        #region "private properties"

        /// <summary>
        /// RangeStyleクラスインスタンス
        /// </summary>
        private RangeStyle RangeStyle
        {
            get
            {
                //最初にアクセスしたときにコンストラクト
                if (_RangeStyle == null) { _RangeStyle = new RangeStyle(this.Parent); }
                return _RangeStyle;
            }
        }

        #endregion

        #endregion
    }
}
