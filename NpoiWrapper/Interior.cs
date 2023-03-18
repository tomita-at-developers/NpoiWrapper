﻿using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Styles.Properties;
using Developers.NpoiWrapper.Styles.Utils;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util.Collections;
using SixLabors.ImageSharp.Metadata.Profiles.Iptc;
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
                return (short?)RangeStyle.GetCommonProperty(new CellStyleParam(StyleName.Interior.ColorIndex));
            }
            set
            {
                ColorPallet Pallet = new ColorPallet(value);
                if (Pallet.Index != null)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.ColorIndex, Pallet.Index) } };
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
                ColorPallet Pallet = new ColorPallet(value);
                if (Pallet.Index != null)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(StyleName.Interior.PatternColorIndex, Pallet.Index) } };
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
