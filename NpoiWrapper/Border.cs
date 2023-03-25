using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Styles.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Border interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Border
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    object Color { get; set; }
    //    object ColorIndex { get; set; }
    //    object LineStyle { get; set; }
    //    object Weight { get; set; }
    //    object ThemeColor { get; set; }
    //    object TintAndShade { get; set; }
    //}

    /// <summary>
    /// Borderクラス
    /// Microsoft.Office.Interop.Excel.Borderをエミュレート
    /// </summary>
    public class Border
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRanget">Rangeインスタンス</param>
        /// <param name="BordersIndex">XlBordersIndex値</param>
        internal Border(Range ParentRanget, XlBordersIndex? BordersIndex)
        {
            //親Range情報の保存
            this.Parent = ParentRanget;
            //Border情報の保存
            this.BordersIndex = BordersIndex;
            //RangeBorderStyle生成
            BorderStyle = new RangeBorderStyle(this.PoiSheet, this.SafeAddressList, this.BordersIndex);
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// 罫線スタイル(EXCEL)
        /// XlLineStyle
        /// </summary>
        public object LineStyle
        {
            get
            {
                return BorderStyle.GetCommonProperty(new BorderStyleParam(StyleName.XlsBorder.LineStyle));
            }
            set
            {
                XlLineStyle XlValue;
                if (XlLineStyleParser.Try(value, out XlValue))
                {
                    BorderStyle.UpdateProperty(new BorderStyleParam(StyleName.XlsBorder.LineStyle, XlValue));
                }
                else
                {
                    throw new ArgumentException("LineStyle");
                }
            }
        }

        /// <summary>
        /// 罫線太さ(EXCEL)
        /// XlBorderWeight
        /// </summary>
        public object Weight
        {
            get
            {
                return BorderStyle.GetCommonProperty(new BorderStyleParam(StyleName.XlsBorder.Weight));
            }
            set
            {
                if (value is XlBorderWeight SafeValue)
                {
                    BorderStyle.UpdateProperty(new BorderStyleParam(StyleName.XlsBorder.Weight, SafeValue));
                }
                else
                {
                    throw new ArgumentException("Weight");
                }
            }
        }

        /// <summary>
        /// 罫線色(EXCEL(NPOIでも同じ値)
        /// short
        /// </summary>
        public object ColorIndex
        {
            get
            {
                return BorderStyle.GetCommonProperty(new BorderStyleParam(StyleName.XlsBorder.ColorIndex));
            }
            set
            {
                if (value is short SafeValue)
                {
                    BorderStyle.UpdateProperty(new BorderStyleParam(StyleName.XlsBorder.ColorIndex, SafeValue));
                }
                else
                {
                    throw new ArgumentException("ColorIndex");
                }
            }
        }

        #endregion

        #region "private properties"

        /// <summary>
        /// ISheetインスタンス
        /// </summary>
        private ISheet PoiSheet { get { return Parent.Parent.PoiSheet; } }

        /// <summary>
        /// CellRangeAddressListインスタンス
        /// </summary>
        private CellRangeAddressList SafeAddressList { get { return this.Parent.SafeAddressList; } }

        /// <summary>
        /// 親RangeのSafeAddressList
        /// </summary>
        private XlBordersIndex? BordersIndex { get; }

        /// <summary>
        /// Range情報読み取り書込み
        /// </summary>
        private RangeBorderStyle BorderStyle { get; }

        #endregion

        #endregion
    }
}
