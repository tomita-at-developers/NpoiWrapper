using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Styles.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper
{
    // Border interface in Interop.Excel is shown below...
    //  public interface Border
    //  {
    //      Application Application { get; }
    //      XlCreator Creator { get; }
    //      object Parent { get; }
    //      object Color { get; set; }
    //      object ColorIndex { get; set; }
    //      object LineStyle { get; set; }
    //      object Weight { get; set; }
    //      object ThemeColor { get; set; }
    //      object TintAndShade { get; set; }
    //  }

    public class Border
    {
        /// <summary>
        /// ISheetインスタンス
        /// </summary>
        private ISheet PoiSheet { get; set; }

        /// <summary>
        /// CellRangeAddressListインスタンス
        /// </summary>
        private CellRangeAddressList SafeRangeAddressList { get; }

        /// <summary>
        /// 親RangeのSafeAddressList
        /// </summary>
        private XlBordersIndex? BordersIndex { get; }

        /// <summary>
        /// Range情報読み取り書込み
        /// </summary>
        private RangeBorderStyle BorderStyle { get; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        /// <param name="BordersIndex">XlBordersIndex値</param>
        public Border(
            ISheet PoiSheet, CellRangeAddressList SafeAddressList, XlBordersIndex? BordersIndex)
        {
            //親Range情報の保存
            this.PoiSheet = PoiSheet;
            this.SafeRangeAddressList = SafeAddressList;
            //Border情報の保存
            this.BordersIndex = BordersIndex;
            //RangeBorderStyle生成
            BorderStyle = new RangeBorderStyle(this.PoiSheet, this.SafeRangeAddressList, this.BordersIndex);
        }

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
                if (value is XlLineStyle SafeValue)
                {
                    BorderStyle.UpdateProperty(new BorderStyleParam(StyleName.XlsBorder.LineStyle, SafeValue));
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
    }
}
