using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper.Styles
{
    /// <summary>
    /// RangeStyleManger
    /// Excel名前体系(Range)とPOI名前体系(ICellSttyle)の間で、読み書きを仲介する。
    /// </summary>
    internal class RangeStyleManager : RangeStyle
    {
        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// Fontクラスインスタンス
        /// </summary>
        private Font _Font = null;

        /// <summary>
        /// Bordersクラスインスタンス
        /// </summary>
        private Borders _Borders = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        public RangeStyleManager(ISheet PoiSheet, CellRangeAddressList SafeAddressList)
            : base(PoiSheet, SafeAddressList)
        {
            //何もしない
        }

        public Font Font
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Font == null) { _Font = new Font(base.PoiSheet, base.SafeRangeAddressList); }
                return _Font;
            }
        }

        public Borders Borders
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Borders == null) { _Borders = new Borders(base.PoiSheet, base.SafeRangeAddressList); }
                return _Borders;
            }
        }
        /// <summary>
        /// HorizontalAlignment
        /// XlHAlignはNPOI.SS.UserModel.HorizontalAlignmentの値で定義しており同義。
        /// </summary>
        public object HorizontalAlignment
        {
            get
            {
                return (XlHAlign?)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.HorizontaiAlignment));
            }
            set
            {
                if (value is XlHAlign SafeValue)
                {
                    List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                    { { new Properties.CellStyleParam(Properties.StyleName.HorizontaiAlignment, (HorizontalAlignment)SafeValue) } };
                    UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("HorizontalAlignment");
                }
            }
        }

        /// <summary>
        /// VerticalAlignment
        /// </summary>
        public object VerticalAlignment
        {
            get
            {
                return (XlVAlign?)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.VerticalAlignment));
            }
            set
            {
                if (value is XlVAlign SafeValue)
                {
                    List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                    { { new Properties.CellStyleParam(Properties.StyleName.VerticalAlignment, (VerticalAlignment)SafeValue) } };
                    UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("VerticalAlignment");
                }
            }
        }

        /// <summary>
        /// 文字書式
        /// string
        /// </summary>
        public object NumberFormatLocal
        {
            get
            {
                return (string)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.NumberFormat));
            }
            set
            {
                //nullでも書きにいく
                List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                { { new Properties.CellStyleParam(Properties.StyleName.NumberFormat, value) } };
                UpdateProperties(Params);
            }
        }

        /// <summary>
        /// 文字書式
        /// string
        /// </summary>
        public object NumberFormat
        {
            get
            {
                return (string)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.NumberFormat));
            }
            set
            {
                //nullでも書きにいく
                List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                { { new Properties.CellStyleParam(Styles.Properties.StyleName.NumberFormat, value) } };
                UpdateProperties(Params);
            }
        }

        /// <summary>
        /// セルの保護
        /// </summary>
        public object Locked
        {
            get
            {
                return (bool)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.Locked));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                    { { new Properties.CellStyleParam(Styles.Properties.StyleName.Locked, SafeValue) } };
                    UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("Locked");
                }
            }
        }

        /// <summary>
        /// 文字列の折り返し
        /// </summary>
        public object WrapText
        {
            get
            {
                return (bool)GetCommonProperty(new Properties.CellStyleParam(Properties.StyleName.WrapText));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    List<Properties.CellStyleParam> Params = new List<Properties.CellStyleParam>
                    { { new Properties.CellStyleParam(Properties.StyleName.WrapText, SafeValue) } };
                    UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("WrapText");
                }
            }
        }
        /// <summary>
        /// 囲み罫線
        /// </summary>
        /// <param name="LineStyle">線のスタイル</param>
        /// <param name="Weight">線の太さ</param>
        /// <param name="ColorIndex">線の色(カラーパレット上の色インデックス</param>
        /// <param name="Color">RGBの色指定(未サポート)</param>
        /// <returns></returns>
        public bool BorderAround(
                            object LineStyle = null, XlBorderWeight Weight = XlBorderWeight.xlThin,
                            XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic, object Color = null)
        {
            return Borders.Around(LineStyle, Weight, ColorIndex, Color);
        }
    }
}
