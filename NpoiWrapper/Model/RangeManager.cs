using Developers.NpoiWrapper.Model.Param;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper.Model
{
    /// <summary>
    /// RangeStyleManger
    /// Excel名前体系(Range)とPOI名前体系(ICellSttyle)の間で、読み書きを仲介する。
    /// </summary>
    internal class RangeManager : RangeStyle
    {
        #region "fields"

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
        /// Interiorクラスインスタンス
        /// </summary>
        private Interior _Interior = null;

        /// <summary>
        /// RangeCommentクラスインスタンス
        /// </summary>
        private RangeComment _RangeComment = null;

        /// <summary>
        /// Validationクラスインスタンス
        /// </summary>
        private Validation _Validation = null;

        /// <summary>
        /// RangeValueクラスインスタンス
        /// </summary>
        private RangeValue _RangeValue = null;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        public RangeManager(Range ParentRange)
            : base(ParentRange)
        {
            //何もしない
        }

        #endregion

        #region "properties"

        /// <summary>
        /// 文字フォント
        /// </summary>
        public Font Font
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Font == null) { _Font = new Font(base.ParentRange); }
                return _Font;
            }
        }

        /// <summary>
        /// 罫線
        /// </summary>
        public Borders Borders
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Borders == null) { _Borders = new Borders(base.ParentRange); }
                return _Borders;
            }
        }

        /// <summary>
        /// 塗りつぶし
        /// </summary>
        public Interior Interior
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Interior == null) { _Interior = new Interior(base.ParentRange); }
                return _Interior;
            }
        }

        /// <summary>
        /// コメント
        /// </summary>
        public Comment Comment
        {
            get
            {
                return RangeComment.Comment;
            }
        }

        /// <summary>
        /// 入力規則
        /// </summary>
        public Validation Validation
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_Validation == null) { _Validation = new Validation(base.ParentRange); }
                return _Validation;
            }
        }

        /// <summary>
        /// セルの値Value
        /// </summary>
        public object Value
        {
            get { return RangeValue.Value; }
            set { RangeValue.Value = value; }   
        }

        /// <summary>
        /// セルの値Value2
        /// </summary>
        public object Value2
        {
            get { return RangeValue.Value2; }
            set { RangeValue.Value2 = value; }
        }

        /// <summary>
        /// セルの値Text
        /// </summary>
        public object Text
        {
            get { return RangeValue.Text; }
        }

        /// <summary>
        /// セルの式
        /// </summary>
        public object Formula
        {
            get { return RangeValue.Formula; }
            set { RangeValue.Formula = value; }
        }

        /// <summary>
        /// HorizontalAlignment
        /// XlHAlignはNPOI.SS.UserModel.HorizontalAlignmentの値で定義しており同義。
        /// </summary>
        public object HorizontalAlignment
        {
            get
            {
                XlHAlign? RetVal = null;
                object RawVal = GetCommonProperty(new CellStyleParam(Utils.StyleName.HorizontaiAlignment));
                if (XlHAlignParser.Try(RawVal, out HorizontalAlignment PoiValue))
                {
                    RetVal = XlHAlignParser.GetXlValue(PoiValue);
                }
                return RetVal;
            }
            set
            {
                if (XlHAlignParser.Try(value, out XlHAlign XlValue))
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(Utils.StyleName.HorizontaiAlignment, XlHAlignParser.GetPoiValue(XlValue)) } };
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
                XlVAlign? RetVal = null;
                object RawVal = GetCommonProperty(new CellStyleParam(Utils.StyleName.VerticalAlignment));
                if (XlVAlignParser.Try(RawVal, out VerticalAlignment PoiValue))
                {
                    RetVal = XlVAlignParser.GetXlValue(PoiValue);
                }
                return RetVal;
            }
            set
            {
                if (XlVAlignParser.Try(value, out XlVAlign XlValue))
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(Utils.StyleName.VerticalAlignment, XlVAlignParser.GetPoiValue(XlValue)) } };
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
                return (string)GetCommonProperty(new CellStyleParam(Utils.StyleName.NumberFormat));
            }
            set
            {
                //nullでも書きにいく
                List<CellStyleParam> Params = new List<CellStyleParam>
                { { new CellStyleParam(Utils.StyleName.NumberFormat, value) } };
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
                return (string)GetCommonProperty(new CellStyleParam(Utils.StyleName.NumberFormat));
            }
            set
            {
                //nullでも書きにいく
                List<CellStyleParam> Params = new List<CellStyleParam>
                { { new CellStyleParam(Model.Utils.StyleName.NumberFormat, value) } };
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
                return (bool?)GetCommonProperty(new CellStyleParam(Utils.StyleName.Locked));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(Model.Utils.StyleName.Locked, SafeValue) } };
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
                return (bool?)GetCommonProperty(new CellStyleParam(Utils.StyleName.WrapText));
            }
            set
            {
                if (value is bool SafeValue)
                {
                    List<CellStyleParam> Params = new List<CellStyleParam>
                    { { new CellStyleParam(Utils.StyleName.WrapText, SafeValue) } };
                    UpdateProperties(Params);
                }
                else
                {
                    throw new ArgumentNullException("WrapText");
                }
            }
        }

        /// <summary>
        /// Cellの値
        /// </summary>
        public RangeValue RangeValue
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_RangeValue == null) { _RangeValue = new RangeValue(base.ParentRange); }
                return _RangeValue;
            }
        }

        /// <summary>
        /// Cellのコメント
        /// </summary>
        private RangeComment RangeComment
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_RangeComment == null) { _RangeComment = new RangeComment(base.ParentRange); }
                return _RangeComment;
            }
        }

        #endregion

        #region "methods"

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

        /// <summary>
        /// コメントの追加
        /// </summary>
        /// <param name="Text">コメント文字列</param>
        /// <returns></returns>
        public Comment AddComment(object Text = null)
        {
            return RangeComment.AddComment(Text);
        }

        #endregion
    }
}
