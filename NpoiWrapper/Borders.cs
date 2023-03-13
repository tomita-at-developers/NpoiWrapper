using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    // Border interface in Interop.Excel is shown below...
    //  public interface Borders
    //  {
    //      Application Application { get; }
    //      XlCreator Creator { get; }
    //      object Parent { get; }
    //      object Color { get; set; }
    //      object ColorIndex { get; set; }
    //      int Count { get; }
    //      Border Item { get; }
    //      object LineStyle { get; set; }
    //      object Value { get; set; }
    //      object Weight { get; set; }
    //      Border this[XlBordersIndex Index] { get; }
    //      object ThemeColor { get; set; }
    //      object TintAndShade { get; set; }
    //  }

    /// <summary>
    /// Bordersクラス
    /// Excelにあるが、このクラスではサポートしていないプロパティは以下の通り。
    ///     Application Application
    ///     XlCreator Creator
    ///     object Parent
    ///     object Color
    ///     Border Item
    ///     object ThemeColor
    ///     object TintAndShade
    /// </summary>
    public class Borders : IEnumerable, IEnumerator
    {
        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// ISheetインスタンス
        /// </summary>
        private ISheet PoiSheet { get; }

        /// <summary>
        /// CellRangeAddressListインスタンス
        /// </summary>
        private CellRangeAddressList SafeRangeAddressList { get; }

        /// <summary>
        /// Borderリスト
        /// </summary>
        private Dictionary<XlBordersIndex, Border> IndexedBorderList { get; } = new Dictionary<XlBordersIndex, Border>();

        /// <summary>
        /// 全てのXlBordersIndexを支配するBorderインスタンス
        /// </summary>
        private Border EntireRangeBorder { get; }
        
        /// <summary>
        /// BorderIndexリスト
        /// </summary>
        private List<XlBordersIndex> BoerderIndexList { get; } = new List<XlBordersIndex>()
        {
            XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom,
            XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight,
            XlBordersIndex.xlInsideHorizontal, XlBordersIndex.xlInsideVertical
        };

        /// <summary>
        /// IndexedBorderList取り出しインデクス
        /// </summary>
        private int EnumBorderIndex { get; set; } = -1;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        internal Borders(ISheet PoiSheet, CellRangeAddressList SafeAddressList)
        {
            //親Range情報の保存
            this.PoiSheet = PoiSheet;
            this.SafeRangeAddressList = SafeAddressList;
            //IndexedBorderListのメンバー生成
            foreach (XlBordersIndex Index in BoerderIndexList)
            {
                IndexedBorderList.Add(Index, new Border(PoiSheet, SafeAddressList, Index));
            }
            //EntireRangeBorderのインスタンス生成(BordersIndexはnull)
            EntireRangeBorder = new Border(this.PoiSheet, this.SafeRangeAddressList, null);
        }

        /// <summary>
        /// GetEnumeratorの実装
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            Reset();
            return (IEnumerator)this;
        }
        /// <summary>
        /// IEnumerator.MoveNextの実装
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            bool RetVal = false;
            EnumBorderIndex += 1;
            if (EnumBorderIndex < BoerderIndexList.Count)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public virtual object Current
        {
            get
            {
                //EnumBorderIndexでBoerderIndexListからDictionaryのキーを取り出してそのキーで､､､
                return IndexedBorderList[BoerderIndexList[EnumBorderIndex]];
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumBorderIndex = -1;
        }

        /// <summary>
        /// Borders.Countプロパティ
        /// </summary>
        public int Count
        {
            get
            {
                //Interop.Excelでは、どうやら常に6になる模様。
                return BoerderIndexList.Count;
            }
        }

        /// <summary>
        /// Bordersインデクサ
        /// </summary>
        public Border this[XlBordersIndex Index]
        {
            get
            {
                return IndexedBorderList[Index];
            }
        }

        /// <summary>
        /// (XlLineStyle)Borders.LineStyleプロパティ
        /// </summary>
        public object LineStyle
        {
            get
            {
                return EntireRangeBorder.LineStyle;
            }
            set
            {
                if (value is XlLineStyle SafeValue)
                {
                    EntireRangeBorder.LineStyle = SafeValue;
                }
                else
                {
                    throw new ArgumentException("LineStyle");
                }
            }
        }

        /// <summary>
        /// Borders.Valueプロパティ(LineStyleに等しい)
        /// </summary>
        public object Value
        {
            get
            {
                return EntireRangeBorder.LineStyle;
            }
            set
            {
                if (value is XlLineStyle SafeValue)
                {
                    EntireRangeBorder.LineStyle = SafeValue;
                }
                else
                {
                    throw new ArgumentException("Value(LineStyle)");
                }
            }
        }

        /// <summary>
        /// (XlBorderWeight)Borders.Weightプロパティ
        /// </summary>
        public object Weight
        {
            get
            {
                return EntireRangeBorder.Weight;
            }
            set
            {
                if (value is XlBorderWeight SafeValue)
                {
                    EntireRangeBorder.Weight = SafeValue;
                }
                else
                {
                    throw new ArgumentException("Weight");
                }
            }
        }

        /// <summary>
        /// (short)Borders.ColorIndexプロパティ
        /// </summary>
        public object ColorIndex
        {
            get
            {
                return EntireRangeBorder.ColorIndex;
            }
            set
            {
                if (value is short SafeValue)
                {
                    EntireRangeBorder.ColorIndex = SafeValue;
                }
                else
                {
                    throw new ArgumentException("ColorIndex");
                }
            }
        }

        /// <summary>
        /// 囲み線の設定
        /// </summary>
        /// <param name="LineStyle"></param>
        /// <param name="Weight"></param>
        /// <param name="ColorIndex"></param>
        /// <param name="Color"></param>
        /// <returns></returns>
        internal bool Around(object LineStyle, XlBorderWeight Weight, XlColorIndex ColorIndex, object Color)
        {
            //４周囲のターゲットとする。
            List<XlBordersIndex> TargetIndex = new List<XlBordersIndex>
            { XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight };
            //ターゲットに対して指定された更新を適用
            foreach (XlBordersIndex Index in TargetIndex)
            {
                //LineStyle指定があれば適用
                if (LineStyle is XlLineStyle SafeLineStyle)
                {
                    this.IndexedBorderList[Index].LineStyle = SafeLineStyle;
                }
                //Weightを適用
                this.IndexedBorderList[Index].Weight = Weight;
                //色指定判断
                short? IndexedColor;
                //自動指定
                if (ColorIndex == XlColorIndex.xlColorIndexAutomatic)
                {
                    IndexedColor = IndexedColors.Automatic.Index;
                }
                //なし？？？
                else if (ColorIndex == XlColorIndex.xlColorIndexNone)
                {
                    IndexedColor = null;
                }
                //その他(具体的なカラーパレット上の色インデックス)
                else
                {
                    IndexedColor = (short)ColorIndex;
                }
                //ColorIndexに有効な指定があれば適用
                if (IndexedColor != null)
                {
                    this.IndexedBorderList[Index].ColorIndex = IndexedColor;
                }
                //Color設定がある場合
                if (Color != null)
                {
                    //未サポートにつき無視
                }
            }
            //常にtrueでリターン
            return true;
        }
    }
}
