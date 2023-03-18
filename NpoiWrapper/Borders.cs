using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Borders interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
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
    //      [IndexerName("_Default")]
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
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// Borderリスト
        /// </summary>
        private readonly Dictionary<XlBordersIndex, Border> _Item = new Dictionary<XlBordersIndex, Border>();

        /// <summary>
        /// BorderIndexリスト
        /// foreachで取得する場合は斜線２種類を除く下記６種類
        /// </summary>
        private readonly List<XlBordersIndex> BoerderIndexList = new List<XlBordersIndex>()
        {
            XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom,
            XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight,
            XlBordersIndex.xlInsideHorizontal, XlBordersIndex.xlInsideVertical
        };

        /// <summary>
        /// 全てのXlBordersIndexを支配するBorderインスタンス
        /// </summary>
        private readonly Border EntireRangeBorder;

        /// <summary>
        /// Enumerator用インデクス
        /// </summary>
        private int EnumeratorIndex = -1;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">Rangeインスタンス</param>
        internal Borders(Range ParentRange)
        {
            //親Range情報の保存
            this.Parent = ParentRange;
            //XlBordersIndexで定義される全８種類のメンバーをすべて生成
            foreach (XlBordersIndex Index in Enum.GetValues(typeof(XlBordersIndex)))
            {
                _Item.Add(Index, new Border(this.Parent, Index));
            }
            //EntireRangeBorderのインスタンス生成(BordersIndexはnull)
            EntireRangeBorder = new Border(this.Parent, null);
        }

        #endregion

        #region "interface implementations"

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
            EnumeratorIndex += 1;
            if (EnumeratorIndex < BoerderIndexList.Count)
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
                return _Item[BoerderIndexList[EnumeratorIndex]];
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumeratorIndex = -1;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// C#では引数を指定するプロパティが実現できなかった。
        /// なのでItemは実装しない。Indexer経由での取得のみサポート。
        /// </summary>
        //public Border Item { get; }

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

        #endregion

        #endregion

        #region "methods"

        #region "internal methods"

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
                    this._Item[Index].LineStyle = SafeLineStyle;
                }
                //Weightを適用
                this._Item[Index].Weight = Weight;
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
                    this._Item[Index].ColorIndex = IndexedColor;
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

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// Bordersインデクサ
        /// </summary>
        [IndexerName("_Default")]
        public Border this[XlBordersIndex Index]
        {
            get
            {
                return this._Item[Index];
            }
        }

        #endregion

    }
}
