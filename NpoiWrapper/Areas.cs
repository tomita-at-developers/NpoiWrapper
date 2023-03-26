using NPOI.SS.Util;
using System.Collections;
using System.Runtime.CompilerServices;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Areas interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Areas
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    int Count { get; }
    //    Range Item { get; }
    //    [IndexerName("_Default")]
    //    Range this[int Index] { get; }
    //}

    /// <summary>
    /// Areasクラス
    /// Microsoft.Office.Interop.Excel.Areasをエミュレート
    /// Rageクラスプロパティとしてのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Areas : IEnumerable, IEnumerator
    {
        #region "fileds"

        /// <summary>
        /// Itemの実体。
        /// インデックス１開始と言いつつ_Item[0]は存在し、_Item[1]から使い始めているだけ。
        /// なので_Item.LengthはAreas.Countよりひとつ多い。
        /// Array.CreateInstance()でlowerBoundsを指定できるが、多次元のみで1次元では正しく作れない模様。
        /// Array.CreateInstance()で1次元指定し作成：FullName = "Developers.NpoiWrapper.Range[*]  => [*]となる
		/// 期待する結果：FullName = "Developers.NpoiWrapper.Range[]
        /// 一見作れたように見えるがRange[]にはキャストできない。
        /// </summary>
        private readonly Range[] _Item;
        /// <summary>
        /// Enumrator用インデクス
        /// </summary>
        private int EnumeratorIndex = 0;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        internal Areas(Range ParentRange)
        {
            this.Parent = ParentRange;
            //Itemの生成(インデックス１開始にするために１個余計に作る。)
            this._Item = new Range[this.RawAddressList.CountRanges() + 1];
            //CellRangeAddressアドレスループ
            for (int a = 0; a < this.RawAddressList.CountRanges(); a++)
            {
                //Rangeオブジェクトの生成(作り置き)
                CellRangeAddressList AddressList = new CellRangeAddressList();
                AddressList.AddCellRangeAddress(RawAddressList.GetCellRangeAddress(a).Copy());
                //インデックスは１開始なので要注意。
                this._Item[ a + 1 ] = new Range(this.Parent.Parent, AddressList);
            }
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
            if (EnumeratorIndex < this._Item.Length)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public virtual object Current { get { return this._Item[EnumeratorIndex]; } }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumeratorIndex = 0;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// Areasに含まれるRangeの数。Item.Lengthは1多いのでCellRangeAddressの数を返している。
        /// </summary>
        public int Count { get { return this.RawAddressList.CountRanges(); } }
        /// <summary>
        /// AreasがもつRange（インデックスは１開始)
        /// </summary>
        public Range[] Item { get { return this._Item; } }

        #endregion

        #region "private properties"

        /// <summary>
        /// Rangeアドレスリト
        /// </summary>
        private CellRangeAddressList RawAddressList { get { return Parent.RawAddressList; } }

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="index">インデックス(1開始。ただし0も存在しておりnullがセットされる。)</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public Range this[int index] { get{ return this._Item[index]; } }

        #endregion
    }
}
