using NPOI.SS.Formula.Functions;
using NPOI.SS.Util;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Collections;

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
        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }
        public Range[] Item { get; }

        private CellRangeAddressList RawAddressList { get { return Parent.RawAddressList; } }

        /// <summary>
        /// Enumrator用インデクス
        /// </summary>
        private int EnumeratorIndex = -1;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        internal Areas(Range ParentRange)
        {
            this.Parent = ParentRange;
            //Itemの生成(コンストラクト時に作り置き)
            this.Item = new Range[this.RawAddressList.CountRanges()];
            for (int a = 0; a < this.RawAddressList.CountRanges(); a++)
            {
                CellRangeAddressList AddressList = new CellRangeAddressList();
                AddressList.AddCellRangeAddress(RawAddressList.GetCellRangeAddress(a).Copy());
                Item[a] = new Range(this.Parent.Parent, AddressList);
            }
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
            EnumeratorIndex += 1;
            if (EnumeratorIndex < Item.Length)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public virtual object Current { get { return Item[EnumeratorIndex]; } }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumeratorIndex = -1;
        }

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="index">インデックス(１開始)</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public Range this[int index] { get{ return Item[index]; } }

        /// <summary>
        /// Areasに含まれるRangeの数
        /// </summary>
        public int Count { get{ return Item.Length; } }
    }
}
