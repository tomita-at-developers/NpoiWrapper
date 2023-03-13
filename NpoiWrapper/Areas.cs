using NPOI.SS.Util;

namespace Developers.NpoiWrapper
{
    // Areas interface in Interop.Excel is shown below...
    //  public interface Areas
    //  {
    //      Application Application { get; }
    //      XlCreator Creator { get; }
    //      object Parent { get; }
    //      int Count { get; }
    //      Range Item { get; }
    //      [IndexerName("_Default")]
    //      Range this[int Index] { get; }
    //  }

    /// <summary>
    /// Areasクラス
    /// Microsoft.Office.Interop.Excel.Areasをエミュレート
    /// Rageクラスプロパティとしてのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Areas
    {
        private Worksheet ParentSheet { get; set; }
        private CellRangeAddressList RawAddressList { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        internal Areas(Worksheet ParentSheet, CellRangeAddressList RangeAddressList)
        {
            this.ParentSheet = ParentSheet;
            this.RawAddressList = RangeAddressList;
        }

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="index">インデックス(１開始)</param>
        /// <returns></returns>
        public Range this[int index]
        {
            get
            {
                CellRangeAddressList AddressList = new CellRangeAddressList();
                AddressList.AddCellRangeAddress(RawAddressList.GetCellRangeAddress(index - 1).Copy());
                return new Range(ParentSheet, AddressList);
            }
        }

        /// <summary>
        /// Areasに含まれるRangeの数
        /// </summary>
        public int Count
        {
            get
            {
                return RawAddressList.CountRanges();
            }
        }
    }
}
