using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper
{
    using Range = _Range;

    /// <summary>
    /// Cellsクラス
    /// インデクサOverrideのためだけにあるクラス
    /// (Office.Interop.Excelには存在しないクラス)
    /// </summary>
    public class Cells : Range
    {
        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        /// <param name="RelativeTo">アドレスの開始位置を示すCellRangeAddressインスタンス</param>
        internal Cells(Worksheet ParentSheet, CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo = null)
            : base(ParentSheet, RangeAddressList, RelativeTo)
        {
            //なにもしない
        }

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        public override Range this[object RowIndex, object ColumnIndex]
        {
            get
            {
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                //インデックスはintであること
                if (RowIndex is int row && ColumnIndex is int column)
                {
                    //RangeAddressを生成
                    RangeAddressList.AddCellRangeAddress(
                        new CellRangeAddress(row - 1, row - 1, column - 1, column - 1));
                }
                //上記以外は例外スロー
                else
                {
                    throw new ArgumentException("Type of RowIndex, ColumnIndex should be int.");
                }
                //Rangeクラスインスタンス生成
                return new Range(ParentSheet, RangeAddressList, RelativeTo);
            }
        }
    }
}
