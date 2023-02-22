using Developers.NpoiWrapper;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
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
        internal Cells(Worksheet ParentSheet, CellRangeAddressList RangeAddressList)
            : base(ParentSheet, RangeAddressList)
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
                if (RowIndex is int row && ColumnIndex is int column)
                {
                    //RangeAddressを生成
                    RangeAddressList.AddCellRangeAddress(
                        new CellRangeAddress(row - 1, row - 1, column - 1, column - 1));
                }
                else
                {

                }
                //Rangeクラスインスタンス生成
                return new Range(ParentSheet, RangeAddressList);
            }
        }
    }
}
