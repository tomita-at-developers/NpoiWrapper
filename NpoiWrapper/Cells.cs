using Developers.NpoiWrapper.Utils;
using NPOI.SS.Util;
using System;
using System.Runtime.CompilerServices;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Cellsクラス
    /// インデクサOverrideのためだけにあるクラス
    /// (Office.Interop.Excelには存在しないクラス)
    /// </summary>
    public class Cells : Range
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        #endregion

        #region "constructors"

        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        /// <param name="RelativeTo">アドレスの開始位置を示すCellRangeAddressインスタンス</param>
        internal Cells(Worksheet ParentSheet, CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo = null)
            : base(ParentSheet, RangeAddressList, RelativeTo)
        {
            Logger.Debug(RangeUtil.CellRangeAddressListToString(RangeAddressList));
            //なにもしない
        }

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns>
        /// 行と列を数字で指定する形式のみサポート
        /// </returns>
        [IndexerName("_Default")]
        public override Range this[object RowIndex, object ColumnIndex]
        {
            get
            {
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                //インデックスはintであること
                if (base.TryConvertToInt(RowIndex, out int row) && base.TryConvertToInt(ColumnIndex, out int column))
                {
                    //１から始まるIndexであること
                    if (row > 0 && column > 0)
                    {
                        //RangeAddressを生成
                        RangeAddressList.AddCellRangeAddress(
                            new CellRangeAddress(row - 1, row - 1, column - 1, column - 1));
                    }
                    else
                    {
                        throw new ArgumentException("RowIndex, ColumnIndex should be 1-origined integer.");
                    }
                }
                //上記以外は例外スロー
                else
                {
                    throw new ArgumentException("Type of RowIndex, ColumnIndex should be int.");
                }
                //Rangeクラスインスタンス生成
                return new Range(base.Parent, RangeAddressList, base.RelativeTo);
            }
        }

        #endregion
    }
}
