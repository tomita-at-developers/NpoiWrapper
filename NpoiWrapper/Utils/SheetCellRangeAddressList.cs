using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Utils
{
    /// <summary>
    /// SheetCellRangeAddressList
    /// CellRangeAddressListとの違いは以下の通り。
    /// ・SheetIndex情報を保持している。
    /// ・SheetIndex評価を加えたInRangeメソッドを実装している。
    /// </summary>
    internal class SheetCellRangeAddressList
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SheetCellRangeAddressList()
            : this(-1, new CellRangeAddressList())
        {
            //何もしない
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="SheetIndex">シート番号(0開始)</param>
        /// <param name="RangeAddressList">CellRangeAddressList</param>
        public SheetCellRangeAddressList(int SheetIndex, CellRangeAddressList CellRangeAddressList)
        {
            this.SheetIndex = SheetIndex;
            this.RangeAddressList = CellRangeAddressList;
        }

        #endregion

        #region "properties"

        /// <summary>
        /// シート番号
        /// </summary>
        public int SheetIndex { get; set; } = -1;
        public CellRangeAddressList RangeAddressList { get; set; }

        #endregion

        #region "methods"

        /// <summary>
        /// 指定されたシートがRangeListに含まれるか判定
        /// </summary>
        /// <param name="SheetIndex"></param>
        /// <returns></returns>
        public bool IsSheetInRange(int SheetIndex)
        {
            return (SheetIndex == this.SheetIndex);
        }

        /// <summary>
        /// 指定された行がRangeListに含まれるか判定
        /// </summary>
        /// <param name="SheetIndex"></param>
        /// <param name="RowIndex"></param>
        /// <returns></returns>
        public bool IsRowInRange(int SheetIndex, int RowIndex)
        {
            bool RetVal = false;
            if (IsSheetInRange(SheetIndex))
            {
                for (int a = 0; a < RangeAddressList.CountRanges(); a++)
                {
                    CellRangeAddress RangeAddress = RangeAddressList.GetCellRangeAddress(a);
                    if (RangeAddress.FirstRow < 0 || RangeAddress.FirstRow <= RowIndex)
                    {
                        if (RangeAddress.LastRow < 0 || RowIndex <= RangeAddress.LastRow)
                        {
                            RetVal = true;
                            break;
                        }
                    }

                }
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたセルがRangeListに含まれるか判定
        /// </summary>
        /// <param name="SheetIndex"></param>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        public bool IsCellInRange(int SheetIndex, int RowIndex, int ColumnIndex)
        {
            bool RetVal = false;
            if (IsSheetInRange(SheetIndex))
            {
                for (int a = 0; a < RangeAddressList.CountRanges(); a++)
                {
                    CellRangeAddress RangeAddress = RangeAddressList.GetCellRangeAddress(a);
                    if (RangeAddress.IsInRange(RowIndex, ColumnIndex))
                    {
                        RetVal = true;
                        break;
                    }

                }
            }
            return RetVal;
        }

        #endregion
    }
}

