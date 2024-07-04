using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper.Common
{
    /// <summary>
    /// レンジアドレスクラス
    /// </summary>
    public class CellRange
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="Range">A1形式のレンジアドレス(カンマ区切りで複数指定された場合は先頭のみ参照)</param>
        public CellRange(string References)
        {
            string [] ReferenceList = References.Split(',');
            Address = CellRangeAddress.ValueOf(ReferenceList[0]);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="FirstRow">先頭行Index(1-base)</param>
        /// <param name="LastRow">最終行Index(1-base)</param>
        /// <param name="FirstColumn">先頭列Index(1-base)</param>
        /// <param name="LastColumn">最終列Index(1-base)</param>
        public CellRange(int RowIndex, int ColumnIndex)
        {
            Address = new CellRangeAddress(RowIndex, RowIndex, ColumnIndex, ColumnIndex);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="FirstRow">先頭行Index(1-base)</param>
        /// <param name="LastRow">最終行Index(1-base)</param>
        /// <param name="FirstColumn">先頭列Index(1-base)</param>
        /// <param name="LastColumn">最終列Index(1-base)</param>
        public CellRange(int FirstRow, int LastRow, int FirstColumn, int LastColumn)
        {
            if (FirstRow >= 1 && LastRow >= 1 && FirstColumn >= 1 && LastColumn >= 1)
            {
                Address = new CellRangeAddress(FirstRow - 1, LastRow - 1, FirstColumn - 1, LastColumn - 1);
            }
            else
            {
                throw new ArgumentException("Index out of range.");
            }
        }

        #endregion

        #region "properties"

        private CellRangeAddress Address { get; set; } = new CellRangeAddress(0, 0, 0, 0);

        /// <summary>
        /// A1形式のレンジアドレス
        /// </summary>
        public string Reference
        {
            get { return Address.FormatAsString(null, false); }
        }

        /// <summary>
        /// この範囲に含まれるセルの個数
        /// </summary>
        public int CellCount
        {
            get { return Address.NumberOfCells; }
        }

        /// <summary>
        /// 先頭行Index(1-base)
        /// </summary>
        public int FirstRowIndex
        {
            get { return Address.FirstRow + 1; }
        }

        /// <summary>
        /// 最終行Index(1-base)
        /// </summary>
        public int LastRowIndex
        {
            get { return Address.LastRow + 1; }
        }

        /// <summary>
        /// 先頭列Index(1-base)
        /// </summary>
        public int FirstColumnIndex
        {
            get { return Address.FirstColumn + 1; }
        }

        /// <summary>
        /// 最終列Index(1-base)
        /// </summary>
        public int LastColumnIndex
        {
            get { return Address.LastColumn + 1; }
        }

        #endregion

    }
}
