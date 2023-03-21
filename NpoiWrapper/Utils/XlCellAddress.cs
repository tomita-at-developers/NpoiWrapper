using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Utils
{
    /// <summary>
    /// ExcelのCellアドレス表現
    /// </summary>
    public class XlCellAddress
    {
        #region "fields"

        /// <summary>
        /// CellAddressインスタンス
        /// </summary>
        private readonly CellAddress _Address;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ(R1C1形式)
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        public XlCellAddress(int RowIndex, int ColumnIndex)
        {
            _Address = new CellAddress(RowIndex - 1, ColumnIndex - 1);
        }

        /// <summary>
        /// コンストラクタ(A1形式)
        /// </summary>
        /// <param name="Address"></param>
        public XlCellAddress(string Address)
        {
            _Address = new CellAddress(Address);
        }

        #endregion

        #region "properties"

        /// <summary>
        /// 行番号(1開始)
        /// </summary>
        public int RowIndex { get { return _Address.Row + 1; } }
        /// <summary>
        /// 列番号(1開始)
        /// </summary>
        public int ColumnIndex { get { return _Address.Column + 1; } }
        /// <summary>
        /// A1形式のアドレス文字列
        /// </summary>
        public string A1Format { get { return _Address.FormatAsString(); } }

        #endregion
    }
}
