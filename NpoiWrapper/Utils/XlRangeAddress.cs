using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper.Utils
{
    /// <summary>
    /// ExcelのRangeアドレス表現
    /// </summary>
    public class XlRangeAddress
    {
        #region "fields"

        /// <summary>
        /// CellRangeAddressインスタンス
        /// </summary>
        private readonly CellRangeAddress _Address;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ(R1C1形式)
        /// </summary>
        /// <param name="FirstRowIndex">先頭行番号(1開始)</param>
        /// <param name="FirstColumnInex">先頭列番号(1開始)</param>
        /// <param name="LastRowIndex">最終行番号(1開始)</param>
        /// <param name="LastColumnIndex">最終列番号(1開始)</param>
        public XlRangeAddress(int FirstRowIndex, int FirstColumnInex, int LastRowIndex, int LastColumnIndex)
        {
            _Address = new CellRangeAddress(FirstRowIndex - 1, LastRowIndex - 1, FirstColumnInex - 1, LastColumnIndex - 1);
        }

        #endregion

        #region "properties"

        /// <summary>
        /// コンストラクタ(A1形式)
        /// </summary>
        /// <param name="Reference">A1形式のRangeアドレス</param>
        public XlRangeAddress(string Reference)
        {
            _Address = CellRangeAddress.ValueOf(Reference);
        }

        /// <summary>
        /// 先頭行番号(1開始)
        /// </summary>
        public int FirstRowIndex { get { return _Address.FirstRow + 1; } }
        /// <summary>
        /// 先頭列番号(1開始)
        /// </summary>
        public int FirstColumnIndex { get { return _Address.FirstColumn + 1; } }
        /// <summary>
        /// 最終行番号(1開始)
        /// </summary>
        public int LastRowIndex { get { return _Address.LastRow + 1; } }
        /// <summary>
        /// 最終列番号
        /// </summary>
        public int LastColumnIndex { get { return _Address.LastColumn + 1; } }
        /// <summary>
        /// A1形式のRangeアドレス文字列
        /// </summary>
        public string A1Format { get { return _Address.FormatAsString(); } }

        #endregion
    }
}
