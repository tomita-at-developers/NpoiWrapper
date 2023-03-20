using NPOI.POIFS.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper.Utils
{
    internal static class CellUtil
    {
        /// <summary>
        /// 指定した位置のセルを取得する(なければ生成)
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="Address">セルのCellAddress</param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(ISheet Sheet, CellAddress Address)
        {
            return GetOrCreateCell(Sheet, Address.Row, Address.Column);
        }

        /// <summary>
        /// 指定した位置のセルを取得する(なければ生成)
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行Index</param>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            IRow Row = Sheet.GetRow(RowIndex) ?? Sheet.CreateRow(RowIndex);
            return Row.GetCell(ColumnIndex) ?? Row.CreateCell(ColumnIndex);
        }

        /// <summary>
        /// 指定した位置のセルを取得する(なければnullでリターン)
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="Address">セルのCellAddress</param>
        /// <returns></returns>
        public static ICell GetCell(ISheet Sheet, CellAddress Address)
        {
            return GetCell(Sheet, Address.Row, Address.Column);
        }

        /// <summary>
        /// 指定した位置のセルを取得する(なければnullでリターン)
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行Index</param>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public static ICell GetCell(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            ICell RetVal = null;
            IRow Row = Sheet.GetRow(RowIndex);
            if (Row != null)
            {
                RetVal = Row.GetCell(ColumnIndex);
            }
            return RetVal;
        }
    }
}
