using log4net.Repository.Hierarchy;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Utils
{
    internal static class CellUtil
    {
        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

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
            IRow Row = Sheet.GetRow(RowIndex);
            if (Row == null)
            {
                Row = Sheet.CreateRow(RowIndex);
                Logger.Debug(
                    "Sheet[" + Sheet.SheetName + "]:Row[" + RowIndex + "] *** Row Created. ***");
            }
            ICell Cell = Row.GetCell(ColumnIndex);
            if (Cell == null)
            {
                Cell = Row.CreateCell(ColumnIndex);
                Logger.Debug(
                    "Sheet[" + Sheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex + "] *** Column Created. ***");
            }
            return Cell;
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
