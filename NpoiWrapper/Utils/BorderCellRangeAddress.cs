using NPOI.SS;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper.Utils
{
    [Flags]
    enum BorderInRange
    {
        EdgeTop = 1,
        EdgeBottom = 2,
        EdgeLeft = 4,
        EdgeRight = 8,
        InsideTop = 16,
        InsideBottom = 32,
        InsideLeft = 64,
        InsideRight = 128
    }

    /// <summary>
    /// BorderCellRangeAddress
    /// 罫線の種類を意識したCellRangeAddressクラス
    /// </summary>
    internal class BorderCellRangeAddress : CellRangeAddress
    {
        //Range分割配列Index
        public const int FirstRangeIndex = 0;
        public const int InternalRangeIndex = 1;
        public const int LastRangeIndex = 2;

        /// <summary>
        /// バージョン情報。シートの最大行数/列数の判断に利用。
        /// </summary>
        private SpreadsheetVersion SheetVersion { get; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="FirstRow">先頭行Index</param>
        /// <param name="LastRow">最終行Index</param>
        /// <param name="FirstColumn">先頭列Index</param>
        /// <param name="LastColumn">最終列Index</param>
        /// <param name="SheetVersion">シートバージョン</param>
        public BorderCellRangeAddress(int FirstRow, int LastRow, int FirstColumn, int LastColumn, SpreadsheetVersion Version)
            :base(FirstRow, LastRow, FirstColumn, LastColumn)
        {
            this.SheetVersion = Version;
            //念のため安全化
            base.FirstRow = base.FirstRow < 0 ? 0 : base.FirstRow;
            base.LastRow = base.LastRow < 0 ? this.SheetVersion.LastRowIndex : base.LastRow;
            base.FirstColumn = base.FirstColumn < 0 ? 0 : FirstColumn;
            base.LastColumn = base.LastColumn < 0 ? this.SheetVersion.LastColumnIndex : base.LastColumn;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="CellRangeAddress">CellRangeAddress</param>
        public BorderCellRangeAddress(CellRangeAddress CellRangeAddress, SpreadsheetVersion Version)
            : this(
                  CellRangeAddress.FirstRow, CellRangeAddress.LastRow,
                  CellRangeAddress.FirstColumn, CellRangeAddress.LastColumn, Version)
        {
        }

        /// <summary>
        /// Rangeに含まれる行の数
        /// </summary>
        public int NumberOfRows { get{ return base.LastRow - base.FirstRow + 1; } }
        /// <summary>
        /// Rangeに含まれる列の数
        /// </summary>
        public int NumberOfColumns { get{ return base.LastColumn - base.FirstColumn + 1; } }
        /// <summary>
        /// このRangeがInsideVerticalを持つかどうか判定
        /// </summary>
        public bool HasInsideVertical { get { return (NumberOfRows > 1); } }
        /// <summary>
        /// このRangeがInsideHorizontalを持つかどうか判定
        /// </summary>
        public bool HasInsideHorizotal { get { return (NumberOfColumns > 1); } }
        /// <summary>
        /// 指定された行が先頭行か判定
        /// </summary>
        /// <param name="RowIndex">行Index</param>
        /// <returns></returns>
        public bool IsFirstRow(int RowIndex) { return (RowIndex == base.FirstRow); }
        /// <summary>
        /// 指定された行が最終行か判定
        /// </summary>
        /// <param name="RowIndex">行Index</param>
        /// <returns></returns>
        public bool IsLastRow(int RowIndex) { return (RowIndex == base.LastRow); }
        /// <summary>
        /// 指定された行の前に行があるか判定
        /// </summary>
        /// <param name="RowIndex">行Index</param>
        /// <returns></returns>
        public bool HasPreviousRow(int RowIndex) { return (base.FirstRow < RowIndex); }
        /// <summary>
        /// 指定された行の後に行があるか判定
        /// </summary>
        /// <param name="RowIndex">行Index</param>
        /// <returns></returns>
        public bool HasNextRow(int RowIndex) { return (RowIndex < base.LastRow); }
        /// <summary>
        /// 指定された列が先頭列か判定
        /// </summary>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public bool IsFirstColumn(int ColumnIndex) { return (ColumnIndex == base.FirstColumn); }
        /// <summary>
        /// 指定された列が最終列か判定
        /// </summary>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public bool IsLastColumn(int ColumnIndex) { return (ColumnIndex == base.LastColumn); }
        /// <summary>
        /// 指定された行の前に行があるか判定
        /// </summary>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public bool HasPreviousColumn(int ColumnIndex) { return (base.FirstColumn < ColumnIndex); }
        /// <summary>
        /// 指定された列の後に列があるか判定
        /// </summary>
        /// <param name="ColumnIndex">列Index</param>
        /// <returns></returns>
        public bool HasNextColumn(int ColumnIndex) { return (ColumnIndex < base.LastColumn); }

        /// <summary>
        /// 指定されたBorderIndexに対応するCellRangeAddressの取得
        /// </summary>
        /// <param name="BorderIndex">BorderIndex</param>
        /// <returns></returns>
        public BorderCellRangeAddress GetIndexedBorderCellRangeAddress(XlBordersIndex? BorderIndex)
        {
            int FirstRow;
            int LastRow;
            int FirstColumn;
            int LastColumn;
            XlBordersIndex Index = BorderIndex ?? (XlBordersIndex)( -1);
            switch (Index)
            {
                //上枠(BorderTopの更新対象)
                case XlBordersIndex.xlEdgeTop:
                    FirstRow = base.FirstRow;
                    LastRow = base.FirstRow;
                    FirstColumn = base.FirstColumn;
                    LastColumn = base.LastColumn;
                    break;
                //下枠(BorderBottomの更新対象)
                case XlBordersIndex.xlEdgeBottom:
                    FirstRow = base.LastRow;
                    LastRow = base.LastRow;
                    FirstColumn = base.FirstColumn;
                    LastColumn = base.LastColumn;
                    break;
                //左枠(BorderLeftの更新対象)
                case XlBordersIndex.xlEdgeLeft:
                    FirstRow = base.FirstRow;
                    LastRow = base.LastRow;
                    FirstColumn = base.FirstColumn;
                    LastColumn = base.FirstColumn;
                    break;
                //右枠(BorderRigtの更新対象)
                case XlBordersIndex.xlEdgeRight:
                    FirstRow = base.FirstRow;
                    LastRow = base.LastRow;
                    FirstColumn = base.LastColumn;
                    LastColumn = base.LastColumn;
                    break;
                //Inside系, Diagonal系は全Range
                case XlBordersIndex.xlInsideHorizontal:
                case XlBordersIndex.xlInsideVertical:
                case XlBordersIndex.xlDiagonalUp:
                case XlBordersIndex.xlDiagonalDown:
                default:
                    FirstRow = base.FirstRow;
                    LastRow = base.LastRow;
                    FirstColumn = base.FirstColumn;
                    LastColumn = base.LastColumn;
                    break;
            }
            return new BorderCellRangeAddress(FirstRow, LastRow, FirstColumn, LastColumn, this.SheetVersion);
        }

        /// <summary>
        /// 保持しているCellRangeAddressを垂直３種類に分割する
        /// </summary>
        /// <returns></returns>
        public BorderCellRangeAddress[] VerticalSplit()
        {
            //リターン値初期化
            BorderCellRangeAddress[] RetVal = new BorderCellRangeAddress[LastRangeIndex + 1];
            //先頭行Range
            if (NumberOfRows > 0)
            {
                RetVal[FirstRangeIndex]
                    = new BorderCellRangeAddress(
                        base.FirstRow, base.FirstRow, base.FirstColumn, base.LastColumn, this.SheetVersion);
            }
            //中間行Range
            if (NumberOfRows > 3)
            {
                RetVal[InternalRangeIndex] 
                    = new BorderCellRangeAddress(
                            base.FirstRow + 1, base.LastRow - 1, base.FirstColumn, base.LastColumn, this.SheetVersion);
            }
            //最終行Range
            if (NumberOfRows > 2)
            {
                RetVal[LastRangeIndex]
                    = new BorderCellRangeAddress(
                            base.LastRow, base.LastRow, base.FirstColumn, base.LastColumn, this.SheetVersion);
            }
            return RetVal;
        }

        /// <summary>
        /// 保持しているCellRangeAddressを水平３種類に分割する
        /// </summary>
        /// <returns></returns>
        public BorderCellRangeAddress[] HorizontalSplit()
        {
            //リターン値初期化
            BorderCellRangeAddress[] RetVal = new BorderCellRangeAddress[LastRangeIndex + 1];
            //先頭列Range
            if (NumberOfColumns > 0)
            {
                RetVal[FirstRangeIndex]
                    = new BorderCellRangeAddress(
                            base.FirstRow, base.LastRow, base.FirstColumn, base.FirstColumn, this.SheetVersion);
            }
            //中間列Range
            if (NumberOfColumns > 3)
            {
                RetVal[InternalRangeIndex]
                    = new BorderCellRangeAddress(
                            base.FirstRow, base.LastRow, base.FirstColumn + 1, base.LastColumn - 1, this.SheetVersion);
            }
            //最終列Range
            if (NumberOfColumns > 2)
            {
                RetVal[LastRangeIndex]
                    = new BorderCellRangeAddress(
                            base.FirstRow, base.LastRow, base.LastColumn, base.LastColumn, this.SheetVersion);
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたセルが必要とする罫線情報(BorderInRangeフラグ)を取得
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns>BorderInRangeフラグ</returns>
        public List<BorderInRange> GetBorderInRange(int RowIndex, int ColumnIndex)
        {
            List<BorderInRange> RetVal = new List<BorderInRange>();
            //先頭行
            if (RowIndex == base.FirstRow)
            {
                //BorderTop : XlBordersIndex.xlEdgeTop
                RetVal.Add(BorderInRange.EdgeTop);
                //下に行がある場合
                if (RowIndex != LastRow)
                {
                    //BorderBottom : XlBordersIndex.xlInsideVertical
                    RetVal.Add(BorderInRange.InsideBottom);
                }
            }
            //最終行
            if (RowIndex == base.LastRow)
            {
                //BorderBottm : XlBordersIndex.xlEdgeBottom
                RetVal.Add(BorderInRange.EdgeBottom);
                //上に行がある場合
                if (RowIndex != FirstRow)
                {
                    //BorderTop : XlBordersIndex.xlInsideVertical
                    RetVal.Add(BorderInRange.InsideTop);
                }
            }
            //先頭列
            if (ColumnIndex == base.FirstColumn)
            {
                //BorderLeft : XlBordersIndex.xlEdgeLeft
                RetVal.Add(BorderInRange.EdgeLeft);
                //右に列がある場合
                if (ColumnIndex != LastColumn)
                {
                    //BorderRight : XlBordersIndex.xlInsideVertical
                    RetVal.Add(BorderInRange.InsideRight);
                }
            }
            //最終列
            if (ColumnIndex == base.LastColumn)
            {
                //BorderRight ; XlBordersIndex.xlEdgeRight
                RetVal.Add(BorderInRange.EdgeRight);
                //左に列がある場合
                if (ColumnIndex != FirstColumn)
                {
                    //BorderLeft :XlBordersIndex.xlInsideVertical
                    RetVal.Add(BorderInRange.InsideRight);
                }
            }
            //中間行
            if (RowIndex != base.FirstRow && RowIndex == base.LastRow)
            {
                //BorderTop : XlBordersIndex.xlInsideHorizontal
                RetVal.Add(BorderInRange.InsideTop);
                //BorderBottom  : XlBordersIndex.xlInsideHorizontal
                RetVal.Add(BorderInRange.InsideBottom);
            }
            //中間列
            if (RowIndex != base.FirstColumn && RowIndex == base.LastColumn)
            {
                //BorderLeft : XlBordersIndex.xlInsideVertical
                RetVal.Add(BorderInRange.InsideLeft);
                //BorderRight : XlBordersIndex.xlInsideVertical
                RetVal.Add(BorderInRange.InsideRight);
            }
            return RetVal;
        }
    }
}
