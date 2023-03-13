using NPOI.SS;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper.Utils
{
    static public class RangeUtil
    {
        /// <summary>
        /// CellRangeAddressListに含まれるアドレスからA1形式の文字列アドレスリストを生成。
        /// </summary>
        /// <param name="RangeAddressList"></param>
        /// <returns></returns>
        static internal string CellRangeAddressListToString(CellRangeAddressList RangeAddressList)
        {
            string RetVal = "Range";
            for (int a = 0; a < RangeAddressList.CountRanges(); a++)
            {
                RetVal += "[" + a + "]" + RangeAddressList.GetCellRangeAddress(a).FormatAsString()
                            + "(" + RangeAddressList.GetCellRangeAddress(a).NumberOfCells + ") ";
            }
            RetVal = RetVal.Trim();
            return RetVal;
        }

        /// <summary>
        /// 行/列インデックス[-1]をそれが意味する最小/最大値に変換し、インデクスとして安全に利用できる対象CellRangeAddressListを生成する。
        /// とはいえ最大値の場合、メモリ不足の例外が発生する可能性が高い
        /// </summary>
        /// <param name="RangeAddressList">対象CellRangeAddressList</param>
        /// <param name="Version">IBookが示すSpreadsheetVersion(最大値取得用)</param>
        /// <returns>評価済Rangeアドレスリスト</returns>
        static public CellRangeAddressList CreateSafeCellRangeAddressList(CellRangeAddressList RangeAddressList, SpreadsheetVersion Version)
        {
            CellRangeAddressList RetVal = new CellRangeAddressList();
            for (int i = 0; i < RangeAddressList.CountRanges(); i++)
            {
                //安全化したアドレスを生成＆追記
                RetVal.AddCellRangeAddress(CreatetSafeCellRangeAddress(RangeAddressList.GetCellRangeAddress(i), Version));
            }
            return RetVal;
        }

        /// <summary>
        /// 行/列インデックス[-1]をそれが意味する最小/最大値に変換し、インデクスとして安全に利用できるCellRangeAddressを生成する。
        /// とはいえ最大値の場合、メモリ不足の例外が発生する可能性が高い
        /// </summary>
        /// <param name="RangeAddressList">対象CellRangeAddress</param>
        /// <param name="Version">IBookが示すSpreadsheetVersion(最大値取得用)</param>
        /// <returns>評価済Rangeアドレス</returns>
        static public CellRangeAddress CreatetSafeCellRangeAddress(CellRangeAddress RangeAddress, SpreadsheetVersion Version)
        {
            //参照を断った新しいアドレスを生成
            CellRangeAddress RetVal = RangeAddress.Copy();
            //0未満(-1)の場合は仕様上の最大、最小値を利用する
            if (RetVal.FirstRow < 0)
            {
                RetVal.FirstRow = 0;
            }
            if (RetVal.LastRow < 0)
            {
                RetVal.LastRow = Version.MaxRows - 1;
            }
            if (RetVal.FirstColumn < 0)
            {
                RetVal.FirstColumn = 0;
            }
            if (RetVal.LastColumn < 0)
            {
                RetVal.LastColumn = Version.MaxColumns - 1;
            }
            return RetVal;
        }

        /// <summary>
        /// 指定された基点アドレスを基点に絶対アドレスを持つアドレスリストを生成する。
        /// </summary>
        /// <param name="RangeAddressList">評価対象Rangeアドレスリスト</param>
        /// <param name="RelativeTo">評価基点アドレス</param>
        /// <returns>絶対アドレスリスト</returns>
        public static CellRangeAddressList CreateAbsoluteCellRangeAddressList(CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo)
        {

            CellRangeAddressList RetVal = new CellRangeAddressList();
            for (int i = 0; i < RangeAddressList.CountRanges(); i++)
            {
                //参照を断った新しいアドレスを生成
                CellRangeAddress Address = RangeAddressList.GetCellRangeAddress(i).Copy();
                //基点アドレスがあれば絶対アドレス計算
                if (RelativeTo != null)
                {
                    //実アドレスならば基点アドレスを加算して絶対アドレスを算出
                    //最小/最大アドレス[-1]なら相対てきにも絶対的にも最小/最大アドレスなのでそのままにしておく
                    //基点アドレスのFirstRow/FirstColumnが[-1]なら0に変換して加算に利用する
                    if (Address.FirstRow >= 0)
                    {
                        Address.FirstRow += RelativeTo.FirstRow > 0 ? RelativeTo.FirstRow : 0;
                    }
                    if (Address.LastRow >= 0)
                    {
                        Address.LastRow += RelativeTo.FirstRow > 0 ? RelativeTo.FirstRow : 0;
                    }
                    if (Address.FirstColumn >= 0)
                    {
                        Address.FirstColumn += RelativeTo.FirstColumn > 0 ? RelativeTo.FirstColumn : 0;
                    }
                    if (Address.LastColumn >= 0)
                    {
                        Address.LastColumn += RelativeTo.FirstColumn > 0 ? RelativeTo.FirstColumn : 0;
                    }
                }
                RetVal.AddCellRangeAddress(Address);
            }
            return RetVal;

        }

        /// <summary>
        /// CellRangeAddressListの全アドレスを統合し論理和的な唯一つのアドレスを持つリストを生成する
        /// </summary>
        /// <param name="RangeAddressList">処理対象のCellRangeAddressList</param>
        /// <returns>CellRangeAddressList</returns>
        static internal CellRangeAddressList CreateMergedAddressList(CellRangeAddressList RangeAddressList)
        {
            List<int> FirstRow = new List<int>();
            List<int> LastRow = new List<int>();
            List<int> FirstColumn= new List<int>();
            List<int> LastColumn = new List<int>();
            foreach (CellRangeAddress adr in RangeAddressList.CellRangeAddresses)
            {
                FirstRow.Add(adr.FirstRow);
                LastRow.Add(adr.LastRow);
                FirstColumn.Add(adr.FirstColumn);
                LastColumn.Add(adr.LastColumn);
            }
            return new CellRangeAddressList(
                            FirstRow.Min(),
                            LastRow.Min() < 0 ? -1 : LastRow.Max(),
                            FirstColumn.Min(),
                            LastColumn.Min() < 0 ? -1 : LastColumn.Max());
        }

        /// <summary>
        /// object二次元配列の生成(インデックスを１開始にする)
        /// </summary>
        /// <param name="RowCount">行数</param>
        /// <param name="ColumnCount">列数</param>
        /// <returns>生成された配列</returns>
        public static object[,] CreateCellArray(int RowCount, int ColumnCount)
        {
            int[] LowerBoundArray = { 1, 1 };
            int[] LengthArray = { RowCount, ColumnCount };
            object[,] dataArray = (object[,])Array.CreateInstance(typeof(object), LengthArray, LowerBoundArray);
            return dataArray;
        }

        /// <summary>
        /// 二次元配列の生成
        /// </summary>
        /// <param name="ElementType">要素の型</param>
        /// <param name="Lengths">要素数配列{一次元長, 二次元長}</param>
        /// <param name="ColumnCount">開始インデクス{一次元開始値, 二次元開始値}</param>
        /// <returns></returns>
        static internal dynamic CreateArrayInstance(Type ElementType, int[] Lengths, int[] LowerBounds)
        {
            dynamic dataArray = (object[,])Array.CreateInstance(ElementType, Lengths, LowerBounds);
            return dataArray;
        }
    }
}

