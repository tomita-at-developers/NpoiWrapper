using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Util
{
    static public class Uitl
    {
        /// <summary>
        /// RangeにペーストするためのCellクラス配列の生成
        /// IndexはExcel.Interopに倣って１開始
        /// </summary>
        /// <param name="RowCount">行数</param>
        /// <param name="ColumnCount">列数</param>
        /// <returns></returns>
        static public Cell[,] CreateRangePastObject(int RowCount, int ColumnCount)
        {
            //Cellクラス配列の生成
            int[] Lengths= new int[] { RowCount, ColumnCount };
            int[] LwerBounds = new int[] { 1, 1 };
            dynamic dataArray = CreateArrayInstance(typeof(Cell), Lengths, LwerBounds);
            //配列要素にCellクラスのインスタンスを詰める
            for(int r = 1; r <= RowCount; r++)
            {
                for (int c = 1; c <= ColumnCount; c++)
                {
                    dataArray[r, c] = new Cell();
                }
            }   
            return (Cell[,])dataArray;
        }

        /// <summary>
        /// A1形式のアドレス修正
        /// </summary>
        /// <param name="Row">行インデックス(１開始)</param>
        /// <param name="Column">列インデックス(１開始)</param>
        /// <returns></returns>
        static public string CreateAddress(int Row, int Column)
        {
            CellAddress Adr = new CellAddress(Row - 1, Column - 1);
            return Adr.FormatAsString();
        }

        /// <summary>
        /// A1:B1形式のアドレス修正
        /// </summary>
        /// <param name="FirstRow">開始行インデックス(１開始)</param>
        /// <param name="LastRow">終了行インデックス(１開始)</param>
        /// <param name="FirstColumn">開始列インデックス(１開始)</param>
        /// <param name="LastColumn">終了列インデックス(１開始)</param>
        /// <returns></returns>
        static public string CreateRangeAddress(int FirstRow, int LastRow, int FirstColumn, int LastColumn)
        {
            CellRangeAddress Adr = new CellRangeAddress(FirstRow - 1, LastRow - 1, FirstColumn - 1, LastColumn - 1);
            return Adr.FormatAsString();
        }

        /// <summary>
        /// 二次元配列の生成
        /// </summary>
        /// <param name="ElementType">要素の型</param>
        /// <param name="Lengths">要素数配列{一次元長, 二次元長}</param>
        /// <param name="ColumnCount">開始インデクス{一次元開始値, 二次元開始値}</param>
        /// <returns></returns>
        static private dynamic CreateArrayInstance(Type ElementType, int[] Lengths, int[] LowerBounds)
        {
            dynamic dataArray = (object[,])Array.CreateInstance(ElementType, Lengths, LowerBounds);
            return dataArray;
        }

    }

}

