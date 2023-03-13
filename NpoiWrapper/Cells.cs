﻿using Developers.NpoiWrapper.Utils;
using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Cellsクラス
    /// インデクサOverrideのためだけにあるクラス
    /// (Office.Interop.Excelには存在しないクラス)
    /// </summary>
    public class Cells : Range
    {
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

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

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        public override Range this[object RowIndex, object ColumnIndex]
        {
            get
            {
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                //インデックスはintであること
                if (RowIndex is int row && ColumnIndex is int column)
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
                return new Range(base.ParentSheet, RangeAddressList, base.RelativeTo);
            }
        }
    }
}