using Developers.NpoiWrapper.Utils;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Rowsクラス
    /// インデクサOverrideのためだけにあるクラス
    /// (Office.Interop.Excelには存在しないクラス)
    /// </summary>
    public class Rows : Range
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        #endregion

        #region "constructors"

        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        /// <param name="RelativeTo">アドレスの開始位置を示すCellRangeAddressインスタンス</param>
        internal Rows(Worksheet ParentSheet, CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo = null)
            : base(ParentSheet, RangeAddressList, RelativeTo, CountType.Rows)
        {
            Logger.Debug(RangeUtil.CellRangeAddressListToString(RangeAddressList));
            //なにもしない
        }

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="RowIndex">行オフセット(先頭アドレスの先頭行に対する１開始のオフセット)を数字で指定</param>
        /// <param name="Ignroed">無視される</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public override Range this[object RowIndex = null, object Ignroed = null]
        {
            get
            {
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                //インデックスはintであること
                if (base.TryConvertToInt(RowIndex, out int Index))
                {
                    //１から始まるIndexであること
                    if (Index > 0)
                    {
                        //絶対行インデックス生成
                        int TargetIndex = this.SafeAddressList.GetCellRangeAddress(0).FirstRow + (Index - 1);
                        //絶対行インデックスがシート範囲内であること
                        if (TargetIndex <= this.Parent.Parent.PoiBook.SpreadsheetVersion.MaxRows - 1)
                        {
                            //先頭生アドレス取得
                            CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(0).Copy();
                            //指定された１行を選択
                            RawAddress.FirstRow = (RawAddress.FirstRow  < 0 ? 0 : RawAddress.FirstRow) +(Index - 1);
                            RawAddress.LastRow = RawAddress.FirstRow;
                            //アドレスリストに追加
                            RangeAddressList.AddCellRangeAddress(RawAddress);
                        }
                        else
                        {
                            throw new ArgumentException("RowIndex is out of range.");
                        }
                    }
                    else
                    {
                        throw new ArgumentException("RowIndex should be 1-origined integer.");
                    }
                }
                //上記以外は例外スロー
                else
                {
                    throw new ArgumentException("Type of RowIndex should be int.");
                }
                //Rangeクラスインスタンス生成
                return new Range(base.Parent, RangeAddressList, base.RelativeTo, CountType.Rows);
            }
        }

        #endregion
    }
}
