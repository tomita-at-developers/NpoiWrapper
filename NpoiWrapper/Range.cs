using Developers.NpoiWrapper.Styles;
using Developers.NpoiWrapper.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Rangeクラス
    /// 実体は_Rangeクラス。
    /// クラスと同名のプロパティRange.Rangeを公開するためだけに強引なOverdirdeとCastをしている。
    /// </summary>
    public class Range : _Range
    {
        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        internal Range(Worksheet ParentSheet, CellRangeAddressList RangeAddressList)
            : base(ParentSheet, RangeAddressList)
        {
        }

        /// <summary>
        /// コンストラクタ(Range.Range, Range.Cellsを生成する場合に使用｡)
        /// Rangeクラス内でしか利用しないのでprivateとしている。
        /// </summary>
        /// <param name="ParentSheet"></param>
        /// <param name="RangeAddressList">相対表現のアドレスリスト</param>
        /// <param name="RelativeTo">基点アドレス</param>
        internal Range(
                    Worksheet ParentSheet, CellRangeAddressList RangeAddressList,
                    CellRangeAddress RelativeTo)
            :base(ParentSheet, RangeAddressList, RelativeTo)
        {
        }

        /// <summary>
        /// コンストラクタ(Range.Rows, Range.Columnsを生成する場合に使用｡)
        /// </summary>
        /// <param name="ParentSheet"></param>
        /// <param name="RangeAddressList">相対表現のアドレスリスト</param>
        /// <param name="RelativeTo">基点アドレス</param>
        /// <param name="CountAs"></param>
        internal Range(
                    Worksheet ParentSheet, CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo,
                    CountType CountAs)
            : base(ParentSheet, RangeAddressList, RelativeTo, CountAs)
        {
        }
    }
}
