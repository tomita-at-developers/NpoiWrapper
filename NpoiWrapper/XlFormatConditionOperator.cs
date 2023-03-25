using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlFormatConditionOperator in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditionoperator
    //----------------------------------------------------------------------------------------------
    //public enum XlFormatConditionOperator
    //{
    //    xlBetween = 1,
    //    xlNotBetween,
    //    xlEqual,
    //    xlNotEqual,
    //    xlGreater,
    //    xlLess,
    //    xlGreaterEqual,
    //    xlLessEqual
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public static class OperatorType
    //{
    //    public const int BETWEEN = 0;
    //    public const int NOT_BETWEEN = 1;
    //    public const int EQUAL = 2;
    //    public const int NOT_EQUAL = 3;
    //    public const int GREATER_THAN = 4;
    //    public const int LESS_THAN = 5;
    //    public const int GREATER_OR_EQUAL = 6;
    //    public const int LESS_OR_EQUAL = 7;
    //    public const int IGNORED = 0;
    //}

    /// <summary>
    /// 評価条件。ValidationのOperationに指定できる。
    /// </summary>
    public enum XlFormatConditionOperator : int
    {
        xlBetween = 1,
        xlNotBetween = 2,
        xlEqual = 3,
        xlNotEqual = 4,
        xlGreater = 5,
        xlLess = 6,
        xlGreaterEqual = 7,
        xlLessEqual = 8,
    }

    /// <summary>
    /// XlFormatConditionOperatorParserとNPOI.SS.UserModel.OperatorTypeの相互変換
    /// </summary>
    internal static class XlFormatConditionOperatorParser
    {
        private static readonly Dictionary<XlFormatConditionOperator, int> _Map = new Dictionary<XlFormatConditionOperator, int>()
        {
            { XlFormatConditionOperator.xlBetween,      OperatorType.BETWEEN            },
            { XlFormatConditionOperator.xlNotBetween,   OperatorType.NOT_BETWEEN        },
            { XlFormatConditionOperator.xlEqual,        OperatorType.EQUAL              },
            { XlFormatConditionOperator.xlNotEqual,     OperatorType.NOT_EQUAL          },
            { XlFormatConditionOperator.xlGreater,      OperatorType.GREATER_THAN       },
            { XlFormatConditionOperator.xlLess,         OperatorType.LESS_THAN          },
            { XlFormatConditionOperator.xlGreaterEqual, OperatorType.GREATER_OR_EQUAL   },
            { XlFormatConditionOperator.xlLessEqual,    OperatorType.LESS_OR_EQUAL      }
        };
        /// <summary>
        /// XlFormatConditionOperator値を指定してOperatorType値を取得。
        /// </summary>
        /// <param name="XlValue">XlFormatConditionOperator</param>
        /// <returns>OperatorType値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static int GetPoiValue(XlFormatConditionOperator XlValue)
        {
            if (XlValue == 0)
            {
                return OperatorType.IGNORED;
            }
            else if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlFormatConditionOperator.");
            }
        }
        /// <summary>
        /// OperatorType値を指定してXlFormatConditionOperator値を取得。
        /// </summary>
        /// <param name="PoiValue">OperatorType値</param>
        /// <returns>XlFormatConditionOperator値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlFormatConditionOperator GetXlValue(int PoiValue)
        {
            if (PoiValue == OperatorType.IGNORED)
            {
                return 0;
            }
            else if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlFormatConditionOperator, int> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of NPOI.SS.UserModel.OperatorType.");
            }
        }
    }
}