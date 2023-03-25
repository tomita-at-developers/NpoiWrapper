using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Developers.NpoiWrapper;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlDVType in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xldvtype
    //----------------------------------------------------------------------------------------------
    //public enum XlDVType
    //{
    //    xlValidateInputOnly,
    //    xlValidateWholeNumber,
    //    xlValidateDecimal,
    //    xlValidateList,
    //    xlValidateDate,
    //    xlValidateTime,
    //    xlValidateTextLength,
    //    xlValidateCustom
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public static class ValidationType
    //{
    //    public const int ANY = 0;
    //    public const int INTEGER = 1;
    //    public const int DECIMAL = 2;
    //    public const int LIST = 3;
    //    public const int DATE = 4;
    //    public const int TIME = 5;
    //    public const int TEXT_LENGTH = 6;
    //    public const int FORMULA = 7;
    //}

    /// <summary>
    /// データバリデーションの種類
    /// </summary>
    public enum XlDVType : int
    {
        xlValidateInputOnly = 0,
        xlValidateWholeNumber = 1,
        xlValidateDecimal = 2,
        xlValidateList = 3,
        xlValidateDate = 4,
        xlValidateTime = 5,
        xlValidateTextLength = 6,
        xlValidateCustom = 7,
    }
    /// <summary>
    /// XlDVTypeとNPOI.SS.UserModel.ValidationTypeの相互変換
    /// </summary>
    internal static class XlDVTypeParser
    {
        private static readonly Dictionary<XlDVType, int> _Map = new Dictionary<XlDVType, int>()
        {
            { XlDVType.xlValidateInputOnly,     ValidationType.ANY          },
            { XlDVType.xlValidateWholeNumber,   ValidationType.INTEGER      },
            { XlDVType.xlValidateDecimal,       ValidationType.DECIMAL      },
            { XlDVType.xlValidateList,          ValidationType.LIST         },
            { XlDVType.xlValidateDate,          ValidationType.DATE         },
            { XlDVType.xlValidateTime,          ValidationType.TIME         },
            { XlDVType.xlValidateTextLength,    ValidationType.TEXT_LENGTH  },
            { XlDVType.xlValidateCustom,        ValidationType.FORMULA      }
        };
        /// <summary>
        /// XlDVType値を指定してValidationType値を取得。
        /// </summary>
        /// <param name="XlValue">XlDVType値</param>
        /// <returns>ValidationType値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static int GetPoiValue(XlDVType XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlDVType.");
            }
        }
        /// <summary>
        /// ValidationType値を指定してXlDVType値を取得。
        /// </summary>
        /// <param name="PoiValue">ValidationType値</param>
        /// <returns>XlDVType値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static XlDVType GetXlValue(int PoiValue)
        {
            if (_Map.ContainsValue(PoiValue))
            {
                KeyValuePair<XlDVType, int> Pair = _Map.FirstOrDefault(c => c.Value == PoiValue);
                return Pair.Key;
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of NPOI.SS.UserModel.ValidationType.");
            }
        }
    }
}
