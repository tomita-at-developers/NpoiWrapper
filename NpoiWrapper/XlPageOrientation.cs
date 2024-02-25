using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlPageOrientation in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlpageorientation
    //----------------------------------------------------------------------------------------------
    //public enum XlPageOrientation
    //{
    //    xlLandscape = 2,
    //    xlPortrait = 1
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface IPrintSetup
    //{
    //    bool Landscape { get; set; }
    //}

    public enum XlPageOrientation
    {
        xlLandscape = 2,
        xlPortrait = 1
    }

    internal static class XlPageOrientationParser
    {
        private static readonly Dictionary<XlPageOrientation, bool> _Map = new Dictionary<XlPageOrientation, bool>()
        {
            { XlPageOrientation.xlLandscape,    true    },
            { XlPageOrientation.xlPortrait,     false   }
        };
        /// <summary>
        /// XlPageOrientation値を指定してLandscape値を取得。
        /// </summary>
        /// <param name="XlValue">XlPageOrientation</param>
        /// <returns>andscape</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static bool GetPoiValue(XlPageOrientation XlValue)
        {
            if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid value is spesified as a member of XlPageOrientation.");
            }
        }
    }

}
