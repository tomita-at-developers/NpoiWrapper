using System;
using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Common
{
    public class ColumnReference
    {
        /// <summary>
        /// A形式の列アドレスを列インデックスに変換(1-base)
        /// </summary>
        /// <param name="Reference"></param>
        /// <returns></returns>
        public static int StringToIndex(string Reference)
        {
            return CellReference.ConvertColStringToIndex(Reference) + 1;
        }

        /// <summary>
        /// 列インデックス(1-base)をA形式の列アドレスに変換
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static string IndexToString(int Index)
        {
            if ( Index >= 1)
            {
                return CellReference.ConvertNumToColString(Index - 1);
            }
            else
            {
                throw new ArgumentException("Column index out of range.");
            }
        }
    }
}
