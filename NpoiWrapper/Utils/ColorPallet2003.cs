using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper.Utils
{
    internal static class ColorPallet2003
    {
        #region "fields"

        private static readonly Dictionary<int, short> Excel2003ColorMap;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        static ColorPallet2003()
        {
            Excel2003ColorMap = new Dictionary<int, short>()
            {
                { 1, 8 },
                { 2, 9 },
                { 3, 10 },
                { 4, 11 },
                { 5, 12 },
                { 6, 13 },
                { 7, 14 },
                { 8, 15 },
                { 9, 16 },
                { 10, 17 },
                { 11, 18 },
                { 12, 19 },
                { 13, 20 },
                { 14, 21 },
                { 15, 22 },
                { 16, 23 },
                { 17, 24 },
                { 18, 61 },
                { 19, 26 },
                { 20, 41 },
                { 21, 28 },
                { 22, 29 },
                { 23, 30 },
                { 24, 31 },
                { 25, 18 },
                { 26, 14 },
                { 27, 13 },
                { 28, 15 },
                { 29, 20 },
                { 30, 16 },
                { 31, 21 },
                { 32, 12 },
                { 33, 40 },
                { 34, 41 },
                { 35, 42 },
                { 36, 43 },
                { 37, 44 },
                { 38, 45 },
                { 39, 46 },
                { 40, 47 },
                { 41, 48 },
                { 42, 49 },
                { 43, 50 },
                { 44, 51 },
                { 45, 52 },
                { 46, 53 },
                { 47, 54 },
                { 48, 55 },
                { 49, 56 },
                { 50, 57 },
                { 51, 58 },
                { 52, 59 },
                { 53, 60 },
                { 54, 61 },
                { 55, 62 },
                { 56, 63 },
            };
        }

        #endregion

        #region "methods"

        /// <summary>
        /// Excel2003のColorIndex値からIndecedColorsの該当色Index値を取得する。
        /// パレットがカスタマイズされている場合は必ずしも正しい値が取得できるとは限らない。
        /// </summary>
        /// <param name="Excel2003Index"></param>
        /// <returns></returns>
        public static short GetPoiIndex(int Excel2003Index)
        {
            short RetVal = IndexedColors.Automatic.Index;
            if (Excel2003ColorMap.ContainsKey(Excel2003Index))
            {
                RetVal = Excel2003ColorMap[Excel2003Index];
            }
            return RetVal;
        }

        /// <summary>
        /// IndecedColorsのIndex値からExcel2003のColorIndex値を取得する。
        /// 一対多なので必ずしももとには戻らない。
        /// </summary>
        /// <param name="PoiIndex"></param>
        /// <returns>Excel2003の標準ColorIndex値</returns>
        public static int GetExcel2003Index(short PoiIndex)
        {
            //xlColorIndexAutomatic = -4105
            int RetVal = (int)XlColorIndex.xlColorIndexAutomatic;
            List<KeyValuePair<int, short>> Matches = Excel2003ColorMap.Where(m => m.Value == PoiIndex).ToList();
            if (Matches.Count > 0)
            {
                RetVal = Matches[0].Key;
            }
            return RetVal;
        }

        #endregion
    }
}
