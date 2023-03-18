using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace Developers.NpoiWrapper.Utils
{
    internal static class StyleUtil
    {
        /// <summary>
        /// Bookが持つCellStyleの全てに対し、そのStyleの利用セル数をカウントする。
        /// forでGetRow/GetColumnするよりforeachでrow in sheet, cell in rowのほうが高速だった。
        /// </summary>
        /// <param name="PoiBook"></param>
        /// <returns></returns>
        public static Dictionary<short, int> GetCellStyleUsage(IWorkbook PoiBook)
        {
            //リターン値
            Dictionary<short, int> RetVal = new Dictionary<short, int>();
            //全スタイルをDictionaryに列挙
            for (int i = 0; i < PoiBook.NumCellStyles; i++)
            {
                RetVal.Add(PoiBook.GetCellStyleAt(i).Index, 0);
            }
            //シートループ
            for (int SIdx = 0; SIdx < PoiBook.NumberOfSheets; SIdx++)
            {
                ISheet s = PoiBook.GetSheetAt(SIdx);
                //行ループ
                foreach (IRow row in s)
                {
                    //列ループ
                    foreach (NPOI.SS.UserModel.ICell cell in row)
                    {
                        short styleIndex = cell.CellStyle.Index;
                        if (RetVal.ContainsKey(styleIndex))
                        {
                            RetVal[styleIndex] += 1;
                        }
                        else
                        {
                            RetVal.Add(styleIndex, 1);
                        }
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// 定義済のNumberFormatのIndexととそれが持つ書式文字列をリストアップする。
        /// </summary>
        /// <param name="Book"></param>
        /// <returns></returns>
        public static SortedDictionary<short, string> GetNumberFormats(IWorkbook Book)
        {
            SortedDictionary<short, string> RetVal = new SortedDictionary<short, string>();
            if (Book is XSSFWorkbook xssfbook)
            {
                NPOI.XSSF.Model.StylesTable StyleTable = xssfbook.GetStylesSource();
                RetVal = (SortedDictionary<short, string>)(StyleTable.GetNumberFormats());
            }
            else if (Book is HSSFWorkbook hssfbook)
            {
                NPOI.HSSF.Model.InternalWorkbook internalbook = hssfbook.InternalWorkbook;
                List<NPOI.HSSF.Record.FormatRecord> Formats = internalbook.Formats;
                foreach (NPOI.HSSF.Record.FormatRecord r in Formats)
                {
                    RetVal.Add((short)r.IndexCode, r.FormatString);
                }
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたNumberFormat値から書式文字列を取得する
        /// </summary>
        /// <param name="Book"></param>
        /// <param name="NumberFormat"></param>
        /// <returns></returns>
        public static string GetNumberFormatString(IWorkbook Book, short NumberFormat)
        {
            string RetVal = string.Empty;
            SortedDictionary<short, string> Formats = GetNumberFormats(Book);
            if (Formats.ContainsKey(NumberFormat))
            {
                RetVal = Formats[NumberFormat];
            }
            return RetVal;
        }
    }
}
