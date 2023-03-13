using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    public class Sheets : IEnumerable, IEnumerator
    {
        [Flags]
        public enum SheetType
        {
            None = 0,
            Worksheet = 1,
            ChartSheet = 2,
            DialogSheet = 4
        }

        internal Workbook ParentBook { get; private set; }
        protected SheetType SheetTypes { get; private set; }
        protected int EnumSheetIndex { get; set; } = -1;

        /// <summary>
        /// コンストラクタ(全種類のシート)
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        internal Sheets(Workbook ParentWorkbook)
            : this(ParentWorkbook, (SheetType.Worksheet | SheetType.ChartSheet | SheetType.DialogSheet))
        {
        }

        /// <summary>
        /// コンストラクタ(SheetTypesフラグで指定された種類のシート)
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        /// <param name="SheetTypes">対象とするシートのタイプ</param>
        internal Sheets(Workbook ParentWorkbook, SheetType SheetTypes)
        {
            ParentBook = ParentWorkbook;
            this.SheetTypes = SheetTypes;
        }

        /// <summary>
        /// GetEnumeratorの実装
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            Reset();
            return (IEnumerator)this;
        }
        /// <summary>
        /// IEnumerator.MoveNextの実装
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            bool RetVal = false;
            EnumSheetIndex += 1;
            if (EnumSheetIndex < GetSheetIndexList(SheetTypes).Count)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public virtual object Current
        {
            get
            {
                //とりあえずWorksheetを返す
                return new Worksheet(
                    ParentBook,
                    ParentBook.PoiBook.GetSheetAt(GetSheetIndexList(SheetTypes)[EnumSheetIndex]));
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumSheetIndex = -1;
        }

        /// <summary>
        /// インデクサ(Index指定)
        /// </summary>
        /// <param name="Index">シートIndex(１開始)</param>
        /// <returns></returns>
        public virtual dynamic this[int Index]
        {
            get
            {
                //とりあえずWorksheetを返す
                return new Worksheet(ParentBook, ParentBook.PoiBook.GetSheetAt(Index - 1));
            }
        }

        /// <summary>
        /// インデクサ(名前指定)
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public virtual dynamic this[string Name]
        {
            get
            {
                //とりあえずWorksheetを返す
                return new Worksheet(ParentBook, ParentBook.PoiBook.GetSheet(Name));
            }
        }

        /// <summary>
        /// シート数の取得
        /// </summary>
        public int Count
        {
            get
            {
                return GetSheetIndexList(SheetTypes).Count;
            }
        }

        /// <summary>
        /// シートの追加
        /// ★常に末尾に追加される。追加位置の指定はできない。
        /// </summary>
        /// <returns></returns>
        public Worksheet Add()
        {
            return new Worksheet(ParentBook, ParentBook.PoiBook.CreateSheet());
        }

        /// <summary>
        /// このBookに含まれるSheetのIndexリストを取得する
        /// </summary>
        /// <param name="SheetTypes">対象とするシートの種類</param>
        /// <returns>Indexリスト</returns>
        protected List<int> GetSheetIndexList(SheetType SheetTypes)
        {
            List<int> SheetIndex = new List<int>();
            for (int i = 0; i < ParentBook.PoiBook.NumberOfSheets; i++)
            {
                ISheet sheet = ParentBook.PoiBook.GetSheetAt(i);
                //ワークシートが指定されている場合
                if(SheetTypes.HasFlag(SheetType.Worksheet))
                {
                    //ワークシートの選別(ただしHSSFSheetは選別不能！)
                    if (sheet is HSSFSheet
                        || (sheet is XSSFSheet && !(sheet is XSSFChartSheet) && !(sheet is XSSFDialogsheet)))
                    {
                        SheetIndex.Add(i);
                    }
                }
                //チャートシートが指定されている場合
                if (SheetTypes.HasFlag(SheetType.ChartSheet))
                {
                    //ワークシートの選別(ただしHSSFSheetは選別不能！)
                    if (sheet is XSSFSheet && (sheet is XSSFChartSheet))
                    {
                        SheetIndex.Add(i);
                    }
                }
                //ダイアログシートが指定されている場合
                if (SheetTypes.HasFlag(SheetType.DialogSheet))
                {
                    //ワークシートの選別(ただしHSSFSheetは選別不能！)
                    if (sheet is XSSFSheet && (sheet is XSSFDialogsheet))
                    {
                        SheetIndex.Add(i);
                    }
                }
            }
            return SheetIndex;
        }
    }
}
