using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Sheets interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Sheets : IEnumerable
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    int Count { get; }
    //    object Item { get; }
    //    HPageBreaks HPageBreaks { get; }
    //    VPageBreaks VPageBreaks { get; }
    //    object Visible { get; set; }
    //    [IndexerName("_Default")]
    //    object this[object Index] { get; }
    //    object Add([Optional] object Before, [Optional] object After, [Optional] object Count, [Optional] object Type);
    //    void Copy([Optional] object Before, [Optional] object After);
    //    void Delete();
    //    void FillAcrossSheets([MarshalAs(UnmanagedType.Interface)] Range Range, XlFillWith Type = XlFillWith.xlFillWithAll);
    //    void Move([Optional] object Before, [Optional] object After);
    //    new IEnumerator GetEnumerator();
    //    void _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate);
    //    void PrintPreview([Optional] object EnableChanges);
    //    void Select([Optional] object Replace);
    //    void PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void PrintOutEx([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName, [Optional] object IgnorePrintAreas);
    //}

    /// <summary>
    /// Sheetsクラス
    /// Workbookクラスコンストラクト時にWorkbook.Sheetsとしてコンストラクトされる。
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Sheets : IEnumerable, IEnumerator
    {
        #region "fields"

        /// <summary>
        /// シートのタイプを示すフラグ
        /// </summary>
        protected SheetType SheetTypes;

        /// <summary>
        /// IEnumerator用列インデクス
        /// </summary>
        protected int EnumSheetIndex = -1;

        #endregion

        #region "constructors"

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
            Parent = ParentWorkbook;
            this.SheetTypes = SheetTypes;
        }

        #endregion

        #region "enums"

        /// <summary>
        /// 管理対象とするシートの種類を示すフラグ
        /// </summary>
        [Flags]
        public enum SheetType
        {
            /// <summary>
            /// なし
            /// </summary>
            None = 0,
            /// <summary>
            /// ワークシート
            /// </summary>
            Worksheet = 1,
            /// <summary>
            /// チャートシート(グラフシート)
            /// </summary>
            ChartSheet = 2,
            /// <summary>
            /// ダイアログシート
            /// </summary>
            DialogSheet = 4
        }

        #endregion

        #region "interface implementations"

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
                    Parent,
                    Parent.PoiBook.GetSheetAt(GetSheetIndexList(SheetTypes)[EnumSheetIndex]));
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumSheetIndex = -1;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Workbook Parent { get; private set; }

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

        #endregion

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// シートの追加(常に末尾に追加)
        /// </summary>
        /// <param name="Before">このシートの前に追加(無視されます)</param>
        /// <param name="After">このシートの後に追加(無視されます)</param>
        /// <param name="Count">追加するシートの数(無視されます)</param>
        /// <param name="Type">ワークシートの種類(無視されます)</param>
        /// <returns></returns>
        public Worksheet Add(object Before = null, object After = null, object Count = null, object Type = null)
        {
            return new Worksheet(Parent, Parent.PoiBook.CreateSheet());
        }

        #endregion

        #region "protected methods"

        /// <summary>
        /// このBookに含まれるSheetのIndexリストを取得する
        /// </summary>
        /// <param name="SheetTypes">対象とするシートの種類</param>
        /// <returns>Indexリスト</returns>
        protected List<int> GetSheetIndexList(SheetType SheetTypes)
        {
            List<int> SheetIndex = new List<int>();
            for (int i = 0; i < Parent.PoiBook.NumberOfSheets; i++)
            {
                ISheet sheet = Parent.PoiBook.GetSheetAt(i);
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

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサ(Index指定)
        /// </summary>
        /// <param name="Index">シートIndex(１開始)</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public virtual dynamic this[int Index]
        {
            get
            {
                //とりあえずWorksheetを返す
                return new Worksheet(Parent, Parent.PoiBook.GetSheetAt(Index - 1));
            }
        }

        /// <summary>
        /// インデクサ(名前指定)
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public virtual dynamic this[string Name]
        {
            get
            {
                //とりあえずWorksheetを返す
                return new Worksheet(Parent, Parent.PoiBook.GetSheet(Name));
            }
        }

        #endregion
    }
}
