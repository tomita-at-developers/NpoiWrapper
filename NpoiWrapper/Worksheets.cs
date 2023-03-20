using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Collections.Generic;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Worksheets interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Worksheets
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
    //    void FillAcrossSheets(Range Range, XlFillWith Type = XlFillWith.xlFillWithAll);
    //    void Move([Optional] object Before, [Optional] object After);
    //    IEnumerator GetEnumerator();
    //    void _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate);
    //    void PrintPreview([Optional] object EnableChanges);
    //    void Select([Optional] object Replace);
    //    void PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void PrintOutEx([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName, [Optional] object IgnorePrintAreas);
    //}

    /// <summary>
    /// Worksheetsクラス
    /// Workbookクラスコンストラクト時にWorkbook.Worksheetsとしてコンストラクトされる。
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// Workbook.WorksheetsはSheets型なのでSheetsを継承しているが、Interop.ExcelではインターフェイスとしてのSheetsとWorksheetsに継承関係はない。
    /// Worksheets実装段階ではSheetsを継承しているかも知れない。
    /// </summary>
    public class Worksheets : Sheets
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        internal Worksheets(Workbook ParentWorkbook)
            : base(ParentWorkbook, SheetType.Worksheet)
        {
        }

        #endregion

        #region "interface implementations"

        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public override object Current
        {
            get
            {
                return new Worksheet(
                    Parent,
                    Parent.PoiBook.GetSheetAt(GetSheetIndexList(SheetTypes)[EnumSheetIndex]));
            }
        }

        #endregion

        #region "methods"

        #region "private methods"

        /// <summary>
        /// このBookに含まれるWorksheetのIndexリストを取得する
        /// </summary>
        /// <returns>Indexリスト</returns>
        private List<int> GetWorksheetIndexList()
        {
            List<int> WorksheetIndex = new List<int>();
            for (int i = 0; i < Parent.PoiBook.NumberOfSheets; i++)
            {
                ISheet sheet = Parent.PoiBook.GetSheetAt(i);
                //ワークシートの選別(ただしHSSFSheetは選別不能！)
                if (sheet is HSSFSheet
                    || (sheet is XSSFSheet && !(sheet is XSSFChartSheet))
                    || (sheet is XSSFSheet && !(sheet is XSSFDialogsheet)))
                {
                    WorksheetIndex.Add(i);
                }
            }
            return WorksheetIndex;
        }

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Index">シートインデクス(１開始)</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public override dynamic this[int Index]
        {
            get
            {
                List<int> WorksheetIndex = GetWorksheetIndexList();
                return new Worksheet(Parent, Parent.PoiBook.GetSheetAt(WorksheetIndex[Index - 1]));
            }
        }

        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Name">シート名</param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public override dynamic this[string Name]
        {
            get
            {
                return new Worksheet(Parent, Parent.PoiBook.GetSheet(Name));
            }
        }

        #endregion
    }
}
