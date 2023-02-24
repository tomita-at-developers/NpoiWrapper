using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Collections;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Sheetsクラス
    /// WorkboolクラスのAheetsメソッドのみが本クラスをコンストラクトする
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Worksheets : Sheets
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        internal Worksheets(Workbook ParentWorkbook)
            : base(ParentWorkbook, SheetType.Worksheet)
        {
        }

        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public override object Current
        {
            get
            {
                return new Worksheet(
                    ParentBook,
                    ParentBook.PoiBook.GetSheetAt(GetSheetIndexList(SheetTypes)[EnumSheetIndex]));
            }
        }
 
        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Index">シートインデクス(１開始)</param>
        /// <returns></returns>
        public override dynamic this[int Index]
        {
            get
            {
                List<int> WorksheetIndex = GetWorksheetIndexList();
                return new Worksheet(ParentBook, ParentBook.PoiBook.GetSheetAt(WorksheetIndex[Index-1]));
            }
        }

        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Name">シート名</param>
        /// <returns></returns>
        public override dynamic this[string Name]        {
            get
            {
                return new Worksheet(ParentBook, ParentBook.PoiBook.GetSheet(Name));
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
        /// このBookに含まれるWorksheetのIndexリストを取得する
        /// </summary>
        /// <returns>Indexリスト</returns>
        private List<int> GetWorksheetIndexList()
        {
            List<int> WorksheetIndex = new List<int>();
            for (int i = 0; i < ParentBook.PoiBook.NumberOfSheets; i++)
            {
                ISheet sheet = ParentBook.PoiBook.GetSheetAt(i);
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
    }
}
