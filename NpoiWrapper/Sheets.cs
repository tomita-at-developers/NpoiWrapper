using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Sheetsクラス
    /// WorkboolクラスのAheetsメソッドのみが本クラスをコンストラクトする
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Sheets
    {
        internal Workbook ParentBook { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        internal Sheets(Workbook ParentWorkbook) 
        {
            ParentBook = ParentWorkbook;
        }

        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Index">シートインデクス(１開始)</param>
        /// <returns></returns>
        public Worksheet this[int Index]
        {
            get
            {
                return new Worksheet(ParentBook, ParentBook.PoiBook.GetSheetAt(Index));
            }
        }

        /// <summary>
        /// インデクザ
        /// </summary>
        /// <param name="Name">シート名</param>
        /// <returns></returns>
        public Worksheet this[string Name]        {
            get
            {
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
                return ParentBook.PoiBook.NumberOfSheets; 
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
    }
}
