using System.Collections;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Workbooksクラス
    /// Microsoft.Office.Interop.Excel.Workbooksをエミュレート
    /// NpoiWrapperクラスのプロパティとしてのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Workbooks : IEnumerable, IEnumerator
    {
        internal List<Workbook> Books { get; private set; } = new List<Workbook>();
        protected int EnumBookIndex { get; set; } = -1;

        /// <summary>
        /// コンストラクタ
        /// NoiWrapperのプロパティとしてのみコンストラクトされる
        /// </summary>
        internal Workbooks()
        {
            //なにもしない
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
            EnumBookIndex += 1;
            if (EnumBookIndex < Books.Count)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public object Current
        {
            get
            {
                return Books[EnumBookIndex];
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumBookIndex = -1;
        }

        /// <summary>
        /// 新規Excelブックの追加
        /// </summary>
        /// <param name="Excel97_2003">Excel97-2003形式で作成する場合true(省略時Excel2007以降形式)</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Add(bool Excel97_2003 = false)
        {
            Workbook Book = new Workbook(Excel97_2003);
            Books.Add(Book);
            return Book;
        }

        /// <summary>
        /// 既存Excelブックを開く
        /// </summary>
        /// <param name="FileNanme">フルパスファイ名</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Open(string FileNanme)
        {
            Workbook Book = new Workbook(FileNanme);
            Books.Add(Book);
            return Book;
        }

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Workbook this[int index]
        {
            get
            {
                return Books[index];
            }
        }
    }
}
