using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Workbooksクラス
    /// Microsoft.Office.Interop.Excel.Workbooksをエミュレート
    /// NpoiWrapperクラスのプロパティとしてのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Workbooks
    {
        /// <summary>
        /// コンストラクタ
        /// NoiWrapperのプロパティとしてのみコンストラクトされる
        /// </summary>
        internal Workbooks()
        {
            //なにもしない
        }

        /// <summary>
        /// 新規Excelブックの追加
        /// </summary>
        /// <param name="Excel97_2003">Excel97-2003形式で作成する場合true(省略時Excel2007以降形式)</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Add(bool Excel97_2003 = false)
        {
            return new Workbook(Excel97_2003);
        }

        /// <summary>
        /// 既存Excelブックを開く
        /// </summary>
        /// <param name="FileNanme">フルパスファイ名</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Open(string FileNanme)
        {
            return new Workbook(FileNanme);
        }
    }
}
