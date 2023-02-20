using NPOI.SS.Formula;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// NpoiWrapperクラス
    /// Microsoft.Office.Interop.Excel.Applicationをエミュレート
    /// </summary>
    public class NpoiWrapper
    {
        /// <summary>
        /// Workbooksクラス
        /// </summary>
        public Workbooks Workbooks { get; } = new Workbooks();

    }
}
