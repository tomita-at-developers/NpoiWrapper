namespace Developers.NpoiWrapper
{
    /// <summary>
    /// NpoiWrapperクラス
    /// Microsoft.Office.Interop.Excel.Applicationをエミュレート
    /// Workbooksインスタンスを持つのみのクラス
    /// </summary>
    public class NpoiApplication
    {
        /// <summary>
        /// Workbooksクラス
        /// </summary>
        public Workbooks Workbooks { get; } = new Workbooks();
    }
}
