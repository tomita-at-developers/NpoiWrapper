namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Application interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Application : _Application, AppEvents_Event
    //{
    //}

    /// <summary>
    /// Applicationクラス
    /// 実体は_Applicationクラス。
    /// クラスと同名のプロパティApplication.Applicationを公開するためだけに強引なOverdirdeとCastをしている。
    /// </summary>
    public class Application : _Application
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Application(bool Use2003ColorIndex)
            : base(Use2003ColorIndex)
        {
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Application()
            : base()
        {
        }

        #endregion
    }
}
