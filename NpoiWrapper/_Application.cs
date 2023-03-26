using System;
using System.Collections.Generic;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // _Application interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface _Application
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    Application Parent { get; }
    //    Range ActiveCell { get; }
    //    Chart ActiveChart { get; }
    //    DialogSheet ActiveDialog { get; }
    //    MenuBar ActiveMenuBar { get; }
    //    string ActivePrinter { get; set; }
    //    object ActiveSheet { get; }
    //    Window ActiveWindow { get; }
    //    Workbook ActiveWorkbook { get; }
    //    AddIns AddIns { get; }
    //    Assistant Assistant { get; }
    //    Range Cells { get; }
    //    Sheets Charts { get; }
    //    Range Columns { get; }
    //    CommandBars CommandBars { get; }
    //    int DDEAppReturnCode { get; }
    //    Sheets DialogSheets { get; }
    //    MenuBars MenuBars { get; }
    //    Modules Modules { get; }
    //    Names Names { get; }
    //    Range Range { get; }
    //    Range Rows { get; }
    //    object Selection { get; }
    //    Sheets Sheets { get; }
    //    Menu ShortcutMenus { get; }
    //    Workbook ThisWorkbook { get; }
    //    Toolbars Toolbars { get; }
    //    Windows Windows { get; }
    //    Workbooks Workbooks { get; }
    //    WorksheetFunction WorksheetFunction { get; }
    //    Sheets Worksheets { get; }
    //    Sheets Excel4IntlMacroSheets { get; }
    //    Sheets Excel4MacroSheets { get; }
    //    bool AlertBeforeOverwriting { get; set; }
    //    string AltStartupPath { get; set; }
    //    bool AskToUpdateLinks { get; set; }
    //    bool EnableAnimations { get; set; }
    //    AutoCorrect AutoCorrect { get; }
    //    int Build { get; }
    //    bool CalculateBeforeSave { get; set; }
    //    XlCalculation Calculation { get; set; }
    //    object Caller { get; }
    //    bool CanPlaySounds { get; }
    //    bool CanRecordSounds { get; }
    //    string Caption { get; set; }
    //    bool CellDragAndDrop { get; set; }
    //    object ClipboardFormats { get; }
    //    bool DisplayClipboardWindow { get; set; }
    //    bool ColorButtons { get; set; }
    //    XlCommandUnderlines CommandUnderlines { get; set; }
    //    bool ConstrainNumeric { get; set; }
    //    bool CopyObjectsWithCells { get; set; }
    //    XlMousePointer Cursor { get; set; }
    //    int CustomListCount { get; }
    //    XlCutCopyMode CutCopyMode { get; set; }
    //    int DataEntryMode { get; set; }
    //    string _Default { get; }
    //    string DefaultFilePath { get; set; }
    //    Dialogs Dialogs { get; }
    //    bool DisplayAlerts { get; set; }
    //    bool DisplayFormulaBar { get; set; }
    //    bool DisplayFullScreen { get; set; }
    //    bool DisplayNoteIndicator { get; set; }
    //    XlCommentDisplayMode DisplayCommentIndicator { get; set; }
    //    bool DisplayExcel4Menus { get; set; }
    //    bool DisplayRecentFiles { get; set; }
    //    bool DisplayScrollBars { get; set; }
    //    bool DisplayStatusBar { get; set; }
    //    bool EditDirectlyInCell { get; set; }
    //    bool EnableAutoComplete { get; set; }
    //    XlEnableCancelKey EnableCancelKey { get; set; }
    //    bool EnableSound { get; set; }
    //    bool EnableTipWizard { get; set; }
    //    object FileConverters { get; }
    //    FileSearch FileSearch { get; }
    //    IFind FileFind { get; }
    //    bool FixedDecimal { get; set; }
    //    int FixedDecimalPlaces { get; set; }
    //    double Height { get; set; }
    //    bool IgnoreRemoteRequests { get; set; }
    //    bool Interactive { get; set; }
    //    object International { get; }
    //    bool Iteration { get; set; }
    //    bool LargeButtons { get; set; }
    //    double Left { get; set; }
    //    string LibraryPath { get; }
    //    object MailSession { get; }
    //    XlMailSystem MailSystem { get; }
    //    bool MathCoprocessorAvailable { get; }
    //    double MaxChange { get; set; }
    //    int MaxIterations { get; set; }
    //    int MemoryFree { get; }
    //    int MemoryTotal { get; }
    //    int MemoryUsed { get; }
    //    bool MouseAvailable { get; }
    //    bool MoveAfterReturn { get; set; }
    //    XlDirection MoveAfterReturnDirection { get; set; }
    //    RecentFiles RecentFiles { get; }
    //    string Name { get; }
    //    string NetworkTemplatesPath { get; }
    //    ODBCErrors ODBCErrors { get; }
    //    int ODBCTimeout { get; set; }
    //    string OnCalculate { get; set; }
    //    string OnData { get; set; }
    //    string OnDoubleClick { get; set; }
    //    string OnEntry { get; set; }
    //    string OnSheetActivate { get; set; }
    //    string OnSheetDeactivate { get; set; }
    //    string OnWindow { get; set; }
    //    string OperatingSystem { get; }
    //    string OrganizationName { get; }
    //    string Path { get; }
    //    string PathSeparator { get; }
    //    object PreviousSelections { get; }
    //    bool PivotTableSelection { get; set; }
    //    bool PromptForSummaryInfo { get; set; }
    //    bool RecordRelative { get; }
    //    XlReferenceStyle ReferenceStyle { get; set; }
    //    object RegisteredFunctions { get; }
    //    bool RollZoom { get; set; }
    //    bool ScreenUpdating { get; set; }
    //    int SheetsInNewWorkbook { get; set; }
    //    bool ShowChartTipNames { get; set; }
    //    bool ShowChartTipValues { get; set; }
    //    string StandardFont { get; set; }
    //    double StandardFontSize { get; set; }
    //    string StartupPath { get; }
    //    object StatusBar { get; set; }
    //    string TemplatesPath { get; }
    //    bool ShowToolTips { get; set; }
    //    double Top { get; set; }
    //    XlFileFormat DefaultSaveFormat { get; set; }
    //    string TransitionMenuKey { get; set; }
    //    int TransitionMenuKeyAction { get; set; }
    //    bool TransitionNavigKeys { get; set; }
    //    double UsableHeight { get; }
    //    double UsableWidth { get; }
    //    bool UserControl { get; set; }
    //    string UserName { get; set; }
    //    string Value { get; }
    //    VBE VBE { get; }
    //    string Version { get; }
    //    bool Visible { get; set; }
    //    double Width { get; set; }
    //    bool WindowsForPens { get; }
    //    XlWindowState WindowState { get; set; }
    //    int UILanguage { get; set; }
    //    int DefaultSheetDirection { get; set; }
    //    int CursorMovement { get; set; }
    //    bool ControlCharacters { get; set; }
    //    bool EnableEvents { get; set; }
    //    bool DisplayInfoWindow { get; set; }
    //    bool ExtendList { get; set; }
    //    OLEDBErrors OLEDBErrors { get; }
    //    COMAddIns COMAddIns { get; }
    //    DefaultWebOptions DefaultWebOptions { get; }
    //    string ProductCode { get; }
    //    string UserLibraryPath { get; }
    //    bool AutoPercentEntry { get; set; }
    //    LanguageSettings LanguageSettings { get; }
    //    object Dummy101 { get; }
    //    AnswerWizard AnswerWizard { get; }
    //    int CalculationVersion { get; }
    //    bool ShowWindowsInTaskbar { get; set; }
    //    MsoFeatureInstall FeatureInstall { get; set; }
    //    bool Ready { get; }
    //    CellFormat FindFormat { get; set; }
    //    CellFormat ReplaceFormat { get; set; }
    //    UsedObjects UsedObjects { get; }
    //    XlCalculationState CalculationState { get; }
    //    XlCalculationInterruptKey CalculationInterruptKey { get; set; }
    //    Watches Watches { get; }
    //    bool DisplayFunctionToolTips { get; set; }
    //    MsoAutomationSecurity AutomationSecurity { get; set; }
    //    FileDialog FileDialog { get; }
    //    bool DisplayPasteOptions { get; set; }
    //    bool DisplayInsertOptions { get; set; }
    //    bool GenerateGetPivotData { get; set; }
    //    AutoRecover AutoRecover { get; }
    //    int Hwnd { get; }
    //    int Hinstance { get; }
    //    ErrorCheckingOptions ErrorCheckingOptions { get; }
    //    bool AutoFormatAsYouTypeReplaceHyperlinks { get; set; }
    //    SmartTagRecognizers SmartTagRecognizers { get; }
    //    NewFile NewWorkbook { get; }
    //    SpellingOptions SpellingOptions { get; }
    //    Speech Speech { get; }
    //    bool MapPaperSize { get; set; }
    //    bool ShowStartupDialog { get; set; }
    //    string DecimalSeparator { get; set; }
    //    string ThousandsSeparator { get; set; }
    //    bool UseSystemSeparators { get; set; }
    //    Range ThisCell { get; }
    //    RTD RTD { get; }
    //    bool DisplayDocumentActionTaskPane { get; set; }
    //    bool ArbitraryXMLSupportAvailable { get; }
    //    int MeasurementUnit { get; set; }
    //    bool ShowSelectionFloaties { get; set; }
    //    bool ShowMenuFloaties { get; set; }
    //    bool ShowDevTools { get; set; }
    //    bool EnableLivePreview { get; set; }
    //    bool DisplayDocumentInformationPanel { get; set; }
    //    bool AlwaysUseClearType { get; set; }
    //    bool WarnOnFunctionNameConflict { get; set; }
    //    int FormulaBarHeight { get; set; }
    //    bool DisplayFormulaAutoComplete { get; set; }
    //    XlGenerateTableRefs GenerateTableRefs { get; set; }
    //    IAssistance Assistance { get; }
    //    bool EnableLargeOperationAlert { get; set; }
    //    int LargeOperationCellThousandCount { get; set; }
    //    bool DeferAsyncQueries { get; set; }
    //    MultiThreadedCalculation MultiThreadedCalculation { get; }
    //    int ActiveEncryptionSession { get; }
    //    bool HighQualityModeForGraphics { get; set; }
    //    void Calculate();
    //    void DDEExecute(int Channel, string String);
    //    int DDEInitiate(string App, string Topic);
    //    void DDEPoke(int Channel, object Item, object Data);
    //    object DDERequest(int Channel, string Item);
    //    void DDETerminate(int Channel);
    //    object Evaluate(object Name);
    //    object _Evaluate(object Name);
    //    object ExecuteExcel4Macro(string String);
    //    Range Intersect(Range Arg1, Range Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    object Run([Optional] object Macro, [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    object _Run2([Optional] object Macro, [Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    void SendKeys(object Keys, [Optional] object Wait);
    //    Range Union(Range Arg1, Range Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    void ActivateMicrosoftApp(XlMSApplication Index);
    //    void AddChartAutoFormat(object Chart, string Name, [Optional] object Description);
    //    void AddCustomList(object ListArray, [Optional] object ByRow);
    //    double CentimetersToPoints(double Centimeters);
    //    bool CheckSpelling(string Word, [Optional] object CustomDictionary, [Optional] object IgnoreUppercase);
    //    object ConvertFormula(object Formula, XlReferenceStyle FromReferenceStyle, [Optional] object ToReferenceStyle, [Optional] object ToAbsolute, [Optional] object RelativeTo);
    //    object Dummy1([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4);
    //    object Dummy2([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8);
    //    object Dummy3();
    //    object Dummy4([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15);
    //    object Dummy5([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13);
    //    object Dummy6();
    //    object Dummy7();
    //    object Dummy8([Optional] object Arg1);
    //    object Dummy9();
    //    bool Dummy10([Optional] object arg);
    //    void Dummy11();
    //    void DeleteChartAutoFormat(string Name);
    //    void DeleteCustomList(int ListNum);
    //    void DoubleClick();
    //    void _FindFile();
    //    object GetCustomListContents(int ListNum);
    //    int GetCustomListNum(object ListArray);
    //    object GetOpenFilename([Optional] object FileFilter, [Optional] object FilterIndex, [Optional] object Title, [Optional] object ButtonText, [Optional] object MultiSelect);
    //    object GetSaveAsFilename([Optional] object InitialFilename, [Optional] object FileFilter, [Optional] object FilterIndex, [Optional] object Title, [Optional] object ButtonText);
    //    void Goto([Optional] object Reference, [Optional] object Scroll);
    //    void Help([Optional] object HelpFile, [Optional] object HelpContextID);
    //    double InchesToPoints(double Inches);
    //    object InputBox(string Prompt, [Optional] object Title, [Optional] object Default, [Optional] object Left, [Optional] object Top, [Optional] object HelpFile, [Optional] object HelpContextID, [Optional] object Type);
    //    void MacroOptions([Optional] object Macro, [Optional] object Description, [Optional] object HasMenu, [Optional] object MenuText, [Optional] object HasShortcutKey, [Optional] object ShortcutKey, [Optional] object Category, [Optional] object StatusBar, [Optional] object HelpContextID, [Optional] object HelpFile);
    //    void MailLogoff();
    //    void MailLogon([Optional] object Name, [Optional] object Password, [Optional] object DownloadNewMail);
    //    Workbook NextLetter();
    //    void OnKey(string Key, [Optional] object Procedure);
    //    void OnRepeat(string Text, string Procedure);
    //    void OnTime(object EarliestTime, string Procedure, [Optional] object LatestTime, [Optional] object Schedule);
    //    void OnUndo(string Text, string Procedure);
    //    void Quit();
    //    void RecordMacro([Optional] object BasicCode, [Optional] object XlmCode);
    //    bool RegisterXLL(string Filename);
    //    void Repeat();
    //    void ResetTipWizard();
    //    void Save([Optional] object Filename);
    //    void SaveWorkspace([Optional] object Filename);
    //    void SetDefaultChart([Optional] object FormatName, [Optional] object Gallery);
    //    void Undo();
    //    void Volatile([Optional] object Volatile);
    //    void _Wait(object Time);
    //    object _WSFunction([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    bool Wait(object Time);
    //    string GetPhonetic([Optional] object Text);
    //    void Dummy12(PivotTable p1, PivotTable p2);
    //    void CalculateFull();
    //    bool FindFile();
    //    object Dummy13(object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    //    void Dummy14();
    //    void CalculateFullRebuild();
    //    void CheckAbort([Optional] object KeepAbort);
    //    void DisplayXMLSourcePane([Optional] object XmlMap);
    //    object Support([MarshalAs(UnmanagedType.IDispatch)] object Object, int ID, [Optional] object arg);
    //    object Dummy20(int grfCompareFunctions);
    //    void CalculateUntilAsyncQueriesDone();
    //    int SharePointVersion(string bstrUrl);
    //}

    /// <summary>
    /// NpoiWrapperクラス
    /// Microsoft.Office.Interop.Excel.Applicationをエミュレート
    /// </summary>
    public class _Application
    {
        #region "fields"

        /// <summary>
        /// お決まりの３プロパティの元となる情報
        /// </summary>
        public readonly Application Application;
        public readonly XlCreator Creator = XlCreator.xlCreatorCode;
        public readonly Application Parent;

        /// <summary>
        /// Excel標準の色インデックスを使用するかどうかを示すフラグ(現在未使用)
        /// </summary>
        internal readonly bool Use2003ColorIndex = false;

        /// <summary>
        /// Selectionプロパティの実体
        /// </summary>
        private readonly Dictionary<string, object> _selection = new Dictionary<string, object>();

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public _Application(bool Use2003ColorIndex)
        {
            this.Application = (Application)this;
            this.Parent = (Application)this;
            this.Windows = new Windows(this);
            this.Workbooks = new Workbooks((Application)this);
            this.Use2003ColorIndex = Use2003ColorIndex;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public _Application()
            : this(false)
        {
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Windows Windows { get; }
        public Workbooks Workbooks { get; }
        /// <summary>
        /// ActiveWindow
        /// Workbookにて勝手にセットしてもらう。
        /// </summary>
        public Window ActiveWindow { get; internal set; }
        /// <summary>
        /// ActiveWorkbook
        /// Workbookにて勝手にセットしてもらう。
        /// </summary>
        public Workbook ActiveWorkbook { get; internal set; }
        /// <summary>
        /// ActiveSheet
        /// Worksheetにて勝手にセットしてもらう。
        /// </summary>
        public object ActiveSheet { get; internal set; }
        /// <summary>
        /// ActiveCell
        /// 未サポート
        /// </summary>
        public Range ActiveCell { get; internal set; }
        /// <summary>
        /// Visible(ダミー実装)
        /// </summary>
        public bool Visible { get; internal set; } = false;
        /// <summary>
        /// DisplayAlerts(ダミー実装)
        /// </summary>
        public bool DisplayAlerts { get; internal set; } = false;

        /// <summary>
        /// Selectionプロパティ(getterのみ)
        /// ActiveSheetにセットされたWorksheetオブジェクトからWorkbook.IndexとWorksheet.Nameを取り出し、、
        /// その２つをキーに格納されたオブジェクトを取り出す。
        /// </summary>
        public dynamic Selection
        {
            get
            {
                dynamic RetVal = null;
                //ActiveSheetが存在すること
                if (ActiveSheet != null)
                {
                    //_SelectionにこのSheetのオブジェクトがあればそれを返す。
                    Worksheet Target = (Worksheet)ActiveSheet;
                    string KeyString = Target.Parent.Index.ToString() + ":" + Target.Name;
                    if (_selection.ContainsKey(KeyString))
                    {
                        RetVal = _selection[Target.Parent.Index.ToString() + ":" + Target.Name];
                    }
                }
                return RetVal;
            }
        }

        #endregion

        #endregion

        #region "methods"

        #region "internal methods"

        /// <summary>
        /// Selectionセッター
        /// </summary>
        /// <param name="Setter">セット主体のWorksheet</param>
        /// <param name="SelectedObject">セットするオブジェクト</param>
        /// <exception cref="InvalidOperationException"></exception>
        internal void SetSelection(Worksheet Setter, object SelectedObject)
        {
            Worksheet Current = (Worksheet)ActiveSheet;
            if (Current?.Parent.Index == Setter.Parent.Index && Current?.Name == Setter.Name)
            {

                _selection[Setter.Parent.Index.ToString() + ":" + Setter.Name] = SelectedObject;
            }
            else
            {
                throw new InvalidOperationException("This sheet is not active.");
            }
        }

        #endregion

        #endregion
    }
}
