using MathNet.Numerics;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices;
using System.Web.UI.WebControls;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Workbook interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Workbook : _Workbook, WorkbookEvents_Event
    //{
    //}
    //----------------------------------------------------------------------------------------------
    // _Workbook interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface _Workbook
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    bool AcceptLabelsInFormulas { get; set; }
    //    Chart ActiveChart { get; }
    //    object ActiveSheet { get; }
    //    string Author { get; set; }
    //    int AutoUpdateFrequency { get; set; }
    //    bool AutoUpdateSaveChanges { get; set; }
    //    int ChangeHistoryDuration { get; set; }
    //    object BuiltinDocumentProperties { get; }
    //    Sheets Charts { get; }
    //    string CodeName { get; }
    //    string _CodeName { get; set; }
    //    object Colors { get; set; }
    //    CommandBars CommandBars { get; }
    //    string Comments { get; set; }
    //    XlSaveConflictResolution ConflictResolution { get; set; }
    //    object Container { get; }
    //    bool CreateBackup { get; }
    //    object CustomDocumentProperties { get; }
    //    bool Date { get; set; }
    //    Sheets DialogSheets { get; }
    //    XlDisplayDrawingObjects DisplayDrawingObjects { get; set; }
    //    XlFileFormat FileFormat { get; }
    //    string FullName { get; }
    //    bool HasMailer { get; set; }
    //    bool HasPassword { get; }
    //    bool HasRoutingSlip { get; set; }
    //    bool IsAddin { get; set; }
    //    string Keywords { get; set; }
    //    Mailer Mailer { get; }
    //    Sheets Modules { get; }
    //    bool MultiUserEditing { get; }
    //    string Name { get; }
    //    Names Names { get; }
    //    string OnSave { get; set; }
    //    string OnSheetActivate { get; set; }
    //    string OnSheetDeactivate { get; set; }
    //    string Path { get; }
    //    bool PersonalViewListSettings { get; set; }
    //    bool PersonalViewPrintSettings { get; set; }
    //    bool PrecisionAsDisplayed { get; set; }
    //    bool ProtectStructure { get; }
    //    bool ProtectWindows { get; }
    //    bool ReadOnly { get; }
    //    bool _ReadOnlyRecommended { get; }
    //    int RevisionNumber { get; }
    //    bool Routed { get; }
    //    RoutingSlip RoutingSlip { get; }
    //    bool Saved { get; set; }
    //    bool SaveLinkValues { get; set; }
    //    Sheets Sheets { get; }
    //    bool ShowConflictHistory { get; set; }
    //    Styles Styles { get; }
    //    string Subject { get; set; }
    //    string Title { get; set; }
    //    bool UpdateRemoteReferences { get; set; }
    //    bool UserControl { get; set; }
    //    object UserStatus { get; }
    //    CustomViews CustomViews { get; }
    //    Windows Windows { get; }
    //    Sheets Worksheets { get; }
    //    bool WriteReserved { get; }
    //    string WriteReservedBy { get; }
    //    Sheets ExcelIntlMacroSheets { get; }
    //    Sheets ExcelMacroSheets { get; }
    //    bool TemplateRemoveExtData { get; set; }
    //    bool HighlightChangesOnScreen { get; set; }
    //    bool KeepChangeHistory { get; set; }
    //    bool ListChangesOnNewSheet { get; set; }
    //    VBProject VBProject { get; }
    //    bool IsInplace { get; }
    //    PublishObjects PublishObjects { get; }
    //    WebOptions WebOptions { get; }
    //    HTMLProject HTMLProject { get; }
    //    bool EnvelopeVisible { get; set; }
    //    int CalculationVersion { get; }
    //    bool VBASigned { get; }
    //    bool ShowPivotTableFieldList { get; set; }
    //    XlUpdateLinks UpdateLinks { get; set; }
    //    bool EnableAutoRecover { get; set; }
    //    bool RemovePersonalInformation { get; set; }
    //    string FullNameURLEncoded { get; }
    //    string Password { get; set; }
    //    string WritePassword { get; set; }
    //    string PasswordEncryptionProvider { get; }
    //    string PasswordEncryptionAlgorithm { get; }
    //    int PasswordEncryptionKeyLength { get; }
    //    bool PasswordEncryptionFileProperties { get; }
    //    bool ReadOnlyRecommended { get; set; }
    //    SmartTagOptions SmartTagOptions { get; }
    //    Permission Permission { get; }
    //    SharedWorkspace SharedWorkspace { get; }
    //    Sync Sync { get; }
    //    XmlNamespaces XmlNamespaces { get; }
    //    XmlMaps XmlMaps { get; }
    //    SmartDocument SmartDocument { get; }
    //    DocumentLibraryVersions DocumentLibraryVersions { get; }
    //    bool InactiveListBorderVisible { get; set; }
    //    bool DisplayInkComments { get; set; }
    //    MetaProperties ContentTypeProperties { get; }
    //    Connections Connections { get; }
    //    SignatureSet Signatures { get; }
    //    ServerPolicy ServerPolicy { get; }
    //    DocumentInspectors DocumentInspectors { get; }
    //    ServerViewableItems ServerViewableItems { get; }
    //    TableStyles TableStyles { get; }
    //    object DefaultTableStyle { get; set; }
    //    object DefaultPivotTableStyle { get; set; }
    //    bool CheckCompatibility { get; set; }
    //    bool HasVBProject { get; }
    //    CustomXMLParts CustomXMLParts { get; }
    //    bool Final { get; set; }
    //    Research Research { get; }
    //    OfficeTheme Theme { get; }
    //    bool ExcelCompatibilityMode { get; }
    //    bool ConnectionsDisabled { get; }
    //    bool ShowPivotChartActiveFields { get; set; }
    //    IconSets IconSets { get; }
    //    string EncryptionProvider { get; set; }
    //    bool DoNotPromptForConvert { get; set; }
    //    bool ForceFullCalculation { get; set; }
    //    void Activate();
    //    void ChangeFileAccess(XlFileAccess Mode, [Optional] object WritePassword, [Optional] object Notify);
    //    void ChangeLink(string Name, string NewName, XlLinkType Type = XlLinkType.xlLinkTypeExcelLinks);
    //    void Close([Optional] object SaveChanges, [Optional] object Filename, [Optional] object RouteWorkbook);
    //    void DeleteNumberFormat(string NumberFormat);
    //    bool ExclusiveAccess();
    //    void ForwardMailer();
    //    object LinkInfo(string Name, XlLinkInfo LinkInfo, [Optional] object Type, [Optional] object EditionRef);
    //    object LinkSources([Optional] object Type);
    //    void MergeWorkbook(object Filename);
    //    Window NewWindow();
    //    void OpenLinks(string Name, [Optional] object ReadOnly, [Optional] object Type);
    //    PivotCaches PivotCaches();
    //    void Post([Optional] object DestName);
    //    void _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate);
    //    void PrintPreview([Optional] object EnableChanges);
    //    void _Protect([Optional] object Password, [Optional] object Structure, [Optional] object Windows);
    //    void ProtectSharing([Optional] object Filename, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, [Optional] object SharingPassword);
    //    void RefreshAll();
    //    void Reply();
    //    void ReplyAll();
    //    void RemoveUser(int Index);
    //    void Route();
    //    void RunAutoMacros(XlRunAutoMacro Which);
    //    void Save();
    //    void _SaveAs([Optional] object Filename, [Optional] object FileFormat, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, XlSaveAsAccessMode AccessMode = XlSaveAsAccessMode.xlNoChange, [Optional] object ConflictResolution, [Optional] object AddToMru, [Optional] object TextCodepage, [Optional] object TextVisualLayout);
    //    void SaveCopyAs([Optional] object Filename);
    //    void SendMail(object Recipients, [Optional] object Subject, [Optional] object ReturnReceipt);
    //    void SendMailer([Optional] object FileFormat, XlPriority Priority = XlPriority.xlPriorityNormal);
    //    void SetLinkOnData(string Name, [Optional] object Procedure);
    //    void Unprotect([Optional] object Password);
    //    void UnprotectSharing([Optional] object SharingPassword);
    //    void UpdateFromFile();
    //    void UpdateLink([Optional] object Name, [Optional] object Type);
    //    void HighlightChangesOptions([Optional] object When, [Optional] object Who, [Optional] object Where);
    //    void PurgeChangeHistoryNow(int Days, [Optional] object SharingPassword);
    //    void AcceptAllChanges([Optional] object When, [Optional] object Who, [Optional] object Where);
    //    void RejectAllChanges([Optional] object When, [Optional] object Who, [Optional] object Where);
    //    void PivotTableWizard([Optional] object SourceType, [Optional] object SourceData, [Optional] object TableDestination, [Optional] object TableName, [Optional] object RowGrand, [Optional] object ColumnGrand, [Optional] object SaveData, [Optional] object HasAutoFormat, [Optional] object AutoPage, [Optional] object Reserved, [Optional] object BackgroundQuery, [Optional] object OptimizeCache, [Optional] object PageFieldOrder, [Optional] object PageFieldWrapCount, [Optional] object ReadData, [Optional] object Connection);
    //    void ResetColors();
    //    void FollowHyperlink(string Address, [Optional] object SubAddress, [Optional] object NewWindow, [Optional] object AddHistory, [Optional] object ExtraInfo, [Optional] object Method, [Optional] object HeaderInfo);
    //    void AddToFavorites();
    //    void PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void WebPagePreview();
    //    void ReloadAs(MsoEncoding Encoding);
    //    void Dummy(int calcid);
    //    void sblt(string s);
    //    void BreakLink(string Name, XlLinkType Type);
    //    void Dummy();
    //    void SaveAs([Optional] object Filename, [Optional] object FileFormat, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, XlSaveAsAccessMode AccessMode = XlSaveAsAccessMode.xlNoChange, [Optional] object ConflictResolution, [Optional] object AddToMru, [Optional] object TextCodepage, [Optional] object TextVisualLayout, [Optional] object Local);
    //    void CheckIn([Optional] object SaveChanges, [Optional] object Comments, [Optional] object MakePublic);
    //    bool CanCheckIn();
    //    void SendForReview([Optional] object Recipients, [Optional] object Subject, [Optional] object ShowMessage, [Optional] object IncludeAttachment);
    //    void ReplyWithChanges([Optional] object ShowMessage);
    //    void EndReview();
    //    void SetPasswordEncryptionOptions([Optional] object PasswordEncryptionProvider, [Optional] object PasswordEncryptionAlgorithm, [Optional] object PasswordEncryptionKeyLength, [Optional] object PasswordEncryptionFileProperties);
    //    void Protect([Optional] object Password, [Optional] object Structure, [Optional] object Windows);
    //    void RecheckSmartTags();
    //    void SendFaxOverInternet([Optional] object Recipients, [Optional] object Subject, [Optional] object ShowMessage);
    //    XlXmlImportResult XmlImport(string Url, [MarshalAs(UnmanagedType.Interface)] out XmlMap ImportMap, [Optional] object Overwrite, [Optional] object Destination);
    //    XlXmlImportResult XmlImportXml(string Data, [MarshalAs(UnmanagedType.Interface)] out XmlMap ImportMap, [Optional] object Overwrite, [Optional] object Destination);
    //    void SaveAsXMLData(string Filename, [MarshalAs(UnmanagedType.Interface)] XmlMap Map);
    //    void ToggleFormsDesign();
    //    void RemoveDocumentInformation(XlRemoveDocInfoType RemoveDocInfoType);
    //    void CheckInWithVersion([Optional] object SaveChanges, [Optional] object Comments, [Optional] object MakePublic, [Optional] object VersionType);
    //    void LockServerFile();
    //    WorkflowTasks GetWorkflowTasks();
    //    WorkflowTemplates GetWorkflowTemplates();
    //    void PrintOutEx([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName, [Optional] object IgnorePrintAreas);
    //    void ApplyTheme(string Filename);
    //    void EnableConnections();
    //    void ExportAsFixedFormat(XlFixedFormatType Type, [Optional] object Filename, [Optional] object Quality, [Optional] object IncludeDocProperties, [Optional] object IgnorePrintAreas, [Optional] object From, [Optional] object To, [Optional] object OpenAfterPublish, [Optional] object FixedFormatExtClassPtr);
    //    void ProtectSharingEx([Optional] object Filename, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, [Optional] object SharingPassword, [Optional] object FileFormat);
    //}

    /// <summary>
    /// Workbookクラス
    /// Microsoft.Office.Interop.Excel.Workbookをエミュレート
    /// WorkbooksクラスのAdd、Openでコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Workbook
    {
        public Application Application { get { return this.Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Workbooks Parent { get; }

        internal IWorkbook PoiBook {get; private set; }
        public string FileName { get; private set; } = string.Empty;
        internal Dictionary<string, ICellStyle> CellStyles { get; private set; } = new Dictionary<string, ICellStyle>();
        internal Dictionary<string, Configurations.Models.PageSetup> PageSetups { get; private set; } = new Dictionary<string, Configurations.Models.PageSetup>();
        internal IFont Font { get; private set; } = null;
        public Sheets Worksheets { get; private set; }
        public Sheets Sheets { get; private set; }

        /// <summary>
        /// 新規ファイルを作成する
        /// <param name="ParentApp">親Application</param>
        /// </summary>
        internal Workbook(Workbooks ParentWorkbooks)
            : this(ParentWorkbooks, false)
        {
        }

        /// <summary>
        /// 新規ファイルを作成する
        /// <param name="ParentWorkbooks">親Workbooks</param>
        /// </summary>
        internal Workbook(Workbooks ParentWorkbooks, bool Excel97_2003)
        {
            //親Workbooks保存
            this.Parent = ParentWorkbooks;
            //Excel形式判断
            if (Excel97_2003)
            {
                //Excel97-2003
                PoiBook = new HSSFWorkbook();
            }
            else
            {
                //Excel2007以降
                PoiBook = new XSSFWorkbook();
            }
            //設定ファイルの内容を反映
            ApplyConfigs();
            //新規の場合はシートを一つ追加しておく
            PoiBook.CreateSheet();
            //この時点でファイル名は未定義
            FileName = string.Empty;
            //Sheets,Worksheetsの初期化
            Sheets = new Sheets(this);
            Worksheets = new Worksheets(this);
        }

        /// <summary>
        /// 既存ファイルを開く(読取専用のみサポート)
        /// </summary>
        /// <param name="ParentApp">親Application</param>
        /// <param name="FileName"></param>
        internal Workbook(Workbooks ParentWorkbooks, string FileName)
        {
            ////ファイルを開く
            FileStream Stream = new FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            PoiBook = WorkbookFactory.Create(Stream);
            //ファイル名を保存
            this.FileName = FileName;
            //Sheets,Worksheetsの初期化
            Sheets = new Sheets(this);
            Worksheets = new Worksheets(this);
        }

        /// <summary>
        /// アクティブシートの取得
        /// ★InteropでもWorkbookクラスにある
        /// </summary>
        /// <returns>ISheetインターフェイス</returns>
        public Worksheet ActiveSheet()
        {
            return new Worksheet(this, PoiBook.GetSheetAt(PoiBook.ActiveSheetIndex));
        }

        ///// <summary>
        ///// 開いているファイルの保存
        ///// </summary>
        //public void Save()
        //{
        //    //ファイル保存
        //    using (FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write))
        //    {
        //        PoiBook.Write(fs, false);
        //    }
        //}

        /// <summary>
        /// 名前を付けて保存(保存後は閉じる)
        /// </summary>
        /// <param name="Path">フルパスファイル名</param>
        public void SaveAs(string Path)
        {
            //保存ファイル名生成＆拡張子の取得
            FileName = Path;
            string CurrentExtention = System.IO.Path.GetExtension(FileName);
            //現在のWorkbookの既定拡張子を取得
            string DefaultExtension = PoiBook.SpreadsheetVersion.DefaultExtension;
            //拡張子が既定拡張子と異なる場合は既定拡張子を適用
            if (String.Compare(CurrentExtention, DefaultExtension, true) != 0)
            {
                FileName = System.IO.Path.ChangeExtension(FileName, DefaultExtension);
            }
            //ファイル保存
            using (FileStream Stream = new FileStream(FileName, FileMode.Create, FileAccess.Write))
            {
                PoiBook.Write(Stream, false);
            }
        }

        /// <summary>
        /// 設定ファイル(NpoiWrapper.config)の設定を反映
        /// </summary>
        private void ApplyConfigs()
        {
            //設定ファイル(NpoiWrapper.config)の読込
            Configurations.ConfigurationManager CfgManager = new Configurations.ConfigurationManager();
            //フォントの生成
            if (CfgManager.Configs.Font != null)
            {
                Font = PoiBook.CreateFont();
                Font.FontName = CfgManager.Configs.Font.Name;
                Font.FontHeightInPoints = CfgManager.Configs.Font.Size;
            }
            //ページ設定リストの生成
            foreach (Configurations.Models.PageSetup ps in CfgManager.Configs.PageSetup ?? new List<Configurations.Models.PageSetup>())
            {
                PageSetups.Add(ps.Name, ps);
            }
            //セルスタイルリストの生成
            foreach (Configurations.Models.CellStyle cs in CfgManager.Configs.CellStyle ?? new List<Configurations.Models.CellStyle>())
            {
                ICellStyle pcs = PoiBook.CreateCellStyle();
                pcs.BorderTop = cs.Border.Top;
                pcs.BorderRight = cs.Border.Right;
                pcs.BorderBottom = cs.Border.Bottom;
                pcs.BorderLeft = cs.Border.Left;
                pcs.Alignment = cs.Align.Horizontal;
                pcs.VerticalAlignment = cs.Align.Vertical;
                pcs.WrapText = cs.WrapText.Value;
                pcs.IsLocked = cs.IsLocked.Value;
                if (cs.DataFormat.Value.Length > 0)
                {
                    pcs.DataFormat = PoiBook.CreateDataFormat().GetFormat(cs.DataFormat.Value);
                }
                pcs.FillForegroundColor = cs.Fill.Color;
                pcs.FillPattern = cs.Fill.Pattern;
                if (Font != null)
                {
                    pcs.SetFont(Font);
                }
                CellStyles.Add(cs.Name, pcs);
            }
        }
    }
}
