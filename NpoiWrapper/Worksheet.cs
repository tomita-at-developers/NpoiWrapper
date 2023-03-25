﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using NPOI.HSSF.UserModel;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Worksheet interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Worksheet : _Worksheet, DocEvents_Event
    //{
    //}
    //----------------------------------------------------------------------------------------------
    // _Worksheet interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface _Worksheet
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    string CodeName { get; }
    //    string _CodeName { get; set; }
    //    int Index { get; }
    //    string Name { get; set; }
    //    object Next { get; }
    //    string OnDoubleClick { get; set; }
    //    string OnSheetActivate { get; set; }
    //    string OnSheetDeactivate { get; set; }
    //    PageSetup PageSetup { get; }
    //    object Previous { get; }
    //    bool ProtectContents { get; }
    //    bool ProtectDrawingObjects { get; }
    //    bool ProtectionMode { get; }
    //    bool ProtectScenarios { get; }
    //    XlSheetVisibility Visible { get; set; }
    //    Shapes Shapes { get; }
    //    bool TransitionExpEval { get; set; }
    //    bool AutoFilterMode { get; set; }
    //    bool EnableCalculation { get; set; }
    //    Range Cells { get; }
    //    Range CircularReference { get; }
    //    Range Columns { get; }
    //    XlConsolidationFunction ConsolidationFunction { get; }
    //    object ConsolidationOptions { get; }
    //    object ConsolidationSources { get; }
    //    bool DisplayAutomaticPageBreaks { get; set; }
    //    bool EnableAutoFilter { get; set; }
    //    XlEnableSelection EnableSelection { get; set; }
    //    bool EnableOutlining { get; set; }
    //    bool EnablePivotTable { get; set; }
    //    bool FilterMode { get; }
    //    Names Names { get; }
    //    string OnCalculate { get; set; }
    //    string OnData { get; set; }
    //    string OnEntry { get; set; }
    //    Outline Outline { get; }
    //    Range Range { get; }
    //    Range Rows { get; }
    //    string ScrollArea { get; set; }
    //    double StandardHeight { get; }
    //    double StandardWidth { get; set; }
    //    bool TransitionFormEntry { get; set; }
    //    XlSheetType Type { get; }
    //    Range UsedRange { get; }
    //    HPageBreaks HPageBreaks { get; }
    //    VPageBreaks VPageBreaks { get; }
    //    QueryTables QueryTables { get; }
    //    bool DisplayPageBreaks { get; set; }
    //    Comments Comments { get; }
    //    Hyperlinks Hyperlinks { get; }
    //    int _DisplayRightToLeft { get; set; }
    //    AutoFilter AutoFilter { get; }
    //    bool DisplayRightToLeft { get; set; }
    //    Scripts Scripts { get; }
    //    Tab Tab { get; }
    //    MsoEnvelope MailEnvelope { get; }
    //    CustomProperties CustomProperties { get; }
    //    SmartTags SmartTags { get; }
    //    Protection Protection { get; }
    //    ListObjects ListObjects { get; }
    //    bool EnableFormatConditionsCalculation { get; set; }
    //    Sort Sort { get; }
    //    void Activate();
    //    void Copy([Optional] object Before, [Optional] object After);
    //    void Delete();
    //    void Move([Optional] object Before, [Optional] object After);
    //    void _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate);
    //    void PrintPreview([Optional] object EnableChanges);
    //    void _Protect([Optional] object Password, [Optional] object DrawingObjects, [Optional] object Contents, [Optional] object Scenarios, [Optional] object UserInterfaceOnly);
    //    void _SaveAs(string Filename, [Optional] object FileFormat, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, [Optional] object AddToMru, [Optional] object TextCodepage, [Optional] object TextVisualLayout);
    //    void Select([Optional] object Replace);
    //    void Unprotect([Optional] object Password);
    //    object Arcs([Optional] object Index);
    //    void SetBackgroundPicture(string Filename);
    //    object Buttons([Optional] object Index);
    //    void Calculate();
    //    object ChartObjects([Optional] object Index);
    //    object CheckBoxes([Optional] object Index);
    //    void CheckSpelling([Optional] object CustomDictionary, [Optional] object IgnoreUppercase, [Optional] object AlwaysSuggest, [Optional] object SpellLang);
    //    void ClearArrows();
    //    object Drawings([Optional] object Index);
    //    object DrawingObjects([Optional] object Index);
    //    object DropDowns([Optional] object Index);
    //    object Evaluate(object Name);
    //    object _Evaluate(object Name);
    //    void ResetAllPageBreaks();
    //    object GroupBoxes([Optional] object Index);
    //    object GroupObjects([Optional] object Index);
    //    object Labels([Optional] object Index);
    //    object Lines([Optional] object Index);
    //    object ListBoxes([Optional] object Index);
    //    object OLEObjects([Optional] object Index);
    //    object OptionButtons([Optional] object Index);
    //    object Ovals([Optional] object Index);
    //    void Paste([Optional] object Destination, [Optional] object Link);
    //    void _PasteSpecial([Optional] object Format, [Optional] object Link, [Optional] object DisplayAsIcon, [Optional] object IconFileName, [Optional] object IconIndex, [Optional] object IconLabel);
    //    object Pictures([Optional] object Index);
    //    object PivotTables([Optional] object Index);
    //    PivotTable PivotTableWizard([Optional] object SourceType, [Optional] object SourceData, [Optional] object TableDestination, [Optional] object TableName, [Optional] object RowGrand, [Optional] object ColumnGrand, [Optional] object SaveData, [Optional] object HasAutoFormat, [Optional] object AutoPage, [Optional] object Reserved, [Optional] object BackgroundQuery, [Optional] object OptimizeCache, [Optional] object PageFieldOrder, [Optional] object PageFieldWrapCount, [Optional] object ReadData, [Optional] object Connection);
    //    object Rectangles([Optional] object Index);
    //    object Scenarios([Optional] object Index);
    //    object ScrollBars([Optional] object Index);
    //    void ShowAllData();
    //    void ShowDataForm();
    //    object Spinners([Optional] object Index);
    //    object TextBoxes([Optional] object Index);
    //    void ClearCircles();
    //    void CircleInvalid();
    //    void PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void _CheckSpelling([Optional] object CustomDictionary, [Optional] object IgnoreUppercase, [Optional] object AlwaysSuggest, [Optional] object SpellLang, [Optional] object IgnoreFinalYaa, [Optional] object SpellScript);
    //    void SaveAs(string Filename, [Optional] object FileFormat, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object ReadOnlyRecommended, [Optional] object CreateBackup, [Optional] object AddToMru, [Optional] object TextCodepage, [Optional] object TextVisualLayout, [Optional] object Local);
    //    void PasteSpecial([Optional] object Format, [Optional] object Link, [Optional] object DisplayAsIcon, [Optional] object IconFileName, [Optional] object IconIndex, [Optional] object IconLabel, [Optional] object NoHTMLFormatting);
    //    void Protect([Optional] object Password, [Optional] object DrawingObjects, [Optional] object Contents, [Optional] object Scenarios, [Optional] object UserInterfaceOnly, [Optional] object AllowFormattingCells, [Optional] object AllowFormattingColumns, [Optional] object AllowFormattingRows, [Optional] object AllowInsertingColumns, [Optional] object AllowInsertingRows, [Optional] object AllowInsertingHyperlinks, [Optional] object AllowDeletingColumns, [Optional] object AllowDeletingRows, [Optional] object AllowSorting, [Optional] object AllowFiltering, [Optional] object AllowUsingPivotTables);
    //    Range XmlDataQuery(string XPath, [Optional] object SelectionNamespaces, [Optional] object Map);
    //    Range XmlMapQuery(string XPath, [Optional] object SelectionNamespaces, [Optional] object Map);
    //    void PrintOutEx([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName, [Optional] object IgnorePrintAreas);
    //    void ExportAsFixedFormat([In] XlFixedFormatType Type, [Optional] object Filename, [Optional] object Quality, [Optional] object IncludeDocProperties, [Optional] object IgnorePrintAreas, [Optional] object From, [Optional] object To, [Optional] object OpenAfterPublish, [Optional] object FixedFormatExtClassPtr);
    //}

    /// <summary>
    /// Worksheetクラス
    /// Microsoft.Office.Interop.Excel.Workbookをエミュレート
    /// WorkbookクラスのActiveSheet、SheetsクラスのAddでのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Worksheet
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        /// <param name="Sheet">自ISheet</param>
        internal Worksheet(Workbook ParentWorkbook, ISheet Sheet)
        {
            Logger.Debug("SheetName[" + Sheet.SheetName + "]");

            //親クラスの保存
            this.Parent = ParentWorkbook;
            //POIシートの保存
            PoiSheet = Sheet;
            //Cells, Rangeの初期値をセット
            //インデクスを省略した場合はこの値が取得される(シート全域)
            Cells = new Cells(this, new CellRangeAddressList(-1, -1, -1, -1));
            Range = new Range(this, new CellRangeAddressList(-1, -1, -1, -1));
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Workbook Parent { get; }

        /// <summary>
        /// シート名
        /// </summary>
        public string Name
        {
            get
            {
                return PoiSheet.SheetName;
            }
            set
            {
                Parent.PoiBook.SetSheetName(
                    Parent.PoiBook.GetSheetIndex(PoiSheet.SheetName), value);
            }
        }

        /// <summary>
        /// シートインデックス(1開始)
        /// </summary>
        public long Index { get { return this.PoiIndex + 1; } }

        /// <summary>
        /// Cellsオブジェクト
        /// </summary>
        public Cells Cells { get; }

        /// <summary>
        /// Rangeオブジェクト
        /// </summary>
        public Range Range { get; }

        #endregion

        #region "internal properties"

        /// <summary>
        /// POIシート
        /// </summary>
        internal ISheet PoiSheet { get; private set; }

        /// <summary>
        /// POIのシートIndex(0開始)
        /// </summary>
        internal int PoiIndex { get { return Parent.PoiBook.GetSheetIndex(PoiSheet.SheetName); } }

        /// <summary>
        /// このシートがHSSFSheetかどうか
        /// </summary>
        internal bool IsHSSF { get { return (PoiSheet is HSSFSheet); } }
        /// <summary>
        /// このシートがXSSFSheetかどうか
        /// </summary>
        internal bool IsXSSF { get { return (PoiSheet is XSSFSheet); } }

        #endregion

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// シートの選択
        /// </summary>
        /// <param name="Replace"></param>
        public void Select(object Replace = null)
        {
            bool IsReplace = true;
            if (Replace is bool BoolValue)
            {
                IsReplace = BoolValue;
            }
            //Replace(default:true)ならば他シートのSelect解除。
            if (IsReplace)
            {
                for (int s = 0; s < Parent.PoiBook.NumberOfSheets; s++)
                {
                    Parent.PoiBook.GetSheetAt(s).IsSelected = false;
                }
            }
            //自シートをSelectedにする。
            PoiSheet.IsSelected = true;
        }

        /// <summary>
        /// シートのアクティブ化
        /// </summary>
        public void Activate()
        {
            //親BookをActiveにする(Application.ActiveWorkbook, Application.Windows:Index＝1)
            Parent.Activate();
            //親BookのWindowでActiveSheetに設定する
            Parent.Windows[1].ActiveSheet = this;
            //Application上でActiveSheetにする
            this.Application.ActiveSheet = this;
            //自シートのみをSelectedにする。(これをやらないと過去にActiveにしたシートの幾つかが累積的にSelected状態になってしまう)
            Select(true);
            //自シートをActiveにする。
            Parent.PoiBook.SetActiveSheet(this.PoiIndex);
        }

        /// <summary>
        /// シートのコピー(同一Workbook内のみサポート)
        /// </summary>
        /// <param name="Before">指定したシートの直前にコピー</param>
        /// <param name="After">指定したシートの直後にコピー</param>
        /// <returns></returns>
        public void Copy(object Before = null, object After = null)
        {
            int Type = 0;
            ISheet Position = null;
            //Beforeのチェック
            if (Before is Worksheet SafeBefore)
            {
                Position = SafeBefore.PoiSheet;
                Type += 1;
            }
            //Afterのチェック
            if (After is Worksheet SafeAfter)
            {
                Position = SafeAfter.PoiSheet;
                Type += 2;
            }
            //Before, Afterいずれか一方であること
            if (Type == 1 || Type == 2)
            {
                //同一ブック内であること
                if (Position.Workbook.Equals(Parent.PoiBook))
                {
                    //元シートの名前にサフィックスを付加した名前をもとに一意で安全な名前を取得し、その名前でコピー
                    ISheet Sheet = PoiSheet.CopySheet(GetSafeUniqueSheetName(AddOrUpdateSuffix(PoiSheet.SheetName)));
                    //指定された場所にインデックス変更
                    int PositionIndex = Parent.PoiBook.GetSheetIndex(Position);
                    Parent.PoiBook.SetSheetOrder(Sheet.SheetName, Type == 1 ? PositionIndex : PositionIndex + 1);
                }
                //別ブックならば例外スロー
                else
                {
                    throw new ArgumentException("Copy to another Workbook is not supported in this vsersion.");
                }
            }
            //Before, Afterの両方が指定されていれば例外スロー
            else
            {
                throw new ArgumentException("Both of Before and After are specified."); 
            }
        }

        /// <summary>
        /// シートの削除
        /// </summary>
        public void Delete()
        {
            int SheetIndex = Parent.PoiBook.GetSheetIndex(PoiSheet.SheetName);
            Parent.PoiBook.RemoveSheetAt(SheetIndex);
        }

        /// <summary>
        /// シートの保護
        /// HSSFではこの操作によりオートフィルターが無効化される。
        /// </summary>
        /// <param name="Password">パスワード</param>
        /// <param name="...">Password以外はすべて無視</param>
        public void Protect(
            object Password = null, object DrawingObjects = null, object Contents = null, object Scenarios = null,
            object UserInterfaceOnly = null, object AllowFormattingCells = null, object AllowFormattingColumns = null,
            object AllowFormattingRows = null, object AllowInsertingColumns = null, object AllowInsertingRows = null,
            object AllowInsertingHyperlinks = null, object AllowDeletingColumns = null, object AllowDeletingRows = null,
            object AllowSorting = null, object AllowFiltering = null, object AllowUsingPivotTables = null)
        {
            string Passwd = "";
            if (Password is string SafePqassword)
            {
                Passwd = SafePqassword;
            }
            PoiSheet.ProtectSheet(Passwd);
            //XSSFならロック解除できるのでやっておく
            if (PoiSheet is XSSFSheet xssfSheet)
            {
                xssfSheet.LockAutoFilter(false);
                xssfSheet.LockSort(false);
            }
            else
            {
                //為す術なし
            }
        }

        #endregion

        #region "alternative public methods"

        /// <summary>
        /// ページ設定
        /// </summary>
        /// <param name="StyleName">NpoiWrapper.configのページ設定パターン名(省略時"default")</param>
        /// <param name="HeaderLeft">ヘッダー左文字列</param>
        /// <param name="HeaderCenter">ヘッダー中央文字列</param>
        /// <param name="HeaderRight">ヘッダー右文字列</param>
        /// <param name="FooterLeft">フッター左文字列</param>
        /// <param name="FooterCenter">フッター中央文字列</param>
        /// <param name="FooterRight">フッター右文字列</param>
        public void PageSetup(
            string StyleName = "default",
            string HeaderLeft = "",
            string HeaderCenter = "",
            string HeaderRight = "",
            string FooterLeft = "",
            string FooterCenter = "",
            string FooterRight = "")
        {
            //指定された名前の設定が存在すればそれを適用
            if (Parent.PageSetups.ContainsKey(StyleName))
            {
                Configurations.Models.PageSetup Setup = Parent.PageSetups[StyleName];
                PoiSheet.PrintSetup.Landscape = Setup.Paper.Landscape;
                PoiSheet.PrintSetup.PaperSize = (short)Setup.Paper.Size;
                PoiSheet.PrintSetup.HeaderMargin = Setup.Margins.Header.ValueInInch;
                PoiSheet.PrintSetup.FooterMargin = Setup.Margins.Footer.ValueInInch;
                if (Setup.Scaling.Fit != null)
                {
                    PoiSheet.FitToPage = true;
                    PoiSheet.PrintSetup.FitWidth = Setup.Scaling.Fit.Wide;
                    PoiSheet.PrintSetup.FitHeight = Setup.Scaling.Fit.Tall;
                }
                else if (Setup.Scaling.Adjust != null)
                {
                    PoiSheet.FitToPage = false;
                    PoiSheet.PrintSetup.Scale = Setup.Scaling.Adjust.Scale;
                }
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.TopMargin, Setup.Margins.Body.TopInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.RightMargin, Setup.Margins.Body.RightInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.BottomMargin, Setup.Margins.Body.BottomInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.LeftMargin, Setup.Margins.Body.LeftInInch);
                PoiSheet.HorizontallyCenter = Setup.Center.Horizontally;
                PoiSheet.VerticallyCenter = Setup.Center.Vertically;
                if (Setup.Titles.Row.Length > 0)
                {
                    PoiSheet.RepeatingRows = CellRangeAddress.ValueOf(Setup.Titles.Row);
                }
                if (Setup.Titles.Column.Length > 0)
                {
                    PoiSheet.RepeatingColumns = CellRangeAddress.ValueOf(Setup.Titles.Column);
                }
            }
            //ヘッダー/フッターの文字は引数から適用
            PoiSheet.Header.Left = HeaderLeft;
            PoiSheet.Header.Center = HeaderCenter;
            PoiSheet.Header.Right = HeaderRight;
            PoiSheet.Footer.Left = FooterLeft;
            PoiSheet.Footer.Center = FooterCenter;
            PoiSheet.Footer.Right = FooterRight;
        }

        /// <summary>
        /// 先頭行/先頭列の固定
        /// Interop.Excelでは以下のような指定方法であり、WindowというPOIでは少々捉えにくい概念を含んでいる。
        /// POIのIWorkbook.SetActiveSheetで実現できそうだが、POIではFreezePaneがSheetの機能なので、素直にWorksheetにした。
        /// WorkSheet.Activate()
        /// WorkSheet.Range("A2").Select()
        /// Application.ActiveWindow.FreezePanes = True
        /// </summary>
        /// <param name="TopLeftCell">固定位置(A1形式)</param>
        public void CreateFreezePane(string TopLeftCell)
        {
            PoiSheet.CreateFreezePane(
                CellRangeAddress.ValueOf(TopLeftCell).FirstColumn,
                CellRangeAddress.ValueOf(TopLeftCell).FirstRow);
        }

        /// <summary>
        /// オートフィルターの設定
        /// HSSFではProtectをコールするとオートフィルタが無効化される。
        /// Interop.Excelでは以下のような指定方法であり、Rangeの機能だが、POIではSheetのメソッドとして実装されている。
        /// なので直球勝負でWorksheetに実装してみる。。
        /// myRange = WorkSheet.Range("A1")
        /// myRange.AutoFilter()
        /// </summary>
        /// <param name="FilterRange"></param>
        public void AutoFilter(string FilterRange)
        {
            PoiSheet.SetAutoFilter(CellRangeAddress.ValueOf(FilterRange));
        }

        #endregion

        #region "private methods"

        /// <summary>
        /// 指定された文字列(シート名)にサフィックスを追加。(既にあれば更新。)
        /// </summary>
        /// <param name="OriginalName"></param>
        /// <returns></returns>
        private string AddOrUpdateSuffix(string OriginalName)
        {
            string RetVal;
            Regex Reg1 = new Regex(@"\s\([0-9]{1,}\)$");
            Match Match = Reg1.Match(OriginalName);
            //末尾が" (9)"ならその番号をアップ
            if (Match.Success)
            {
                RetVal = Reg1.Replace(OriginalName, "");
                Regex Reg2 = new Regex(@"[0-9]{1,}");
                Match Mach2 = Reg2.Match(Match.Value);
                string SeqNoString = Mach2.Value;
                int SeqNo = int.Parse(SeqNoString);
                RetVal = RetVal + " (" + (SeqNo + 1).ToString() + ")";
            }
            //単に" (2)"を付加
            else
            {
                RetVal = OriginalName + " (2)";
            }
            return RetVal;
        }

        /// <summary>
        /// 指定された名前を安全かつブック内で重複のない唯一の名前にする。
        /// </summary>
        /// <param name="NameProposal">元とする名前</param>
        /// <returns></returns>
        private string GetSafeUniqueSheetName(string NameProposal)
        {
            //使えない文字をスペースに変換する
            string RetVal = WorkbookUtil.CreateSafeSheetName(NameProposal);
            //既存シート名のリストアップ
            List<string> SheetNameList = new List<string>();
            for (int s = 0; s < Parent.PoiBook.NumberOfSheets; s++)
            {
                SheetNameList.Add(Parent.PoiBook.GetSheetAt(s).SheetName);
            }
            //既存シートとの名前衝突がなくなるまでループ
            bool Unique = false;
            while (Unique == false)
            {
                int Duplication = 0;
                foreach (string SheetName in SheetNameList)
                {
                    if (RetVal == SheetName)
                    {
                        //名前の更新
                        RetVal = AddOrUpdateSuffix(RetVal);
                        //重複カウント
                        Duplication += 1;
                    }
                }
                //重複がなければ脱出
                if (Duplication == 0)
                {
                    Unique = true;
                }
            }
            return RetVal;
        }

        #endregion

        #endregion
    }
}
