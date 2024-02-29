using Developers.NpoiWrapper.Model;
using Developers.NpoiWrapper.Utils;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Range interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Range
    //{
    // +  Application Application { get; }
    // +  XlCreator Creator { get; }
    // +  object Parent { get; }
    //    object AddIndent { get; set; }
    // +  string Address { get; }
    //    string AddressLocal { get; }
    // +  Areas Areas { get; }
    // +  Borders Borders { get; }
    // +  Range Cells { get; }
    //    Characters Characters { get; }
    // +  int Column { get; }
    // +  Range Columns { get; }
    // +  object ColumnWidth { get; set; }
    // +  int Count { get; }
    //    Range CurrentArray { get; }
    //    Range CurrentRegion { get; }
    //    [IndexerName("_Default")]
    // +  object this[[Optional] object RowIndex, [Optional] object ColumnIndex] { get; set; }
    //    Range Dependents { get; }
    //    Range DirectDependents { get; }
    //    Range DirectPrecedents { get; }
    //    Range End { get; }
    // +  Range EntireColumn { get; }
    // +  Range EntireRow { get; }
    // +  Font Font { get; }
    // +  object Formula { get; set; }
    //    object FormulaArray { get; set; }
    //    XlFormulaLabel FormulaLabel { get; set; }
    //    object FormulaHidden { get; set; }
    //    object FormulaLocal { get; set; }
    //    object FormulaR1C1 { get; set; }
    //    object FormulaR1C1Local { get; set; }
    //    object HasArray { get; }
    //    object HasFormula { get; }
    // +  object Height { get; }
    //    object Hidden { get; set; }
    // +  object HorizontalAlignment { get; set; }
    //    object IndentLevel { get; set; }
    // +  Interior Interior { get; }
    //    object Item { get; set; }
    //    object Left { get; }
    //    int ListHeaderRows { get; }
    //    XlLocationInTable LocationInTable { get; }
    // +  object Locked { get; set; }
    //    Range MergeArea { get; }
    //    object MergeCells { get; set; }
    //    object Name { get; set; }
    //    Range Next { get; }
    // +  object NumberFormat { get; set; }
    // -  object NumberFormatLocal { get; set; }
    //    Range Offset { get; }
    //    object Orientation { get; set; }
    //    object OutlineLevel { get; set; }
    //    int PageBreak { get; set; }
    //    PivotField PivotField { get; }
    //    PivotItem PivotItem { get; }
    //    PivotTable PivotTable { get; }
    //    Range Precedents { get; }
    //    object PrefixCharacter { get; }
    //    Range Previous { get; }
    //    QueryTable QueryTable { get; }
    // +  Range Range { get; }
    //    Range Resize { get; }
    // +  int Row { get; }
    // +  object RowHeight { get; set; }
    // +  Range Rows { get; }
    //    object ShowDetail { get; set; }
    //    object ShrinkToFit { get; set; }
    //    SoundNote SoundNote { get; }
    //    object Style { get; set; }
    //    object Summary { get; }
    // +  object Text { get; }
    //    object Top { get; }
    //    object UseStandardHeight { get; set; }
    //    object UseStandardWidth { get; set; }
    // +  Validation Validation { get; }
    // +  object Value { get; set; }
    //    object Value2 { get; set; }
    // +  object VerticalAlignment { get; set; }
    // +  object Width { get; }
    //    Worksheet Worksheet { get; }
    // +  object WrapText { get; set; }
    // +  Comment Comment { get; }
    //    Phonetic Phonetic { get; }
    //    FormatConditions FormatConditions { get; }
    //    int ReadingOrder { get; set; }
    //    Hyperlinks Hyperlinks { get; }
    //    Phonetics Phonetics { get; }
    //    string ID { get; set; }
    //    PivotCell PivotCell { get; }
    //    Errors Errors { get; }
    //    SmartTags SmartTags { get; }
    //    bool AllowEdit { get; }
    //    ListObject ListObject { get; }
    //    XPath XPath { get; }
    //    Actions ServerActions { get; }
    //    string MDX { get; }
    //    object CountLarge { get; }
    //    object Activate();
    //    object AdvancedFilter(XlFilterAction Action, [Optional] object CriteriaRange, [Optional] object CopyToRange, [Optional] object Unique);
    //    object ApplyNames([Optional] object Names, [Optional] object IgnoreRelativeAbsolute, [Optional] object UseRowColumnNames, [Optional] object OmitColumn, [Optional] object OmitRow, XlApplyNamesOrder Order = XlApplyNamesOrder.xlRowThenColumn, [Optional] object AppendLast);
    //    object ApplyOutlineStyles();
    //    string AutoComplete(string String);
    //    object AutoFill(Range Destination, XlAutoFillType Type = XlAutoFillType.xlFillDefault);
    //    object AutoFilter([Optional] object Field, [Optional] object Criteria1, XlAutoFilterOperator Operator = XlAutoFilterOperator.xlAnd, [Optional] object Criteria2, [Optional] object VisibleDropDown);
    // +  object AutoFit();
    //    object AutoFormat(XlRangeAutoFormat Format = XlRangeAutoFormat.xlRangeAutoFormatClassic1, [Optional] object Number, [Optional] object Font, [Optional] object Alignment, [Optional] object Border, [Optional] object Pattern, [Optional] object Width);
    //    object AutoOutline();
    // +  object BorderAround([Optional] object LineStyle, XlBorderWeight Weight = XlBorderWeight.xlThin, XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic, [Optional] object Color);
    //    object Calculate();
    //    object CheckSpelling([Optional] object CustomDictionary, [Optional] object IgnoreUppercase, [Optional] object AlwaysSuggest, [Optional] object SpellLang);
    //    object Clear();
    //    object ClearContents();
    //    object ClearFormats();
    //    object ClearNotes();
    //    object ClearOutline();
    //    Range ColumnDifferences(object Comparison);
    //    object Consolidate([Optional] object Sources, [Optional] object Function, [Optional] object TopRow, [Optional] object LeftColumn, [Optional] object CreateLinks);
    //    object Copy([Optional] object Destination);
    //    int CopyFromRecordset(object Data, [Optional] object MaxRows, [Optional] object MaxColumns);
    //    object CopyPicture(XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, XlCopyPictureFormat Format = XlCopyPictureFormat.xlPicture);
    //    object CreateNames([Optional] object Top, [Optional] object Left, [Optional] object Bottom, [Optional] object Right);
    //    object CreatePublisher([Optional] object Edition, XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, [Optional] object ContainsPICT, [Optional] object ContainsBIFF, [Optional] object ContainsRTF, [Optional] object ContainsVALU);
    //    object Cut([Optional] object Destination);
    //    object DataSeries([Optional] object Rowcol, XlDataSeriesType Type = XlDataSeriesType.xlDataSeriesLinear, XlDataSeriesDate Date = XlDataSeriesDate.xlDay, [Optional] object Step, [Optional] object Stop, [Optional] object Trend);
    //    object Delete([Optional] object Shift);
    //    object DialogBox();
    //    object EditionOptions(XlEditionType Type, XlEditionOptionsOption Option, [Optional] object Name, [Optional] object Reference, XlPictureAppearance Appearance = XlPictureAppearance.xlScreen, XlPictureAppearance ChartSize = XlPictureAppearance.xlScreen, [Optional] object Format);
    //    object FillDown();
    //    object FillLeft();
    //    object FillRight();
    //    object FillUp();
    //    Range Find(object What, [Optional] object After, [Optional] object LookIn, [Optional] object LookAt, [Optional] object SearchOrder, XlSearchDirection SearchDirection = XlSearchDirection.xlNext, [Optional] object MatchCase, [Optional] object MatchByte, [Optional] object SearchFormat);
    //    Range FindNext([Optional] object After);
    //    Range FindPrevious([Optional] object After);
    //    object FunctionWizard();
    //    bool GoalSeek(object Goal, Range ChangingCell);
    //    object Group([Optional] object Start, [Optional] object End, [Optional] object By, [Optional] object Periods);
    //    void InsertIndent(int InsertAmount);
    //    object Insert([Optional] object Shift, [Optional] object CopyOrigin);
    //    object Justify();
    //    object ListNames();
    //    void Merge([Optional] object Across);
    //    void UnMerge();
    //    object NavigateArrow([Optional] object TowardPrecedent, [Optional] object ArrowNumber, [Optional] object LinkNumber);
    //    IEnumerator GetEnumerator();
    //    string NoteText([Optional] object Text, [Optional] object Start, [Optional] object Length);
    //    object Parse([Optional] object ParseLine, [Optional] object Destination);
    //    object _PasteSpecial(XlPasteType Paste = XlPasteType.xlPasteAll, XlPasteSpecialOperation Operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, [Optional] object SkipBlanks, [Optional] object Transpose);
    //    object _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate);
    //    object PrintPreview([Optional] object EnableChanges);
    //    object RemoveSubtotal();
    //    bool Replace(object What, object Replacement, [Optional] object LookAt, [Optional] object SearchOrder, [Optional] object MatchCase, [Optional] object MatchByte, [Optional] object SearchFormat, [Optional] object ReplaceFormat);
    //    Range RowDifferences(object Comparison);
    //    object Run([Optional] object Arg1, [Optional] object Arg2, [Optional] object Arg3, [Optional] object Arg4, [Optional] object Arg5, [Optional] object Arg6, [Optional] object Arg7, [Optional] object Arg8, [Optional] object Arg9, [Optional] object Arg10, [Optional] object Arg11, [Optional] object Arg12, [Optional] object Arg13, [Optional] object Arg14, [Optional] object Arg15, [Optional] object Arg16, [Optional] object Arg17, [Optional] object Arg18, [Optional] object Arg19, [Optional] object Arg20, [Optional] object Arg21, [Optional] object Arg22, [Optional] object Arg23, [Optional] object Arg24, [Optional] object Arg25, [Optional] object Arg26, [Optional] object Arg27, [Optional] object Arg28, [Optional] object Arg29, [Optional] object Arg30);
    // +  object Select();
    //    object Show();
    //    object ShowDependents([Optional] object Remove);
    //    object ShowErrors();
    //    object ShowPrecedents([Optional] object Remove);
    //    object Sort([Optional] object Key1, XlSortOrder Order1 = XlSortOrder.xlAscending, [Optional] object Key2, [Optional] object Type, XlSortOrder Order2 = XlSortOrder.xlAscending, [Optional] object Key3, XlSortOrder Order3 = XlSortOrder.xlAscending, XlYesNoGuess Header = XlYesNoGuess.xlNo, [Optional] object OrderCustom, [Optional] object MatchCase, XlSortOrientation Orientation = XlSortOrientation.xlSortRows, XlSortMethod SortMethod = XlSortMethod.xlPinYin, XlSortDataOption DataOption1 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption3 = XlSortDataOption.xlSortNormal);
    //    object SortSpecial(XlSortMethod SortMethod = XlSortMethod.xlPinYin, [Optional] object Key1, XlSortOrder Order1 = XlSortOrder.xlAscending, [Optional] object Type, [Optional] object Key2, XlSortOrder Order2 = XlSortOrder.xlAscending, [Optional] object Key3, XlSortOrder Order3 = XlSortOrder.xlAscending, XlYesNoGuess Header = XlYesNoGuess.xlNo, [Optional] object OrderCustom, [Optional] object MatchCase, XlSortOrientation Orientation = XlSortOrientation.xlSortRows, XlSortDataOption DataOption1 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption2 = XlSortDataOption.xlSortNormal, XlSortDataOption DataOption3 = XlSortDataOption.xlSortNormal);
    // -  Range SpecialCells(XlCellType Type, [Optional] object Value);
    //    object SubscribeTo(string Edition, XlSubscribeToFormat Format = XlSubscribeToFormat.xlSubscribeToText);
    //    object Subtotal(int GroupBy, XlConsolidationFunction Function, object TotalList, [Optional] object Replace, [Optional] object PageBreaks, XlSummaryRow SummaryBelowData = XlSummaryRow.xlSummaryBelow);
    //    object Table([Optional] object RowInput, [Optional] object ColumnInput);
    //    object TextToColumns([Optional] object Destination, XlTextParsingType DataType = XlTextParsingType.xlDelimited, XlTextQualifier TextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, [Optional] object ConsecutiveDelimiter, [Optional] object Tab, [Optional] object Semicolon, [Optional] object Comma, [Optional] object Space, [Optional] object Other, [Optional] object OtherChar, [Optional] object FieldInfo, [Optional] object DecimalSeparator, [Optional] object ThousandsSeparator, [Optional] object TrailingMinusNumbers);
    //    object Ungroup();
    // +  Comment AddComment([Optional] object Text);
    //    void ClearComments();
    //    void SetPhonetic();
    //    object PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void Dirty();
    //    void Speak([Optional] object SpeakDirection, [Optional] object SpeakFormulas);
    //    object PasteSpecial(XlPasteType Paste = XlPasteType.xlPasteAll, XlPasteSpecialOperation Operation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, [Optional] object SkipBlanks, [Optional] object Transpose);
    //    void RemoveDuplicates([Optional] object Columns, XlYesNoGuess Header = XlYesNoGuess.xlNo);
    //    object PrintOutEx([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    void ExportAsFixedFormat(XlFixedFormatType Type, [Optional] object Filename, [Optional] object Quality, [Optional] object IncludeDocProperties, [Optional] object IgnorePrintAreas, [Optional] object From, [Optional] object To, [Optional] object OpenAfterPublish, [Optional] object FixedFormatExtClassPtr);
    //    object CalculateRowMajorOrder();
    //}

    /// <summary>
    /// Rangeクラス
    /// WorksheetクラスのCells, Rangeプロパティにアクセスすると本クラスのインデクサでコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class _Range : IEnumerable, IEnumerator
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);
        private static readonly log4net.ILog ProcTimeLogger = log4net.LogManager.GetLogger("ProcessingTime");
        /// <summary>
        /// IEnumerator用行インデクス
        /// </summary>
        private int EnumeratorRowIndex = 0;
        /// <summary>
        /// IEnumerator用列インデクス
        /// </summary>
        private int EnumeratorColumnIndex = -1;
        /// <summary>
        /// Rangeのデータ/スタイル情報の管理
        /// </summary>
        private RangeManager _RangeManager = null;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        internal _Range(Worksheet ParentSheet, CellRangeAddressList RangeAddressList)
        {
            Logger.Debug(
                "ShhetName[" + ParentSheet.Name + "] " +
                RangeUtil.CellRangeAddressListToString(RangeAddressList));

            this.Parent = ParentSheet;
            this.RawAddressList = RangeAddressList;
            //アドレスに含まれる[-1]を実インデックスに変換し、安全にアクセス可能とする
            this.SafeAddressList = RangeUtil.CreateSafeCellRangeAddressList(
                                        this.RawAddressList, this.Parent.Parent.PoiBook.SpreadsheetVersion);
        }

        /// <summary>
        /// コンストラクタ(Range.Range, Range.Cellsを生成する場合に使用｡)
        /// Rangeクラス内でしか利用しないのでprivateとしている。
        /// </summary>
        /// <param name="ParentSheet"></param>
        /// <param name="RangeAddressList">相対表現のアドレスリスト</param>
        /// <param name="RelativeTo">基点アドレス</param>
        internal _Range(
                    Worksheet ParentSheet, CellRangeAddressList RangeAddressList,
                    CellRangeAddress RelativeTo)
                    : this(ParentSheet, RangeAddressList)
        {
            //追加パラメータ保存
            this.RelativeTo = RelativeTo;
            //RawAddressListのアドレス再作成(基点アドレスを加算し絶対アドレスに変換)
            this.RawAddressList = RangeUtil.CreateAbsoluteCellRangeAddressList(this.RawAddressList, RelativeTo);
            //SafeAddressList再作成
            this.SafeAddressList = RangeUtil.CreateSafeCellRangeAddressList(
                                        this.RawAddressList, this.Parent.Parent.PoiBook.SpreadsheetVersion);
        }

        /// <summary>
        /// コンストラクタ(Range.Rows, Range.Columnsを生成する場合に使用｡)
        /// </summary>
        /// <param name="ParentSheet"></param>
        /// <param name="RangeAddressList">相対表現のアドレスリスト</param>
        /// <param name="RelativeTo">基点アドレス</param>
        /// <param name="CountAs"></param>
        internal _Range(
                    Worksheet ParentSheet, CellRangeAddressList RangeAddressList, CellRangeAddress RelativeTo,
                    CountType CountAs)
                    : this(ParentSheet, RangeAddressList, RelativeTo)
        {
            //追加パラメータ保存
            this.CountAs = CountAs;
        }

        #endregion

        #region "enums"

        /// <summary>
        /// Countプロパティが示す値の種別(セル数、行数、列数)
        /// </summary>
        internal enum CountType
        {
            /// <summary>
            /// Countはセル数を示す(標準)
            /// </summary>
            Default,
            /// <summary>
            /// Countは行数を示す
            /// </summary>
            Rows,
            /// <summary>
            /// Countは列数を示す
            /// </summary>
            Columns
        }

        #endregion

        #region "interface implementations"

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
            CellRangeAddress Adr = SafeAddressList.GetCellRangeAddress(0);
            //次にカラムがない場合は次の行の先頭カラムへ
            EnumeratorColumnIndex += 1;
            if (Adr.FirstColumn + EnumeratorColumnIndex > Adr.LastColumn)
            {
                EnumeratorRowIndex += 1;
                EnumeratorColumnIndex = 0;
                //まだ行があればture
                if (Adr.FirstRow + EnumeratorRowIndex <= Adr.LastRow)
                {
                    RetVal = true;
                }
            }
            //次にカラムがあれば行を維持し次のカラムへ
            else
            {
                EnumeratorColumnIndex += 0;
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
                CellRangeAddress Adr = SafeAddressList.GetCellRangeAddress(0);
                return new Range(
                    Parent,
                    new CellRangeAddressList(
                        Adr.FirstRow + EnumeratorRowIndex, Adr.FirstRow + EnumeratorRowIndex,
                        Adr.FirstColumn + EnumeratorColumnIndex, Adr.FirstColumn + EnumeratorColumnIndex),
                    RelativeTo);
            }
        }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset()
        {
            EnumeratorRowIndex = 0;
            EnumeratorColumnIndex = -1;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Worksheet Parent { get; }

        /// <summary>
        /// Count
        /// このRageに含まれるセル、行、列の数
        /// </summary>
        public int Count
        {
            get
            {
                int RetVal = 0;
                //行数をカウント
                if (CountAs == CountType.Rows)
                {
                    CellRangeAddress RawAddress = SafeAddressList.GetCellRangeAddress(0);
                    RetVal = RawAddress.LastRow - RawAddress.FirstRow + 1;
                }
                //列数をカウント
                else if (CountAs == CountType.Columns)
                {
                    CellRangeAddress RawAddress = SafeAddressList.GetCellRangeAddress(0);
                    RetVal = RawAddress.LastColumn - RawAddress.FirstColumn + 1;
                }
                //セル数をカウント
                else
                {
                    //RawAddressListではEntireの場合にアドレスが-1となり、Cells数が正しく評価されない
                    //それゆえここではSafeAddressListを使用している
                    for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                    {
                        CellRangeAddress RawAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                        RetVal += RawAddress.NumberOfCells;
                    }
                }
                return RetVal;
            }
        }

        /// <summary>
        /// Address
        /// レンジのアドレス
        /// A1形式固定(Interop.Excel.Range.Addressのデフォルトのみサポート)
        /// </summary>
        public string Address
        {
            get
            {
                string RetVal = string.Empty;
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx);
                    RetVal += RawAddress.FormatAsString(null, true) + ",";

                }
                RetVal = RetVal.TrimEnd(',');
                return RetVal;
            }
        }

        /// <summary>
        /// Areas
        /// 複数Rangeで構成される場合、Areasから個々のRangeを取り出せる
        /// </summary>
        public Areas Areas
        {
            get
            {
                return new Areas((Range)this);
            }
        }

        /// <summary>
        /// Cells
        /// このRangeを起点としたサプRangeとしてのCells  !!!!Cells CellsからRange Cellsに変更！
        /// </summary>
        public Range Cells
        {
            get
            {
                //このRangeの先頭アドレスを起点としたCellsを生成
                return new Cells(
                    Parent,
                    new CellRangeAddressList(-1, -1, -1, -1),
                    RawAddressList.GetCellRangeAddress(0));
            }
        }

        /// <summary>
        /// RelativeRange
        /// Interop.ExcelのRange.Rangeプロパティ。
        /// 現在のRangeの始点を起点("A1")としたRangeが取得できる。
        /// クラス名と同じ名前にできないので、このクラスを_Rangeとして衝突をさけ、本クラスを継承するRangeクラスでラップしている。
        /// </summary>
        public Range Range
        {
            get
            {
                //このRangeの先頭アドレスを起点としたRangeを生成
                return new Range(
                    Parent,
                    new CellRangeAddressList(-1, -1, -1, -1),
                    RawAddressList.GetCellRangeAddress(0));
            }
        }

        /// <summary>
        /// 現在のRangeに含まれる行の全体(全カラム)
        /// </summary>
        public Range EntireRow
        {
            get
            {
                //全レンジを処理
                CellRangeAddressList AddressList = new CellRangeAddressList();
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    //生アドレス取得
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx).Copy();
                    //列を全域に拡張しリストに追加
                    AddressList.AddCellRangeAddress(
                        new CellRangeAddress(RawAddress.FirstRow, RawAddress.LastRow, -1, -1));
                }
                //Rangeクラスインスタンス生成
                return new Range(Parent, AddressList, null, CountType.Rows);
            }
        }

        /// <summary>
        /// 現在のRangeに含まれる列の全体(全行)
        /// </summary>
        public Range EntireColumn
        {
            get
            {
                //全レンジを処理
                CellRangeAddressList AddressList = new CellRangeAddressList();
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    //生アドレス取得
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx).Copy();
                    //行を全域に拡張しリストに追加
                    AddressList.AddCellRangeAddress(
                        new CellRangeAddress(-1, -1, RawAddress.FirstColumn, RawAddress.LastColumn));
                }
                //Rangeクラスインスタンス生成
                return new Range(Parent, AddressList);
            }
        }

        /// <summary>
        /// 先頭アドレスの先頭行インデックス取得(１開始)
        /// </summary>
        public int Row
        {
            get
            {
                return SafeAddressList.GetCellRangeAddress(0).FirstRow + 1;
            }
        }

        /// <summary>
        /// 先頭アドレスのRangeを生成(RangeType.Rows)
        /// </summary>
        public Range Rows
        {
            get
            {
                //先頭アドレスのみ切り出し
                CellRangeAddressList AddressList = new CellRangeAddressList();
                AddressList.AddCellRangeAddress(RawAddressList.GetCellRangeAddress(0).Copy());
                //Rangeクラスインスタンス生成
                return new Range(Parent, AddressList, RelativeTo, CountType.Rows);
            }
        }

        /// <summary>
        /// 先頭アドレスの先頭列インデックス取得(１開始)
        /// </summary>
        public int Column
        {
            get
            {
                return SafeAddressList.GetCellRangeAddress(0).FirstColumn + 1;
            }
        }

        /// <summary>
        /// 先頭アドレスのRangeを生成(RangeType.Columns)
        /// </summary>
        public Range Columns
        {
            get
            {
                //先頭アドレスのみ切り出し
                CellRangeAddressList AddressList = new CellRangeAddressList();
                AddressList.AddCellRangeAddress(RawAddressList.GetCellRangeAddress(0).Copy());
                //Rangeクラスインスタンス生成
                return new Range(Parent, AddressList, RelativeTo, CountType.Columns);
            }
        }

        /// <summary>
        /// レンジの値       !!!!!interface上はobject
        /// </summary>
        public object Value
        {
            get { return RangeManager.Value; }
            set { RangeManager.Value = value; }
        }
        //public dynamic Value
        //{
        //    get
        //    {
        //        //Office.Interop.Excelにならい先頭アドレスのみ参照
        //        CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
        //        //値リストの確保
        //        object[,] Values = RangeUtil.CreateCellArray(
        //            SafeAddress.LastRow - SafeAddress.FirstRow + 1, SafeAddress.LastColumn - SafeAddress.FirstColumn + 1);
        //        //行ループ
        //        for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
        //        {
        //            //行の取得(なければ生成)
        //            IRow row = Parent.PoiSheet.GetRow(RIdx) ?? Parent.PoiSheet.CreateRow(RIdx);
        //            //列ループ
        //            for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
        //            {
        //                //列の取得(なければ生成)
        //                ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
        //                object CelVal;
        //                //セルの型に応じたプロパティを参照する
        //                switch (cell.CellType)
        //                {
        //                    //文字列
        //                    case CellType.String:
        //                        CelVal = cell.StringCellValue;
        //                        break;
        //                    //数値
        //                    case CellType.Numeric:
        //                        if (DateUtil.IsCellDateFormatted(cell))
        //                            CelVal = cell.DateCellValue.ToString();
        //                        else
        //                            CelVal = cell.NumericCellValue.ToString();
        //                        break;
        //                    //Boolean
        //                    case CellType.Boolean:
        //                        CelVal = cell.BooleanCellValue.ToString();
        //                        break;
        //                    //式(評価結果を返す)
        //                    case CellType.Formula:
        //                        IFormulaEvaluator evaluator
        //                            = Parent.Parent.PoiBook.GetCreationHelper().CreateFormulaEvaluator();
        //                        CellValue cellValue = evaluator.Evaluate(cell);
        //                        if (cellValue.CellType == CellType.String)
        //                            CelVal = cellValue.StringValue;
        //                        else
        //                            CelVal = cellValue.NumberValue.ToString();
        //                        break;
        //                    //エラー
        //                    case CellType.Error:
        //                        CelVal = cell.ErrorCellValue.ToString();
        //                        break;
        //                    //空白
        //                    case CellType.Blank:
        //                        CelVal = string.Empty;
        //                        break;
        //                    //その他
        //                    default:
        //                        CelVal = string.Empty;
        //                        break;
        //                }
        //                Values[
        //                    RIdx - SafeAddress.FirstRow + 1,
        //                    CIdx - SafeAddress.FirstColumn + 1] = CelVal;
        //            }
        //        }
        //        //単一セルなら配列ではなく値そのものでリターン
        //        if (Values.Length == 1)
        //        {
        //            return Values[1, 1];
        //        }
        //        return Values;
        //    }
        //    set
        //    {
        //        bool PasteArray = false;
        //        int ValueFirstRow = 0;
        //        int ValueFirstColumn = 0;
        //        //Office.Interop.Excelにならい非連続Rangeの全てに適用
        //        for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
        //        {
        //            //アドレス取得
        //            CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
        //            //供給された値が配列の場合
        //            if (value.GetType().IsArray)
        //            {
        //                //２次元ならばRangeペースト処理を設定
        //                if (((Array)value).Rank == 2)
        //                {
        //                    ValueFirstRow = ((Array)value).GetLowerBound(0);
        //                    ValueFirstColumn = ((Array)value).GetLowerBound(1);
        //                    PasteArray = true;
        //                }
        //            }
        //            //行ループ
        //            for (int RIdx = 0; RIdx <= SafeAddress.LastRow - SafeAddress.FirstRow; RIdx++)
        //            {
        //                //行の取得(なければ生成)
        //                IRow row = Parent.PoiSheet.GetRow(RIdx + SafeAddress.FirstRow)
        //                            ?? Parent.PoiSheet.CreateRow(RIdx + SafeAddress.FirstRow);
        //                //列ループ
        //                for (int CIdx = 0; CIdx <= SafeAddress.LastColumn - SafeAddress.FirstColumn; CIdx++)
        //                {
        //                    //列の取得(なければ生成)
        //                    ICell cell = row.GetCell(CIdx + SafeAddress.FirstColumn)
        //                                    ?? row.CreateCell(CIdx + SafeAddress.FirstColumn);
        //                    //セットする値の特定
        //                    object CValue = value;
        //                    Cell.ValueType CType = Cell.ValueType.Auto;
        //                    //Rangeペースト処理の場合は配列から値を取得
        //                    if (PasteArray)
        //                    {
        //                        CValue = value[RIdx + ValueFirstRow, CIdx + ValueFirstColumn];
        //                        //配列要素がCellクラスなら解読
        //                        if (CValue is Cell)
        //                        {
        //                            Cell c = value[RIdx + ValueFirstRow, CIdx + ValueFirstColumn];
        //                            CValue = c.Value;
        //                            CType = c.Type;
        //                        }
        //                    }
        //                    //文字列固定の場合
        //                    if (CType == Cell.ValueType.String)
        //                    {
        //                        cell.SetCellValue((string)CValue);
        //                        cell.SetCellType(CellType.String);
        //                    }
        //                    //式固定の場合
        //                    else if (CType == Cell.ValueType.Formula)
        //                    {
        //                        cell.SetCellFormula((string)CValue);
        //                        cell.SetCellType(CellType.Formula);
        //                    }
        //                    else
        //                    {
        //                        //日付であっても数値としてセット(ユーザによる書式設定を期待する)
        //                        if (DateTime.TryParse(CValue.ToString(), out DateTime dtm))
        //                        {
        //                            cell.SetCellValue((DateTime)dtm);
        //                            cell.SetCellType(CellType.Numeric);
        //                        }
        //                        //数値であれば数値としてセット
        //                        else if (double.TryParse(CValue.ToString(), out double dbl))
        //                        {
        //                            cell.SetCellValue((double)dbl);
        //                            cell.SetCellType(CellType.Numeric);
        //                        }
        //                        //その他は文字列扱い
        //                        else
        //                        {
        //                            cell.SetCellValue((string)CValue);
        //                            cell.SetCellType(CellType.String);
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// レンジの値       !!!!!interface上はobject
        /// </summary>
        public object Value2
        {
            get { return RangeManager.Value2; }
            set { RangeManager.Value2 = value; }
        }


        /// <summary>
        /// セルの文字列(セットのみ)
        /// </summary>
        public object Text
        {
            get { return RangeManager.Text; }
        }
        //public object Text
        //{
        //    ///セルの値設定
        //    set
        //    {
        //        //Office.Interop.Excelにならい非連続Rangeの全てに適用
        //        for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
        //        {
        //            //アドレス取得
        //            CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
        //            //行ループ
        //            for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
        //            {
        //                //行の取得(なければ生成)
        //                IRow row = Parent.PoiSheet.GetRow(RIdx) ?? Parent.PoiSheet.CreateRow(RIdx);
        //                //列ループ
        //                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
        //                {
        //                    //列の取得(なければ生成)
        //                    ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
        //                    cell.SetCellValue((string)value);
        //                    cell.SetCellType(CellType.String);
        //                }
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// セルの式(セットのみ)
        /// </summary>
        public object Formula
        {
            get { return RangeManager.Formula; }
            set { RangeManager.Formula = value; }
        }

        //public string Formula
        //{
        //    ///セルの値設定
        //    set
        //    {
        //        string Formula = value;
        //        Formula = Formula.TrimStart('=');
        //        //Office.Interop.Excelにならい非連続Rangeの全てに適用
        //        for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
        //        {
        //            //アドレス取得
        //            CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
        //            //行ループ
        //            for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
        //            {
        //                //行の取得(なければ生成)
        //                IRow row = Parent.PoiSheet.GetRow(RIdx) ?? Parent.PoiSheet.CreateRow(RIdx);
        //                //列ループ
        //                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
        //                {
        //                    //列の取得(なければ生成)
        //                    ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
        //                    cell.SetCellFormula(Formula);
        //                    cell.SetCellType(CellType.Formula);
        //                }
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// Rangeの行高さ合計(単位はPoint)
        /// </summary>
        public object Height
        {
            //Rangeに含まれる行の高さ合計値
            get
            {
                //Office.Interop.Excelにならい先頭アドレスのみ参照
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
                float RetVal = 0;
                //行ループ
                for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                {
                    //行の取得(なければデフォルト値を採用)
                    IRow row = Parent.PoiSheet.GetRow(RIdx);
                    if (row != null)
                    {
                        RetVal += row.HeightInPoints;
                    }
                    else
                    {
                        //twipなので20倍してpointに編案
                        RetVal += (Parent.PoiSheet.DefaultRowHeight * 20);
                    }
                }
                return RetVal;
            }
        }

        /// <summary>
        /// Range各行の高さ(単位はPoint)
        /// </summary>
        public object RowHeight
        {
            get
            {
                object RetVal = null;
                List<float> ht = new List<float>();
                //Office.Interop.Excelにならい先頭アドレスのみ参照
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
                //行ループ
                for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                {
                    //行の取得(なければデフォルト値を採用)
                    IRow row = Parent.PoiSheet.GetRow(RIdx);
                    if (row != null)
                    {
                        ht.Add(row.HeightInPoints);
                    }
                    else
                    {
                        //twipなので20倍してpointに編案
                        ht.Add(Parent.PoiSheet.DefaultRowHeight * 20);
                    }
                    //違う高さが検出されたらbreak
                    if (ht.Min() != ht.Max())
                    {
                        break;
                    }
                }
                //全行が同じ高さならその高さでリターン
                if (ht.Min() == ht.Max())
                {
                    RetVal = ht.Min();
                }
                return RetVal;
            }
            set
            {
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //行ループ
                    for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                    {
                        //行の取得(なければ生成)
                        IRow row = Parent.PoiSheet.GetRow(RIdx);
                        if (row == null)
                        {
                            row = Parent.PoiSheet.CreateRow(RIdx);
                            Logger.Debug(
                                "Sheet[" + Parent.PoiSheet.SheetName + "]:Row[" + RIdx + "] *** Row Created. ***");
                        }
                        //高さを設定
                        row.HeightInPoints = (float)value;
                    }
                }
            }
        }

        /// <summary>
        /// Rangeの列幅合計(単位は文字幅の1/20を1とする値であり、Pointではない)
        /// </summary>
        public object Width
        {
            //Rangeに含まれる列の幅合計値
            get
            {
                float RetVal = 0;
                //Office.Interop.Excelにならい先頭アドレスのみ参照
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
                //列ループ
                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                {
                    RetVal += Parent.PoiSheet.GetColumnWidth(CIdx);
                }
                return RetVal;
            }
        }

        /// <summary>
        /// Range各列の幅(単位は文字幅の1/20を1とする値であり、Pointではない)
        /// </summary>
        public object ColumnWidth
        {
            get
            {
                object RetVal = null;
                List<int> wd = new List<int>();
                //Office.Interop.Excelにならい先頭アドレスのみ参照
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
                //列ループ
                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                {
                    wd.Add(Parent.PoiSheet.GetColumnWidth(CIdx));
                    //違う幅さが検出されたらbreak
                    if (wd.Min() != wd.Max())
                    {
                        break;
                    }
                }
                //全列が同じ幅ならその幅でリターン
                if (wd.Min() == wd.Max())
                {
                    RetVal = wd.Min();
                }
                return RetVal;
            }
            set
            {
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //列ループ
                    for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                    {
                        Parent.PoiSheet.SetColumnWidth(CIdx, (int)value);
                    }
                }
            }
        }

        /// <summary>
        /// 罫線情報
        /// </summary>
        public Borders Borders { get { return RangeManager.Borders; } }

        /// <summary>
        /// 文字フォント情報
        /// </summary>
        public Font Font { get { return RangeManager.Font; } }

        /// <summary>
        /// セルの内部(塗りつぶし)
        /// </summary>
        public Interior Interior { get { return RangeManager.Interior; } }

        /// <summary>
        /// 文字位置(水平方向)
        /// </summary>
        public object HorizontalAlignment
        {
            get { return RangeManager.HorizontalAlignment; }
            set { RangeManager.HorizontalAlignment = value; }
        }

        /// <summary>
        /// 文字位置(垂直方向)
        /// </summary>
        public object VerticalAlignment
        {
            get { return RangeManager.VerticalAlignment; }
            set { RangeManager.VerticalAlignment = value; }
        }

        /// <summary>
        /// 文字書式
        /// </summary>
        public object NumberFormat
        {
            get { return RangeManager.NumberFormat; }
            set { RangeManager.NumberFormat = value; }
        }

        /// <summary>
        /// 文字書式(本来はLocaleを反映するが現在未サポートにつきNumberFormatと同じ)
        /// </summary>
        public object NumberFormatLocal
        {
            get { return RangeManager.NumberFormatLocal; }
            set { RangeManager.NumberFormatLocal = value; }
        }

        /// <summary>
        /// セルの保護
        /// </summary>
        public object Locked
        {
            get { return RangeManager.Locked; }
            set { RangeManager.Locked = value; }
        }

        /// <summary>
        /// 文字列の折り返し
        /// </summary>
        public object WrapText
        {
            get { return RangeManager.WrapText; }
            set { RangeManager.WrapText = value; }
        }

        /// <summary>
        /// コメント
        /// </summary>
        public Comment Comment
        {
            get { return RangeManager.Comment; }
        }

        /// <summary>
        /// 入力規則
        /// </summary>
        public Validation Validation
        {
            get { return RangeManager.Validation; }
        }



        #endregion

        #region "internal properties"

        /// <summary>
        /// Rangeアドレス(-1を含む)
        /// </summary>
        internal CellRangeAddressList RawAddressList { get; }
        /// <summary>
        /// Rangeアドレス(-1を含まないのでforで安全にアクセスできる)
        /// </summary>
        internal CellRangeAddressList SafeAddressList { get; }
        /// <summary>
        /// 相対アドレスの基点アドレス
        /// </summary>
        internal CellRangeAddress RelativeTo { get; }

        #endregion

        #region "private properties"

        /// <summary>
        /// Countプロパティが示す値の種別(セル数、行数、列数)
        /// </summary>
        private CountType CountAs { get; } = CountType.Default;
        /// <summary>
        /// スタイル情報管理クラスインスタンスへのアクセサ
        /// </summary>
        private RangeManager RangeManager
        {
            get
            {
                //最初にアクセスされたときにインスタンス化する
                if (_RangeManager == null) { _RangeManager = new RangeManager((Range)this); }
                return _RangeManager;
            }
        }

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="Cell1"></param>
        /// <param name="Cell2"></param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public virtual Range this[object Cell1, object Cell2 = null]
        {
            get
            {
                //アドレス計算用リスト初期化
                CellRangeAddressList AddressList = new CellRangeAddressList();
                //Cells指定の場合
                if (Cell1 is Range cell1)
                {
                    //Cell1が単一セルであること
                    if (cell1.Count == 1)
                    {
                        AddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(cell1.Address));
                    }
                    //上記以外は例外スロー
                    else
                    {
                        throw new ArgumentException("Cell1 contains multiple cells.");
                    }
                    //Cell2の指定があること
                    if (Cell2 != null && Cell2 is Range cell2)
                    {
                        //Cell2が単一セルであること
                        if (cell2.Count == 1)
                        {
                            AddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(cell2.Address));
                        }
                        //上記以外は例外スロー
                        else
                        {
                            throw new ArgumentException("Cell2 contains multiple cells.");
                        }
                    }
                    //Cell2の指定がなければ例外スロー
                    else
                    {
                        throw new ArgumentException("In case type of Cell1 is Cells, Type of Cell2 must be Cells.");
                    }
                    //アドレスの統合
                    AddressList = RangeUtil.CreateMergedAddressList(AddressList);
                }
                //Cell1がStringの場合(A1形式)
                else if (Cell1 is string adr1)
                {
                    string[] AdrLst1 = adr1.Split(',');
                    //複数アドレスの場合
                    if (AdrLst1.Length > 1)
                    {
                        //Cell1の複数アドレスをそのまま使用
                        foreach (string adr in AdrLst1)
                        {
                            AddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(adr));
                        }
                        //Cell2があれば例外スロー
                        if (Cell2 != null)
                        {
                            throw new ArgumentException("In case Cell1 has multiple cells, Cell2 must be null.");
                        }
                    }
                    //単一アドレスの場合
                    else
                    {
                        //Cell1(A1形式)からアドレス生成しアレイに追記
                        AddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(adr1));
                        //Cell2がStringの場合(A1形式)
                        if (Cell2 != null && Cell2 is string adr2)
                        {
                            string[] AdrLst2 = adr2.Split(',');
                            //単一アドレスなら採用
                            if (AdrLst2.Length == 1)
                            {
                                //Cell2(A1形式)からアドレス生成しアレイに追記
                                AddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(adr2));
                            }
                            //複数アドレスなら例外スロー
                            else
                            {
                                throw new ArgumentException("Cell2 contains multiple cells.");
                            }
                        }
                        //アドレスの統合
                        AddressList = RangeUtil.CreateMergedAddressList(AddressList);
                    }
                }
                //Cellsでもstringでもなければ例外スロー
                else
                {
                    throw new ArgumentException("Type of Cell1 must be Cells or string.");
                }
                //Rangeクラスインスタンス生成
                return new Range(Parent, AddressList, RelativeTo);
            }
        }

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// Rangeの選択
        /// </summary>
        public void Select()
        {
            //親BookのWindowにこのRangeをセットする。
            this.Parent.Parent.Windows[1].RangeSelection = (Range)this;
            //ApplicationにこのRangeをセットする。
            this.Application.SetSelection(this.Parent, this);
        }

        /// <summary>
        /// セルのコメントを生成する
        /// </summary>
        /// <param name="CommentText">コメント文字列</param>
        public Comment AddComment(object Text = null)
        {
            return RangeManager.AddComment(Text);
        }

        /// <summary>
        /// 列幅の自動調整
        /// </summary>
        public void AutoFit()
        {
            //デバッグログ用情報
            int FitCount = 0;
            var StopwatchForDebugLog = new System.Diagnostics.Stopwatch();
            StopwatchForDebugLog.Start();
            //行モード
            if (this.CountAs == CountType.Rows)
            {
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //行高自動調整ループ
                    for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                    {
                        //処理数インクリメント
                        FitCount++;
                        IRow Row = Parent.PoiSheet.GetRow(RIdx);
                        if(Row != null)
                        {
                            Row.Height = -1;
                        }
                    }
                }
                //処理時間測定タイマー停止＆ログ出力
                StopwatchForDebugLog.Stop();
                TimeSpan TimeSpanForDebugLog = StopwatchForDebugLog.Elapsed;
                Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for [" + FitCount + "]Rows");
            }
            //列モード
            else
            {
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //列幅自動調整ループ
                    for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                    {
                        //処理数インクリメント
                        FitCount++;
                        //スッピンのAutoSizeでは独自書式(例えば通貨)の幅が少し足りない。
                        //"\"や"､"の増分が考慮されていないような感じ。
                        //ある程度救済するため、一律28%増の処理を行う
                        Parent.PoiSheet.AutoSizeColumn(CIdx);
                        Parent.PoiSheet.SetColumnWidth(
                            CIdx, Parent.PoiSheet.GetColumnWidth(CIdx) * 128 / 100);
                        //処理対象の行が大量の場合に効果があるらしいが、弊害もありそうなのでやめておく。
                        //GC.Collect();
                    }
                }
                //処理時間測定タイマー停止＆ログ出力
                StopwatchForDebugLog.Stop();
                TimeSpan TimeSpanForDebugLog = StopwatchForDebugLog.Elapsed;
                Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for [" + FitCount + "]Columns");
                ProcTimeLogger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for [" + FitCount + "]Columns");
            }
        }

        /// <summary>
        /// 囲み罫線の設定
        /// </summary>
        /// <param name="LineStyle">線の種類</param>
        /// <param name="Weight">線の太さ</param>
        /// <param name="ColorIndex">カラーパレット上の色インデックス。NPOIのColorIndexはInterop.Excelのそれと異なるので要注意。</param>
        /// <param name="Color">未サポート</param>
        /// <returns></returns>
        public object BorderAround(
                            object LineStyle = null, XlBorderWeight Weight = XlBorderWeight.xlThin,
                            XlColorIndex ColorIndex = XlColorIndex.xlColorIndexAutomatic, object Color = null)
        {
            return RangeManager.BorderAround(LineStyle, Weight, ColorIndex, Color);
        }

        /// <summary>
        /// 指定された条件に合致するRangeを取得する
        /// </summary>
        /// <param name="Type">指定条件</param>
        /// <param name="Value">条件パラメータ</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public Range SpecialCells(XlCellType Type = XlCellType.xlCellTypeLastCell, object Value = null)
        {
            Range RetVal;
            int RowIndex = 0;
            int ColumnIndex = 0;
            //先頭アドレス取得
            CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
            //XlCellType.xlCellTypeLastCellのみ処理
            if (Type == XlCellType.xlCellTypeLastCell)
            {
                int LastRowIndex = SafeAddress.LastRow;
                //Range最終行から上に向かって検索
                for (int CIdx = 0; CIdx >= 0; CIdx--)
                {
                    //行が存在すれば列をチェック
                    IRow row = Parent.PoiSheet.GetRow(LastRowIndex + CIdx);
                    if (row != null)
                    {
                        //列が存在するならその列を採用
                        if (row.PhysicalNumberOfCells > 0)
                        {
                            RowIndex = LastRowIndex + CIdx;
                            ColumnIndex = row.LastCellNum - 1;
                            break;
                        }
                    }
                }
                //最終カラムのRangeでリターン
                RetVal = new Range(
                    Parent, new CellRangeAddressList(RowIndex, RowIndex, ColumnIndex, ColumnIndex));
            }
            else
            {
                //ダミーアクセス
                if (Value == null) { }
                //例外スロー
                throw new ArgumentException("SpecialCells supports XlCellType.xlCellTypeLastCell only.");
            }
            return RetVal;
        }

        #endregion

        #region "alternative public methods"

        /// <summary>
        /// セルのバリデーションを生成する
        /// </summary>
        /// <param name="ExplicitList">値リスト(srting[])</param>
        /// <param name="ShowPronptBox">プロンプト表示有無</param>
        /// <param name="PronptBoxTitle">プロンプトタイトル</param>
        /// <param name="PronptBoxText">プロンプト本文</param>
        /// <param name="ShowErrorBox">バリデーションエラー時のエラーボックス表示有無</param>
        /// <param name="ErrorBoxTitle">エラーボックスタイトル</param>
        /// <param name="ErrorBoxText">エラーボックス本文</param>
        public void AddValidation(
            string[] ExplicitList,
            bool ShowPronptBox = true, string PronptBoxTitle = "値選択", string PronptBoxText = "値を選択してください。",
            bool ShowErrorBox = true, string ErrorBoxTitle = "入力エラー", string ErrorBoxText = "正しい値を選択してください。")
        {
            IDataValidationConstraint Cst
                = Parent.PoiSheet.GetDataValidationHelper().CreateExplicitListConstraint(ExplicitList);
            //Office.Interop.Excelにならい非連続Rangeの全てに適用(RawAddressList)
            IDataValidation Val
                = Parent.PoiSheet.GetDataValidationHelper().CreateValidation(Cst, RawAddressList);
            //なぜかHSSFとXSFでは指定が逆になっている模様
            if (Parent.PoiSheet is HSSFSheet)
            {
                Val.SuppressDropDownArrow = false;
            }
            else
            {
                //なぜかサプレスTRUEにすると表示される
                Val.SuppressDropDownArrow = true;
            }
            Val.ErrorStyle = ERRORSTYLE.STOP;
            Val.ShowErrorBox = ShowErrorBox;
            Val.CreateErrorBox(ErrorBoxTitle, ErrorBoxText);
            Val.ShowPromptBox = ShowPronptBox;
            Val.CreatePromptBox(PronptBoxTitle, PronptBoxText);
            Parent.PoiSheet.AddValidationData(Val);
        }

        /// <summary>
        /// Rangeの書式設定
        /// 設定可能なスタイル数には上限があるため、予めPoiWrapper.configに設定しておき、それを使いまわす。
        /// </summary>
        /// <param name="StyleName">PoiWrapper.configで指定したスタイル名</param>
        public void SetStyle(string StyleName)
        {
            //Office.Interop.Excelにならい非連続Rangeの全てに適用
            for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
            {
                //アドレス取得
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                //スタイルリストに存在する場合はその設定を適用
                if (Parent.Parent.CellStyles.ContainsKey(StyleName))
                {
                    //設定の取得
                    ICellStyle Style = Parent.Parent.CellStyles[StyleName];
                    //行ループ
                    for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                    {
                        //行の取得(なければ生成)
                        IRow row = Parent.PoiSheet.GetRow(RIdx);
                        if (row == null)
                        {
                            row = Parent.PoiSheet.CreateRow(RIdx);
                            Logger.Debug(
                                "Sheet[" + Parent.PoiSheet.SheetName + "]:Row[" + RIdx + "] *** Row Created. ***");
                        }
                        //列ループ
                        for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                        {
                            //列の取得(なければ生成)
                            ICell cell = row.GetCell(CIdx);
                            if(cell == null)
                            {
                                cell = row.CreateCell(CIdx);
                                Logger.Debug(
                                    "Sheet[" + Parent.PoiSheet.SheetName + "]:Cell[" + RIdx + "][" + CIdx + "] *** Column Created. ***");
                            }
                            //スタイルの適用
                            cell.CellStyle = Style;
                        }
                    }
                }
            }
        }

        #endregion

        #endregion
    }
}
