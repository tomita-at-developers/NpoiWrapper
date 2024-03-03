using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Validation interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Validation
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    int AlertStyle { get; }
    //    bool IgnoreBlank { get; set; }
    //    int IMEMode { get; set; }
    //    bool InCellDropdown { get; set; }
    //    string ErrorMessage { get; set; }
    //    string ErrorTitle { get; set; }
    //    string InputMessage { get; set; }
    //    string InputTitle { get; set; }
    //    string Formula1 { get; }
    //    string Formula2 { get; }
    //    int Operator { get; }
    //    bool ShowError { get; set; }
    //    bool ShowInput { get; set; }
    //    int Type { get; }
    //    bool Value { get; }
    //    void Add(XlDVType Type, [Optional] object AlertStyle, [Optional] object Operator, [Optional] object Formula1, [Optional] object Formula2);
    //    void Delete();
    //    void Modify([Optional] object Type, [Optional] object AlertStyle, [Optional] object Operator, [Optional] object Formula1, [Optional] object Formula2);
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding interface in NPOI IDataValidation is shown below...
    //----------------------------------------------------------------------------------------------
    //  ＜ 注意＞
    //  inteerpo.excelでは、指定されたRangeが複数のvalidationを持つ場合、取得されるvalidationは不定となり、アクセスすると例外となる。
    //  NPOIでは、Validationの更新削除の関するmethod類がSS系に存在しない。よってHSS/XSSFそれぞれで対処する必要がある。
    //  XSSF : List<NPOI.OpenXmlFormats.Spreadsheet.CT_DataValidation> ValidationList = ((XSSFSheet)Sheet).GetCTWorksheet().dataValidations.dataValidation
    //  HSSF : ???
    //public interface IDataValidation
    //{
    //    //入力規則の定義
    //    IDataValidationConstraint ValidationConstraint { get; }
    //    //エラーメッセージのスタイル：Excelの[(エラーメッセージタブ)スタイル]に相当。
    //    //NPOI.SS.UserModel.ERRORSTYLEクラスで定義されている以下のconstを指定
    //    //  public const int STOP = 0x00;
    //    //  public const int WARNING = 0x01;
    //    //  public const int INFO = 0x02;
    //    int ErrorStyle { get; set; }
    //    //空白セルの許可：Excelの[空白を無視する]と同義
    //    bool EmptyCellAllowed { get; set; }
    //    //ドロップダウンリストの抑制：Excelの[ドロップダウンリストから選択する]に相当。
    //    //  HSSFではtrueで非表示、XSSFではtrueで表示となり、意味合いが逆になっている。
    //    bool SuppressDropDownArrow { get; set; }
    //    //入力時メッセージ：Excelの[セルを選択した時に入力時メッセージを表示する]に相当。
    //    bool ShowPromptBox { get; set; }
    //    //エラーメッセージ：Excelの[無効なデータが入力されたらエラーメッセージを表示する]に相当。
    //    bool ShowErrorBox { get; set; }
    //    //入力時メッセージのタイトル
    //    string PromptBoxTitle { get; }
    //    //入力時メッセージのメッセージ
    //    string PromptBoxText { get; }
    //    //エラーメッセージのタイトル
    //    string ErrorBoxTitle { get; }
    //    //エラーメッセージのメッセージ
    //    string ErrorBoxText { get; }
    //    //適用Range
    //    CellRangeAddressList Regions { get; }
    //    void CreatePromptBox(string title, string text);
    //    void CreateErrorBox(string title, string text);
    //}
    //----------------------------------------------------------------------------------------------
    //IDataValidationConstraint ValidationConstraint
    //----------------------------------------------------------------------------------------------
    //IDataValidationHelper ISheet.GetDataValidationHelper()が提供する各種メソッドで設定する。
    //Excelとの対応関係(XDVTypeの対応)は以下の通り。
    //public enum XlDVType
    //{
    //    xlValidateInputOnly,
    //    xlValidateWholeNumber,
    //    xlValidateDecimal,
    //    xlValidateList,
    //    xlValidateDate,
    //    xlValidateTime,
    //    xlValidateTextLength,
    //    xlValidateCustom
    //}
    //----------------------------------------------------------------------------------------------
    //すべての値：xlValidateInputOnly
    //----------------------------------------------------------------------------------------------
    //整数：xlValidateWholeNumber
    //      データ(OperatorType), 最小値(formura1), 最大値(formula2), 空白を無視する
    //      CreateintConstraint(int operatorType, string formula1, string formula2)
    //      CreateNumericConstraint(int validationType, int operatorType, string formula1, string formula2)
    //----------------------------------------------------------------------------------------------
    //小数点数：xlValidateDecimal
    //      データ(OperatorType), 最小値(formura1), 最大値(formula2), 空白を無視する
    //      CreateDecimalConstraint(int operatorType, string formula1, string formula2)
    //----------------------------------------------------------------------------------------------
    //リスト：xlValidateList
    //      元の値(formula), 空白を無視する, ドロップダウンリストから選択する
    //      CreateExplicitListConstraint(string[] listOfValues)：即値指定
    //      CreateFormulaListConstraint(string listFormula)：レンジ指定
    //----------------------------------------------------------------------------------------------
    //日付：xlValidateDate
    //      データ(OperatorType), 開始日(formura1), 終了日(formula2), 空白を無視する
    //      CreateDateConstraint(int operatorType, string formula1, string formula2, string dateFormat)
    //----------------------------------------------------------------------------------------------
    //時刻：xlValidateTime
    //      データ(OperatorType), 開始時刻(formura1), 終了時刻(formula2), 空白を無視する
    //      CreateTimeConstraint(int operatorType, string formula1, string formula2)
    //----------------------------------------------------------------------------------------------
    //文字列(長さ指定)：xlValidateTextLength
    //      データ(OperatorType), 最小値(formura1), 最大値(formula2), 空白を無視する
    //      CreateTextLengthConstraint(int operatorType, string formula1, string formula2)
    //----------------------------------------------------------------------------------------------
    //ユーザー設定：xlValidateCustom
    //      数式(formula), 空白を無視する
    //      CreateCustomConstraint(string formula)
    //----------------------------------------------------------------------------------------------
    //public static class ValidationType  <--- CreateNumericConstraintメソッドの引数validationType
    //{
    //    public const int ANY = 0;
    //    public const int INTEGER = 1;
    //    public const int DECIMAL = 2;
    //    public const int LIST = 3;
    //    public const int DATE = 4;
    //    public const int TIME = 5;
    //    public const int TEXT_LENGTH = 6;
    //    public const int FORMULA = 7;
    //}
    //----------------------------------------------------------------------------------------------
    //public static class OperatorType    <--- 各Constraintメソッドの引数operatorType
    //{
    //    public const int BETWEEN = 0;
    //    public const int NOT_BETWEEN = 1;
    //    public const int EQUAL = 2;
    //    public const int NOT_EQUAL = 3;
    //    public const int GREATER_THAN = 4;
    //    public const int LESS_THAN = 5;
    //    public const int GREATER_OR_EQUAL = 6;
    //    public const int LESS_OR_EQUAL = 7;
    //    public const int IGNORED = 0;
    //}
    //----------------------------------------------------------------------------------------------

    /// <summary>
    /// Validationクラス
    /// </summary>
    public class Validation
    {
        #region "fields"

        /// <summary>
        /// このオブジェクトにアクセス可能か否か
        /// Interop.Excelの挙動に似せるため＆ニセの初期値を見せないため。
        /// </summary>
        private bool _Accessable = false;
        /// <summary>
        /// IgnoreBlankの実体
        /// </summary>
        private bool _IgnoreBlank = true;
        /// <summary>
        /// InCellDropdownの実体
        /// </summary>
        private bool _InCellDropdown = true;
        /// <summary>
        /// ShowInputの実体
        /// </summary>
        private bool _ShowInput = false;
        /// <summary>
        /// InputTitleの実体
        /// </summary>
        private string _InputTitle = string.Empty;
        /// <summary>
        /// InputMessageの実体
        /// </summary>
        private string _InputMessage = string.Empty;
        /// <summary>
        /// ShowErrorの実体
        /// </summary>
        private bool _ShowError = false;
        /// <summary>
        /// ErrorTitleの実体
        /// </summary>
        private string _ErrorTitle = string.Empty;
        /// <summary>
        /// ErrorMessageの実体
        /// </summary>
        private string _ErrorMessage = string.Empty;
        /// <summary>
        /// IMEModeの実体
        /// </summary>
        private int _IMEMode = 0;
        /// <summary>
        /// AlertStyleの実体
        /// </summary>
        private int _AlertStyle = (int)XlDVAlertStyle.xlValidAlertInformation;
        /// <summary>
        /// Typeの実体
        /// </summary>
        private int _Type = (int)XlDVType.xlValidateInputOnly;
        /// <summary>
        /// Operatorの実体
        /// </summary>
        private int _Operator = (int)XlFormatConditionOperator.xlEqual;
        /// <summary>
        /// Formula1の実体
        /// </summary>
        private string _Formula1 = null;
        /// <summary>
        /// Formula2の実体
        /// </summary>
        private string _Formula2 = null;

        /// <summary>
        /// Valueの実体
        /// </summary>
        private bool _Value = true;

        /// <summary>
        /// このRangeが属しているValidationのリスト
        /// </summary>
        private List<IDataValidation> _BelongsTo = new List<IDataValidation>();

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        internal Validation(Range ParentRange)
        {
            //親Range情報保存
            this.Parent = ParentRange;
            //Validation情報の読込
            GetProperties();
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// 空白セルの許可
        /// </summary>
        public bool IgnoreBlank
        {
            get
            {
                return _IgnoreBlank;
            }
            set
            {
                _IgnoreBlank = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// ドロップダウンリストの表示
        /// </summary>
        public bool InCellDropdown
        {
            get
            {
                return _InCellDropdown;
            }
            set
            {
                _InCellDropdown = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// 入力メッセージの表示
        /// </summary>
        public bool ShowInput
        {
            get
            {
                return _ShowInput;
            }
            set
            {
                _ShowInput = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// 入力メッセージのタイトル
        /// </summary>
        public string InputTitle
        {
            get
            {
                return _InputTitle;
            }
            set
            {
                _InputTitle = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// 入力メッセージの本文
        /// </summary>
        public string InputMessage
        {
            get
            {
                return _InputMessage;
            }
            set
            {
                _InputMessage = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// エラーメッセージの表示
        /// </summary>
        public bool ShowError
        {
            get
            {
                return _ShowError;
            }
            set
            {
                _ShowError = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// エラーメッセージのタイトル
        /// </summary>
        public string ErrorTitle
        {
            get
            {
                return _ErrorTitle;
            }
            set
            {
                _ErrorTitle = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// エラーメッセージの本文
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return _ErrorMessage;
            }
            set
            {
                _ErrorMessage = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// IMEモード(未サポート(0:xlIMEModeNoControl固定)
        /// </summary>
        public int IMEMode
        {
            get
            {
                return _IMEMode;
            }
            set
            {
                _IMEMode = value;
                Apply();
                GetProperties();
            }
        }

        /// <summary>
        /// エラーメッセージのスタイル
        /// </summary>
        public int AlertStyle
        {
            get
            {
                AccessFilter();
                return _AlertStyle;
            }
            private set
            {
                _AlertStyle = value;
            }
        }

        /// <summary>
        /// 入力規則の種類
        /// </summary>
        public int Type
        {
            get
            {
                AccessFilter();
                return _Type;
            }
            private set
            {
                _Type = value;
            }
        }

        /// <summary>
        /// 評価条件の種類
        /// </summary>
        public int Operator
        {
            get
            {
                AccessFilter();
                return _Operator;
            }
            private set
            {
                _Operator = value;
            }
        }

        /// <summary>
        /// 評価パラメータ１
        /// </summary>
        public string Formula1
        {
            get
            {
                AccessFilter();
                return _Formula1;
            }
            private set
            {
                _Formula1 = value;
            }
        }

        /// <summary>
        /// 評価パラメータ２
        /// </summary>
        public string Formula2
        {
            get
            {
                AccessFilter();
                return _Formula2;
            }
            private set
            {
                _Formula2 = value;
            }
        }

        /// <summary>
        /// 全レンジの入力規則評価結果
        /// </summary>
        public bool Value
        {
            get
            {
                AccessFilter();
                return _Value;
            }
            private set
            {
                _Value = value;
            }
        }

        #endregion

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// 入力規則の追加(既にあれば更新)
        /// </summary>
        /// <param name="Type">XlDVTypeで指定する入力規則の種類</param>
        /// <param name="AlertStyle">XlDVAlertStyleで指定するエラーメッセージスタイル</param>
        /// <param name="Operator">XlFormatConditionOperatorで指定する値評価の種類</param>
        /// <param name="Formula1">最小値等の評価パラメータ</param>
        /// <param name="Formula2">最大値等の評価パラメータ</param>
        public void Add(XlDVType Type, object AlertStyle = null, object Operator = null, object Formula1 = null, object Formula2 = null)
        {
            //最新の状態に更新
            GetProperties();
            //このRangeを支配するValidationがなければ処理する
            if (this._BelongsTo.Count == 0)
            {

                //指定されたパラメータの保存(AddのTypeは型指定なので信頼する)
                SetParamsToProperties(Type, AlertStyle, Operator, Formula1, Formula2);
                //Validationの適用
                Apply();
            }
            else
            {
                throw new InvalidOperationException("This Range already has a validation to subject to.");
            }
            //最新の状態に更新
            GetProperties();
        }

        /// <summary>
        /// Validationの削除
        /// </summary>
        public void Delete()
        {
            //最新の状態に更新
            GetProperties();
            //このRangeを支配するValidationがあれば処理する
            if (this._BelongsTo.Count > 0)
            {
                //このRangeからValidationを除去
                UnlinkOrDeleteValidation();
            }
        }

        /// <summary>
        /// 入力規則の更新(なければ追加)
        /// </summary>
        /// <param name="Type">XlDVTypeで指定する入力規則の種類</param>
        /// <param name="AlertStyle">XlDVAlertStyleで指定するエラーメッセージスタイル</param>
        /// <param name="Operator">XlFormatConditionOperatorで指定する値評価の種類</param>
        /// <param name="Formula1">最小値等の評価パラメータ。Cell/Rangeへの参照を指定する場合、参照先は先は同一シート内、またはブックを範囲として指定された名前付きCell/Rangeでなければならない。</param>
        /// <param name="Formula2">最大値等の評価パラメータ。Cell/Rangeへの参照を指定する場合、参照先は先は同一シート内、またはブックを範囲として指定された名前付きCell/Rangeでなければならない。</param>
        public void Modify(object Type = null, object AlertStyle = null, object Operator = null, object Formula1 = null, object Formula2 = null)
        {
            //最新の状態に更新
            GetProperties();
            //このRangeを支配するValidationがあれば処理する
            if (this._BelongsTo.Count > 0)
            {
                //ModifyのTypeは[optional]objectなので存在と正当性をチェックする
                int ValidationType = this._Type;
                if (Type is XlDVType SafeXlDVType)
                {
                    ValidationType = (int)SafeXlDVType;
                }
                if (Type is int SafeInt)
                {
                    if (Enum.IsDefined(typeof(XlDVType), SafeInt))
                    {
                        ValidationType = SafeInt;
                    }
                }
                //指定されたパラメータの保存
                SetParamsToProperties((XlDVType)ValidationType, AlertStyle, Operator, Formula1, Formula2);
                //Validationの適用
                Apply();
            }
            else
            {
                throw new InvalidOperationException("This Range has no validation to modify.");
            }
            //最新の状態に更新
            GetProperties();
        }

        /// <summary>
        /// このシートのValidationsからこのRangeに関する情報を取り出し、プロパティにセットする。
        /// </summary>
        private void GetProperties()
        {
            //このRangeが属するValidationリストの初期化
            this._BelongsTo = new List<IDataValidation>();
            //所存Validationが見つかっていないRange(初期値は全Range)
            CellRangeAddressList Remaining = Parent.SafeAddressList;
            //Validation情報取得(SSレベルのClone情報)
            List <IDataValidation> Validations = Parent.Parent.PoiSheet.GetDataValidations();
            //Validationの検索
            for (int v = 0; v < Validations.Count; v++)
            {
                IDataValidation vd = Validations[v];
                Utils.CellRangeAddressListOperator AdrOper
                    = new Utils.CellRangeAddressListOperator(vd.Regions, Remaining);
                //このRangeの残り部分がこのValidationに含まれる場合
                if(AdrOper.Overlapping.CountRanges() > 0)
                {
                    //所属Validation追記
                    this._BelongsTo.Add(vd);
                }
                //残りRange更新
                Remaining = AdrOper.TargetRemainder;
                //このRangeに所属Validationが見つかっていないRangeがない場合は検索終了
                if (Remaining.CountRanges() == 0)
                {
                    break;
                }
            }
            //このRangeのすべてが唯一のVlidatoinに属している場合
            if(Remaining.CountRanges() == 0 && this._BelongsTo.Count == 1)
            {
                BindingFlags Flag;
                FieldInfo FieldInfo;
                int PoiValidationType;
                int PoiOperatorType;
                bool PoiInCellDropdown;
                int PoiAlertStyle;
                //プロパティアクセスを可能にする。
                this._Accessable = true;
                //Validation情報
                IDataValidation Validation = this._BelongsTo[0];
                PoiOperatorType = Validation.ValidationConstraint.Operator;
                this._Operator = (int)XlFormatConditionOperatorParser.GetXlValue(PoiOperatorType);
                this._Formula1 = Validation.ValidationConstraint.Formula1;
                this._Formula2 = Validation.ValidationConstraint.Formula2;
                this._IgnoreBlank = Validation.EmptyCellAllowed;
                PoiInCellDropdown = InCellDropdownParser.GetXlValue(Parent.Parent, Validation.SuppressDropDownArrow);
                this._InCellDropdown = PoiInCellDropdown;
                this._ShowInput = Validation.ShowPromptBox;
                this._InputTitle = Validation.PromptBoxTitle;
                this._InputMessage = Validation.PromptBoxText;
                this._ShowError = Validation.ShowErrorBox;
                this._ErrorTitle = Validation.ErrorBoxTitle;
                this._ErrorMessage = Validation.ErrorBoxText;
                PoiAlertStyle = Validation.ErrorStyle;
                this._AlertStyle = (int)XlDVAlertStyleParser.GetXlValue(PoiAlertStyle);
                //IMEMpdeの取得
                if (Parent.Parent.IsHSSF)
                {
                    //Typeの取得
                    //IDataValidationの実体はHSSFDataValidation
                    //HSSFDataValidation.ValidationConstraint._validationTypeをreflectionで取得する。
                    //(public)ValidationConstraintの(private)_validationTypeをreflectionで取得
                    Flag = BindingFlags.NonPublic | BindingFlags.Instance;
                    FieldInfo = Validation.ValidationConstraint.GetType().GetField("_validationType", Flag);
                    PoiValidationType = (int)FieldInfo.GetValue(Validation.ValidationConstraint);
                    this._Type = (int)XlDVTypeParser.GetXlValue(PoiValidationType);
                    //InternalSheetの取得
                    NPOI.HSSF.Model.InternalSheet InternalSheet = ((HSSFSheet)Parent.Parent.PoiSheet).Sheet;


                    //InternalSheet._dataValidityTableの取得
                    Flag = BindingFlags.NonPublic | BindingFlags.Instance;
                    FieldInfo = InternalSheet.GetType().GetField("_dataValidityTable", Flag);
                    NPOI.HSSF.Record.Aggregates.DataValidityTable ValidationTable
                        = (NPOI.HSSF.Record.Aggregates.DataValidityTable)FieldInfo.GetValue(InternalSheet);
                    //InternalSheet._dataValidityTable._validationListの取得
                    FieldInfo ListInfo = ValidationTable.GetType().GetField("_validationList", Flag);
                    List<ArrayList> HssfValidation = (List<ArrayList>)ListInfo.GetValue(ValidationTable);
                }
                else
                {
                    //Typeの取得
                    //IDataValidationの実体はXSSFDataValidation
                    //XSSFDataValidation.ValidationConstraint.validationTypeをreflectionで取得する。
                    //(public)ValidationConstraintの(private)validationTypeをreflectionで取得
                    Flag = BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.GetField;
                    FieldInfo = Validation.ValidationConstraint.GetType().GetField("validationType", Flag);
                    PoiValidationType = (int)FieldInfo.GetValue(Validation.ValidationConstraint);
                    this._Type = (int)XlDVTypeParser.GetXlValue(PoiValidationType);
                    //IMEModeの取得
                    //XSSFDataValidation.ctDdataValidation.imeModeを取得する。
                    //(private)ctDdataValidationをreflectionで取得してimeModeにアクセス。
                    Flag = BindingFlags.NonPublic | BindingFlags.Instance;
                    FieldInfo TableInfo = Validation.GetType().GetField("ctDdataValidation", Flag);
                    CT_DataValidation XssfValidation = (CT_DataValidation)TableInfo.GetValue(Validation);
                    this._IMEMode= (int)XlIMEModeParser.GetXlValue(XssfValidation.imeMode);
                }
                //Type別のFormula処理
                switch ((XlDVType)this._Type)
                { 
                    //入力のみ
                    case XlDVType.xlValidateInputOnly:
                        //Formula1/Formula2の指定はないはずなので何もしない
                        break;
                    //受付時、xlValidateWholeNumberは小数点数を例外扱いとしているが、ここではxlValidateWholeNumber, xlValidateDecimalとも
                    //doubleにキャストできれば即値指定と判断している。
                    //またxlValidateTextLengthもFormulaの属性は同じなので共通処理する。
                    case XlDVType.xlValidateWholeNumber:
                    case XlDVType.xlValidateDecimal:
                    case XlDVType.xlValidateTextLength:
                        //即値でなければ式とみなす
                        if (!double.TryParse(this._Formula1, out _))
                        {
                            this._Formula1 = Formularize(this._Formula1);
                        }
                        if (!double.TryParse(this._Formula2, out _))
                        {
                            this._Formula2 = Formularize(this._Formula2);
                        }
                        break;
                    case XlDVType.xlValidateList:
                        //Formula1のみ処理
                        if (this._Formula1 != null)
                        {
                            //ダブルコーテーションで囲まれていれば即値リストと判断する。
                            if (this._Formula1.StartsWith("\"") && this._Formula1.StartsWith("\""))
                            {
                                //ダブルクォーテーションの囲みを除去する。
                                this._Formula1 = this._Formula1.TrimStart('"').TrimEnd('"');
                            }
                            //ダブルコーテーションで囲まれていなければ式と判断
                            else
                            {
                                this._Formula1 = Formularize(this._Formula1);
                            }
                        }
                        //Formula2はないはずはので何もしない
                        break;
                    case XlDVType.xlValidateDate:
                    case XlDVType.xlValidateTime:
                        string ParseFormat = (this._Type == (int)XlDVType.xlValidateDate ? "yyyy/MM/dd" : "HH:mm:ss");
                        string DateTimeString;
                        //日付にキャストできる場合は即値指定と判断して日付フォーマット文字列に変換
                        if (TryParseToDateTime(this._Formula1, ParseFormat, out DateTimeString))
                        {
                            this._Formula1 = DateTimeString;
                        }
                        //即値でなければ式とみなす
                        else
                        {
                            this._Formula1 = Formularize(this._Formula1);
                        }
                        if (TryParseToDateTime(this._Formula2, ParseFormat, out DateTimeString))
                        {
                            this._Formula2 = DateTimeString;
                        }
                        else
                        {
                            this._Formula2 = Formularize(this._Formula2);
                        }
                        break;
                    case XlDVType.xlValidateCustom:
                        //常に式とみなす
                        this._Formula1 = Formularize(this._Formula1);
                        this._Formula2 = Formularize(this._Formula2);
                        break;
                }
            }
            else
            {
                //プロパティアクセスを不能にする。
                _Accessable = false;
                //初期値のセット
                this._Type = (int)XlDVType.xlValidateInputOnly;
                this._Operator = 0; //XlFormatConditionOperatorにはない値だが初期値としてセット
                this._Formula1 = null;
                this._Formula2 = null;
                this._IgnoreBlank = true;
                this._InCellDropdown = true;
                this._ShowInput = false;
                this._InputTitle = string.Empty;
                this._InputMessage = string.Empty;
                this._ShowError = false;
                this._ErrorTitle = string.Empty;
                this._ErrorMessage = string.Empty;
                this._AlertStyle = (int)XlDVAlertStyle.xlValidAlertStop;
                this._IMEMode = (int)XlIMEMode.xlIMEModeNoControl;
            }
        }

        /// <summary>
        /// シートのValidationsにこのRangeを含む定義があれば、そのValidationからこのRangeを除外する。
        /// このRangeのみのValidationであればであればValidation定義そのものを削除する。
        /// </summary>
        private void UnlinkOrDeleteValidation()
        {
            if (Parent.Parent.IsHSSF)
            {

            }
            else
            {
                //XSSFのValidationsを取得しリストがあれば処理
                CT_DataValidations XSSFValidations = ((XSSFSheet)Parent.Parent.PoiSheet).GetCTWorksheet().dataValidations;
                if (XSSFValidations != null)
                {
                    //削除対象リストの確保
                    List<CT_DataValidation> DeleteList = new List<CT_DataValidation>();
                    //XSSFのValidationを取得
                    List<CT_DataValidation> XSSFValidation = XSSFValidations.dataValidation;
                    //Validationの検索
                    for (int v = 0; v < XSSFValidation.Count; v++)
                    {
                        CT_DataValidation vt = XSSFValidation[v];
                        //対象RangeをCellAddressRangeListに変換
                        string[] RangeList = vt.sqref.Split(' ');
                        CellRangeAddressList ValidatinoAddressList = new CellRangeAddressList();
                        foreach (string Range in RangeList)
                        {
                            ValidatinoAddressList.AddCellRangeAddress(CellRangeAddress.ValueOf(Range));
                        }
                        //アドレス評価
                        Utils.CellRangeAddressListOperator AdrOper
                            = new Utils.CellRangeAddressListOperator(ValidatinoAddressList, Parent.SafeAddressList);
                        //このRangeを支配するValidationの場合
                        if (AdrOper.Overlapping.CountRanges() > 0)
                        {
                            //このRangeのみを支配するValidationの場合はValidation自体を削除する
                            if (AdrOper.BaseRemainder.CountRanges() == 0)
                            {
                                //削除リストに追記
                                DeleteList.Add(vt);
                            }
                            //他にも支配するRangeがあれば、このRangeだけを支配から除外する
                            else
                            {
                                vt.sqref = AdrOper.BaseRemainderString;
                            }
                            break;
                        }
                    }
                    //不要となったValidationの削除
                    foreach (CT_DataValidation vt in DeleteList)
                    {
                        XSSFValidation.Remove(vt);
                    }

                }
            }
        }

        /// <summary>
        /// 指定された情報をプロパティに保存
        /// </summary>
        /// <param name="Type">XlDVTypeで指定する入力規則の種類</param>
        /// <param name="AlertStyle">XlDVAlertStyleで指定するエラーメッセージスタイル</param>
        /// <param name="Operator">XlFormatConditionOperatorで指定する値評価の種類</param>
        /// <param name="Formula1">最小値等の評価パラメータ</param>
        /// <param name="Formula2">最大値等の評価パラメータ</param>
        private void SetParamsToProperties(XlDVType Type, object AlertStyle = null, object Operator = null, object Formula1 = null, object Formula2 = null)
        {
            //Typeは型指定があるので信頼する。
            this.Type = (int)Type;
            //AlertStyleはXlDVAlertStyleまたはintでありXlDVAlertStyleに含まれる値であること
            if (AlertStyle is XlDVAlertStyle SafeAlertStyle)
            {
                this.AlertStyle = (int)SafeAlertStyle;
            }
            else if (AlertStyle is int SafeAlertStyleInt)
            {
                if (Enum.IsDefined(typeof(XlDVAlertStyle), SafeAlertStyleInt))
                {
                    this._AlertStyle = SafeAlertStyleInt;
                }
            }
            //OperatorはXlFormatConditionOperatorまたはintでありXlFormatConditionOperatorに含まれる値であること
            if (Operator is XlFormatConditionOperator SafeOperator)
            {
                this._Operator = (int)SafeOperator;
            }
            else if (Operator is int SafeOperatorInt)
            {
                if (Enum.IsDefined(typeof(XlFormatConditionOperator), SafeOperatorInt))
                {
                    this._Operator = SafeOperatorInt;
                }
            }
            //FromulaはNULLでなければ文字列として採用
            //(省略と区別がつかないので上位での明示的なnullクリアはできない。""クリアしてもらうしかない。)
            if (Formula1 != null)
            {
                this._Formula1 = Formula1.ToString();
            }
            if (Formula2 != null)
            {
                this._Formula2 = Formula2.ToString();
            }
            //Type別のFormula処理
            switch ((XlDVType)this._Type)
            {
                case XlDVType.xlValidateInputOnly:
                case XlDVType.xlValidateWholeNumber:
                case XlDVType.xlValidateDecimal:
                case XlDVType.xlValidateList:
                    break;
                case XlDVType.xlValidateDate:
                    string Param;
                    //即値指定の場合はOADateに変換する
                    if (this._Formula1 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOADate(this._Formula1, out Param))
                        {
                            this._Formula1 = Param;
                        }
                    }
                    if (this._Formula2 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOADate(this._Formula2, out Param))
                        {
                            this._Formula2 = Param;
                        }
                    }
                    break;
                case XlDVType.xlValidateTime://TryParseToOATime
                case XlDVType.xlValidateTextLength:
                case XlDVType.xlValidateCustom:
                    break;
            }
        }

        /// <summary>
        /// 入力規則の追加(既にあれば更新)
        /// POIの各ConstraintメソッドはFormulaの文字列が"="で始まれば式、"="でなければ即値と判断している。
        /// ただし即値を"="で指定しても式としては正しい。
        /// なので、ユーザから指定された値をそのまま渡せば良い。
        /// ただしxlValidateListで、ユーザがカンマ区切りの即値を指定する場合は、先頭が"="であってはならない。
        /// なので、xlValidateListで先頭文字が"="でなければ、カンマ分割し、CreateExplicitListConstraint()にstring[]を指定する。
        /// 【重要】
        /// このプロジェクトが使用しているNPOI2.6.0において、Formula内のCell/Range参照は同一シート内でなければならない。
        /// シートを超えて参照が可能なのは、Book範囲で名前付けされたCell/Rangeのみである。
        /// </summary>
        private void Apply()
        {
            //DataValidationHelper取得
            IDataValidationHelper Helper = Parent.Parent.PoiSheet.GetDataValidationHelper();
            //Constraint
            IDataValidationConstraint Cst;
            //XlパラメータのPoi化
            int PoiOperatorType = XlFormatConditionOperatorParser.GetPoiValue((XlFormatConditionOperator)this._Operator);
            ST_DataValidationImeMode PoiIMEMode = XlIMEModeParser.GetPoiValue((XlIMEMode)this._IMEMode);
            int PoiAlertStyle = XlDVAlertStyleParser.GetPoiValue((XlDVAlertStyle)this._AlertStyle);
            bool SuppressDropDownArrow = InCellDropdownParser.GetPoiValue(Parent.Parent, this._InCellDropdown);
            //XlDVType別処理(_Typeを読む。Typeでは追加初回に例外が発生するので。)
            switch ((XlDVType)this._Type)
            {
                //すべての値
                case XlDVType.xlValidateInputOnly:
                    //XSSFはHelper経由のCreateNumericConstraintでANY+IGNOREを指定できないのでコントラクタを直使いする。
                    if (Parent.Parent.IsXSSF)
                    {
                        Cst = new XSSFDataValidationConstraint(ValidationType.ANY, OperatorType.IGNORED, null, null);
                    }
                    else
                    {
                        //HSSFは素直に。
                        Cst = Helper.CreateNumericConstraint(ValidationType.ANY, OperatorType.IGNORED, null, null);
                    }
                    break;
                //整数
                case XlDVType.xlValidateWholeNumber:
                    Cst = Helper.CreateintConstraint(PoiOperatorType, this._Formula1, this._Formula2);
                    break;
                //小数点数
                case XlDVType.xlValidateDecimal:
                    Cst = Helper.CreateDecimalConstraint(PoiOperatorType, this._Formula1, this._Formula2);
                    break;
                //リスト
                case XlDVType.xlValidateList:
                    //"="で始まるならば式として処理
                    if (this._Formula1.StartsWith("="))
                    {
                        //Formula1をそのままCreateFormulaListConstraint()に渡す
                        Cst = Helper.CreateFormulaListConstraint(this._Formula1);
                    }
                    //"="でなければCSVの即値として処理
                    //Excelアプリ, Interop.Excelで即値を"="で始めると、参照先が見つからないとしてエラー/例外となる。
                    else
                    {
                        //カンマで配列に分離し、CreateExplicitListConstraint()に渡す。
                        string[] ExplicitList = this._Formula1.Split(',');
                        Cst = Helper.CreateExplicitListConstraint(ExplicitList);
                    }
                    break;
                //日付
                case XlDVType.xlValidateDate:
                    string DateParam;
                    //即値指定の場合はOADateに変換する
                    if (this._Formula1 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOADate(this._Formula1, out DateParam))
                        {
                            this._Formula1 = DateParam;
                        }
                    }
                    if (this._Formula2 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOADate(this._Formula2, out DateParam))
                        {
                            this._Formula2 = DateParam;
                        }
                    }
                    Cst = Helper.CreateDateConstraint(PoiOperatorType, this._Formula1, this._Formula2, null);
                    break;
                //時刻
                case XlDVType.xlValidateTime:
                    string TimeParam;
                    //即値指定の場合はOADateに変換する
                    if (this._Formula1 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOATime(this._Formula1, out TimeParam))
                        {
                            this._Formula1 = TimeParam;
                        }
                    }
                    if (this._Formula2 != null && !this._Formula1.StartsWith("="))
                    {
                        if (TryParseToOATime(this._Formula2, out TimeParam))
                        {
                            this._Formula2 = TimeParam;
                        }
                    }
                    Cst = Helper.CreateTimeConstraint(PoiOperatorType, this._Formula1, this._Formula2);
                    break;
                //文字列(長さ指定)
                case XlDVType.xlValidateTextLength:
                    Cst = Helper.CreateTextLengthConstraint(PoiOperatorType, this._Formula1, this._Formula2);
                    break;
                //ユーザー設定
                case XlDVType.xlValidateCustom:
                    Cst = Helper.CreateCustomConstraint(this._Formula1);
                    break;
                default:
                    throw new ArgumentException("Validation.Type");
            }
            //バリデーションの作成
            IDataValidation Val
                = Parent.Parent.PoiSheet.GetDataValidationHelper().CreateValidation(Cst, Parent.SafeAddressList);
            //空白セルの許可
            Val.EmptyCellAllowed = this._IgnoreBlank;
            //エラーメッセージ
            Val.ShowErrorBox = this._ShowError;
            Val.CreateErrorBox(this._ErrorTitle, this._ErrorMessage);
            Val.ErrorStyle = PoiAlertStyle;
            //入力メッセージ
            Val.ShowPromptBox = this._ShowInput;
            Val.CreatePromptBox(this._InputTitle, this._InputMessage);
            //ドロップダウンリスト
            Val.SuppressDropDownArrow = SuppressDropDownArrow;
            //IMEモード
            BindingFlags　Flag = BindingFlags.NonPublic | BindingFlags.Instance;
            FieldInfo TableInfo = Val.GetType().GetField("ctDdataValidation", Flag);
            CT_DataValidation XssfValidation = (CT_DataValidation)TableInfo.GetValue(Val);
            XssfValidation.imeMode = PoiIMEMode;
            //このRangeのバリデーションを除去
            UnlinkOrDeleteValidation();
            //バリデーションの登録
            Parent.Parent.PoiSheet.AddValidationData(Val);
        }

        /// <summary>
        /// アクセスフィルター
        /// </summary>
        /// <returns></returns>
        private void AccessFilter()
        {
            //アクセス可能でなければ例外
            if(!this._Accessable)
            {
                throw new InvalidOperationException("Can't identify the Validation information of this Range.");
            }
        }

        /// <summary>
        /// Excelの日付シリアル値(文字列)を指定されたフォーマットの日付時刻文字列に変換する。変換できない場合はstring.Emptyでリターンする。）
        /// </summary>
        /// <param name="OATimeString">Excelの日付シリアル値(文字列)</param>
        /// <param name="Format">フォーマット(yyyy/MM/dd, HH:mmなど)</param>
        /// <param name="DateTimeString">変換結果</param>
        /// <returns></returns>
        private bool TryParseToDateTime(string OATimeString, string Format, out string DateTimeString)
        {
            DateTimeString = null;
            bool RetVal = false;
            if (double.TryParse(OATimeString, out double Temp))
            {
                try
                {
                    //(OfficeAutomationのシリアル日付からDataTimeに変換して文字列化)
                    DateTimeString = DateTime.FromOADate(Temp).ToString(Format);
                    RetVal = true;
                }
                catch
                {
                    DateTimeString = null;
                    RetVal = false;
                }
            }
            return RetVal;
        }

        private bool TryParseToOADate(string DateString, out string OADate)
        {
            OADate = null;
            bool RetVal = false;
            if (DateTime.TryParse(DateString, out DateTime DateTimeValue))
            {
                OADate = DateTimeValue.ToOADate().ToString();
                RetVal = true;
            }
            return RetVal;
        }

        private bool TryParseToOATime(string TimeString, out string OATime)
        {
            OATime = null;
            bool RetVal = false;
            CultureInfo Culture = CultureInfo.CurrentCulture;
            DateTimeStyles Styles = DateTimeStyles.NoCurrentDateDefault;
            string[] TimeParts = TimeString.Split(':');
            if (DateTime.TryParseExact(TimeString, (TimeParts.Length >= 3 ? "H:m:s" : "H:m"), Culture, Styles, out DateTime DateTimeValue))
            {
                OATime = DateTimeValue.ToOADate().ToString();
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// 式として表現する(先頭に"="を付加するのみ)。null/長さ0の場合はnullを返す。
        /// </summary>
        /// <param name="RawString">対象文字列</param>
        /// <returns>式</returns>
        private string Formularize(string RawString)
        {
            string RetVal = null;
            //nullでなく、実際に文字が存在する場合は先頭に"="を付加する。
            if (RawString != null)
            {
                if (RawString.Length > 0)
                {
                    RetVal = "=" + RawString;
                }
            }
            return RetVal;
        }

        #endregion

        #endregion

        /// <summary>
        /// InCellDropdownとSuppressDropDownArrowの相互変換
        /// POIは表示抑止true、Excelは表示trueで反転した情報を管理している。
        /// 基本的にはboolの反転だが、XSSFはバグっているのか意味が反転しているので、さらに反転させる。
        /// この反転考慮はParseする瞬間にとどめ、それ以外では上位/下位ともそれぞれが期待する意味で処理できる。
        /// </summary>
        private static class InCellDropdownParser
        {
            /// <summary>
            /// InCellDropdown値を指定してSuppressDropDownArrow値を取得
            /// </summary>
            /// <param name="XlValue">InCellDropdown値</param>
            /// <returns>SuppressDropDownArrow値</returns>
            public static bool GetPoiValue(Worksheet Sheet, bool XlValue)
            {
                bool RetVal = !XlValue;
                if (Sheet.IsXSSF)
                {
                    RetVal = !RetVal;
                }
                return RetVal;
            }
            /// <summary>
            /// SuppressDropDownArrow値を指定してInCellDropdown値を取得
            /// </summary>
            /// <param name="PoiValue">SuppressDropDownArrow</param>
            /// <returns>InCellDropdown値</returns>
            public static bool GetXlValue(Worksheet Sheet, bool PoiValue)
            {
                bool RetVal = !PoiValue;
                if (Sheet.IsXSSF)
                {
                    RetVal = !RetVal;
                }
                return RetVal;
            }
        }
    }
}
