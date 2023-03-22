using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


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



    internal class Validation
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
        private bool _InCellDropdown = false;
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
        private int _AlertStyle  = (int)XlDVAlertStyle.xlValidAlertInformation;
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
        private string _Formula1 = string.Empty;
        /// <summary>
        /// Formula2の実体
        /// </summary>
        private string _Formula2 = string.Empty;
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

        internal Validation(Range ParentRange)
        {
            this.Parent = ParentRange;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

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
        public bool InCellDropdown
        {
            get
            { 
                return _InCellDropdown;
            }
            set
            {
                _InCellDropdown = value;
            }
        }
        public bool ShowInput
        {
            get
            { 
                return _ShowInput;
            }
            set
            {
                _ShowInput = value;
            }
        }
        public string InputTitle
        {
            get
            { 
                return _InputTitle;
            }
            set
            {
                _InputTitle = value;
            }
        }
        public string InputMessage
        {
            get
            { 
                return _InputMessage;
            }
            set
            { 
                _ErrorMessage = value;
            }
        }
        public bool ShowError
        {
            get
            { 
                return _ShowError;
            }
            set
            { 
                _ShowError = value;
            }
        }
        public string ErrorTitle
        {
            get
            { 
                return _ErrorTitle;
            }
            set
            { 
                _ErrorTitle = value;
            }
        }
        public string ErrorMessage
        {
            get
            {
                return _ErrorMessage;
            }
            set
            {
                _ErrorMessage = value;
            }
        }
        public int IMEMode
        {
            get
            {
                return _IMEMode;
            }
            set
            { 
                _IMEMode = value;
            }
        }
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
        public bool Value
        {
            get
            {
                AccessFilter();
                return _Value
;
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
            //指定されたパラメータの保存(AddのTypeは型指定なので信頼する)
            UpdateProperties(Type, AlertStyle, Operator, Formula1, Formula2);
            //Validationの適用
            Apply();
            //最新の状態に更新
            GetProperties();
        }

        /// <summary>
        /// Validationの削除
        /// </summary>
        public void Delete()
        {
            //このRangeからValidationを除去
            UnlinkOrDeleteValidation();
            //最新の状態に更新
            GetProperties();
        }

        /// <summary>
        /// 入力規則の更新(なければ追加)
        /// </summary>
        /// <param name="Type">XlDVTypeで指定する入力規則の種類</param>
        /// <param name="AlertStyle">XlDVAlertStyleで指定するエラーメッセージスタイル</param>
        /// <param name="Operator">XlFormatConditionOperatorで指定する値評価の種類</param>
        /// <param name="Formula1">最小値等の評価パラメータ</param>
        /// <param name="Formula2">最大値等の評価パラメータ</param>
        public void Modify(object Type = null, object AlertStyle = null, object Operator = null, object Formula1 = null, object Formula2 = null)
        {
            //ModifyのTypeは[optional]objectなので存在と正当性をチェックする
            int ValidationType = this._Type;
            if (Type is int SafeInt)
            {
                if (Enum.IsDefined(typeof(XlDVType), SafeInt))
                {
                    ValidationType = SafeInt;
                }
            }
            //指定されたパラメータの保存
            UpdateProperties((XlDVType)ValidationType, AlertStyle, Operator, Formula1, Formula2);
            //Validationの適用
            Apply();
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
                //プロパティアクセスを可能にする。
                _Accessable = true;
                //Propertiesのセット
            }
        }

        /// <summary>
        /// シートのValidationsにこのRangeを含む定義があれば、そのValidationからこのRangeを除外する。
        /// このRangeのみのValidationであればであればValidation定義そのものを削除する。
        /// </summary>
        private void UnlinkOrDeleteValidation()
        {
            if (Parent.Parent.PoiSheet is HSSFSheet)
            {

            }
            else
            {
                //XSSFのValidationを取得
                List<NPOI.OpenXmlFormats.Spreadsheet.CT_DataValidation> XSSFValidation
                    = ((XSSFSheet)Parent.Parent.PoiSheet).GetCTWorksheet().dataValidations.dataValidation;
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
        private void UpdateProperties(XlDVType Type, object AlertStyle = null, object Operator = null, object Formula1 = null, object Formula2 = null)
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
                    this.AlertStyle = SafeAlertStyleInt;
                }
            }
            //OperatorはXlFormatConditionOperatorまたはintでありXlFormatConditionOperatorに含まれる値であること
            if (Operator is XlFormatConditionOperator SafeOperator)
            {
                this.Operator = (int)SafeOperator;
            }
            else if (Operator is int SafeOperatorInt)
            {
                if (Enum.IsDefined(typeof(XlFormatConditionOperator), SafeOperatorInt))
                {
                    this.Operator = SafeOperatorInt;
                }
            }
            //FromulaはNULLでなければ文字列として採用
            //(省略と区別がつかないので上位での明示的なnullクリアはできない。""クリアしてもらうしかない。)
            if (Formula1 != null)
            {
                this.Formula1 = Formula1.ToString();
            }
            if (Formula2 != null)
            {
                this.Formula2 = Formula2.ToString();
            }
        }

        /// <summary>
        /// 入力規則の追加(既にあれば更新)
        /// </summary>
        private void Apply()
        {
            //DataValidationHelper取得
            IDataValidationHelper Helper = Parent.Parent.PoiSheet.GetDataValidationHelper();
            //Constraint
            IDataValidationConstraint Cst = null;
            //XlDVType別処理(_Typeを読む。Typeでは追加初回に例外が発生するので。)
            switch ((XlDVType)this._Type)
            {
                //すべての値
                case XlDVType.xlValidateInputOnly:
                    Cst = Helper.CreateCustomConstraint("TRUE");
                    break;
                //整数
                case XlDVType.xlValidateWholeNumber:
                    Cst = Helper.CreateintConstraint(this.Operator, this.Formula1, this.Formula2);
                    break;
                //小数点数
                case XlDVType.xlValidateDecimal:
                    Cst = Helper.CreateDecimalConstraint(this.Operator, this.Formula1, this.Formula2);
                    break;
                //リスト
                case XlDVType.xlValidateList:
                    //"="で始まるならば式として処理
                    if (this.Formula1.StartsWith("="))
                    {
                        //POIでは先頭に"="を付けない
                        this.Formula1 = this.Formula1.TrimStart('=');
                        Cst = Helper.CreateFormulaListConstraint(this.Formula1);
                    }
                    //"="でなければCSVの即値として処理
                    else
                    {
                        //カンマで配列に分離
                        string[] ExplicitList = this.Formula1.Split(',');
                        Cst = Helper.CreateExplicitListConstraint(ExplicitList);

                    }
                    break;
                //日付
                case XlDVType.xlValidateDate:
                    Cst = Helper.CreateDateConstraint(this.Operator, this.Formula1, this.Formula2, null);
                    break;
                //時刻
                case XlDVType.xlValidateTime:
                    Cst = Helper.CreateTimeConstraint(this.Operator, this.Formula1, this.Formula2);
                    break;
                //文字列(長さ指定)
                case XlDVType.xlValidateTextLength:
                    Cst = Helper.CreateTextLengthConstraint(this.Operator, this.Formula1, this.Formula2);
                    break;
                //ユーザー設定
                case XlDVType.xlValidateCustom:
                    Cst = Helper.CreateCustomConstraint(this.Formula1);
                    break;
                default:
                    Cst = Helper.CreateCustomConstraint("TRUE");
                    break;
            }
            //バリデーションの作成
            IDataValidation Val
                = Parent.Parent.PoiSheet.GetDataValidationHelper().CreateValidation(Cst, Parent.SafeAddressList);
            //プロパティ指定されている値の設定
            Val.ErrorStyle = ERRORSTYLE.STOP;
            Val.ShowErrorBox = this.ShowError;
            if (this.ShowInput)
            {
                Val.CreatePromptBox(this.InputTitle, this.InputMessage);
            }
            if (this.ShowError)
            {
                Val.CreateErrorBox(this.ErrorTitle, this.ErrorMessage);
            }
            Val.ShowPromptBox = this.ShowInput;
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

        #endregion

        #endregion
    }
}
