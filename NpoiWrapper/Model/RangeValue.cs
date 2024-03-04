using Developers.NpoiWrapper.Utils;
using MathNet.Numerics;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Developers.NpoiWrapper.Model
{
    internal class RangeValue
    {
        #region "constants"

        private const string BUILT_IN_FORMAT_DATE = "m/d/yy";
        private const string BUILT_IN_FORMAT_DATETIME = "m/d/yy h:mm";
        private const string BUILT_IN_FORMAT_TIME = "h:mm";

        #endregion


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
        /// <param name="ParentRange">Rangeインスタンス</param>
        public RangeValue(Range ParentRange)
        {
            //親Range情報の保存
            this.ParentRange = ParentRange;
        }

        #endregion

        #region "constructors"

        /// <summary>
        /// 処理対象プロパティ
        /// </summary>
        private enum TargetType
        {
            Value,
            Value2,
            Text,
            Formula
        }

        /// <summary>
        /// 時刻に含まれる情報
        /// </summary>
        [Flags]
        private enum DateTimePart
        {
            None = 0,
            Date = 1,
            Hour = 2,
            Minute = 4,
            Second = 8,
            Time = 6
        }

        #endregion

        #region "properties"

        /// <summary>
        /// レンジの値(書式は自動判断する)
        /// </summary>
        public object Value
        {
            get
            {
                return Getter(TargetType.Value);
            }
            set
            {
                Setter(TargetType.Value, value);
            }
        }

        /// <summary>
        /// レンジの値(書式なし)
        /// </summary>
        public object Value2
        {
            get
            {
                return Getter(TargetType.Value2);
            }
            set
            {
                Setter(TargetType.Value2, value);
            }
        }

        /// <summary>
        /// セルの文字列(ゲットのみ)
        /// </summary>
        public object Text
        {
            get
            {
                return Getter(TargetType.Text);
            }
        }

        /// <summary>
        /// セルの数式
        /// </summary>
        public object Formula
        {
            get
            {
                return Getter(TargetType.Formula);
            }
            set
            {
                Setter(TargetType.Formula, value);
            }
        }

        /// <summary>
        /// 親Range
        /// </summary>
        protected Range ParentRange { get; }

        /// <summary>
        /// 親IWorkbook
        /// </summary>
        private IWorkbook PoiBook { get { return this.ParentRange.Parent.Parent.PoiBook; } }

        /// <summary>
        /// 親ISheet
        /// </summary>
        private ISheet PoiSheet { get { return this.ParentRange.Parent.PoiSheet; } }

        /// <summary>
        /// 絶対表現(RonwIndex,ColumnIndexとして直接利用可能)されたアドレスリスト
        /// </summary>
        private CellRangeAddressList SafeAddressList { get { return this.ParentRange.SafeAddressList; } }

        #endregion

        #region "methods"

        /// <summary>
        /// Getterサブルーチン
        /// </summary>
        /// <param name="Target">読み取り対象</param>
        /// <returns>読み取り結果</returns>
        private object Getter(TargetType Target)
        {
            object RetVal = null;
            //Office.Interop.Excelにならい先頭アドレスのみ参照
            CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
            //値リストの確保
            object[,] Values = RangeUtil.CreateCellArray(
                SafeAddress.LastRow - SafeAddress.FirstRow + 1, SafeAddress.LastColumn - SafeAddress.FirstColumn + 1);
            int ValueMinRowIndex = Values.GetLowerBound(0);
            int ValueMinColIndex = Values.GetLowerBound(1);
            //行ループ
            for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
            {
                //列ループ
                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                {
                    object CelVal = null;
                    //読み取り対象別の処理
                    switch (Target)
                    {
                        //Valueを取得
                        case TargetType.Value:
                            CelVal = GetValue(PoiSheet, RIdx, CIdx);
                            break;
                        //Value2を取得
                        case TargetType.Value2:
                            CelVal = GetValue2(PoiSheet, RIdx, CIdx);
                            break;
                        //Textを取得
                        case TargetType.Text:
                            CelVal = GetText(PoiSheet, RIdx, CIdx);
                            break;
                        //Formulaを取得
                        case TargetType.Formula:
                            CelVal = GetFormula(PoiSheet, RIdx, CIdx);
                            break;
                        //あり得ないのでなにもしない
                        default:
                            break;
                    }
                    //取得した値を配列に保存(1開始のインデックス)
                    Values[
                        RIdx - SafeAddress.FirstRow + ValueMinRowIndex,
                        CIdx - SafeAddress.FirstColumn + ValueMinColIndex] = CelVal;
                }
            }
            //Textの読み取りならば値の集約を行う
            if(Target == TargetType.Text)
            {
                //配列を直線的リストに変換
                List<object> TextValues = new List<object>();
                foreach (object Text in Values)
                {
                    TextValues.Add(Text);
                }
                //値の集約
                IEnumerable<object> Distinct;
                Distinct = TextValues.Distinct();
                //値が１種類ならその値を採用
                if (Distinct.Count() == 1)
                {
                    RetVal = Distinct.First();
                }
                //複数種類あればNULL
                else
                {
                    RetVal = DBNull.Value;
                }
            }
            //Text以外ならば読み取り結果でリターン
            else
            {
                //基本的には読み取り結果配列でリターン
                RetVal = Values;
                //単一セルなら配列ではなく値そのものでリターン
                if (Values.Length == 1)
                {
                    RetVal = Values[ValueMinRowIndex, ValueMinColIndex];
                }
            }
            return RetVal;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Target">書込み対象</param>
        /// <param name="Value">書き込む値</param>
        private void Setter(TargetType Target, dynamic Value)
        {
            //Paste処理フラグ初期化
            bool PasteArray = false;
            int ValueMinRowIndex = 0;
            int ValueMinColIndex = 0;
            //実体のある値が指定されている場合
            if (Value != null)
            {
                //指定された値が配列の場合
                if (Value.GetType().IsArray)
                {
                    //２次元ならばRangeペースト処理を設定
                    if (((Array)Value).Rank == 2)
                    {
                        ValueMinRowIndex = ((Array)Value).GetLowerBound(0);
                        ValueMinColIndex = ((Array)Value).GetLowerBound(1);
                        PasteArray = true;
                    }
                }
            }
            //Office.Interop.Excelにならい非連続Rangeの全てに適用
            for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
            {
                //アドレス取得
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                //ペースト処理の場合はサイズの一致を確認
                if (PasteArray)
                {
                    int ValueRowCount = ((Array)Value).GetUpperBound(0) - ((Array)Value).GetLowerBound(0) + 1;
                    int ValueColCount = ((Array)Value).GetUpperBound(1) - ((Array)Value).GetLowerBound(1) + 1;
                    if ((SafeAddress.LastRow - SafeAddress.FirstRow + 1) != ValueRowCount
                        || (SafeAddress.LastColumn - SafeAddress.FirstColumn + 1) != ValueColCount)
                    {
                        //サイズ不一致ならば例外スロー
                        throw new ArgumentException("Specified array size diffrers from Range size.");
                    }
                }
                //行ループ
                for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                {
                    //列ループ
                    for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                    {
                        //セットする値の特定
                        object CValue = Value;
                        //Rangeペースト処理の場合は配列から値を取得
                        if (PasteArray)
                        {
                            CValue
                                = Value[
                                    RIdx - SafeAddress.FirstRow + ValueMinRowIndex,
                                    CIdx - SafeAddress.FirstColumn + ValueMinColIndex];
                        }
                        //書込み対象別の処理
                        switch (Target)
                        {
                            //Valueをセット
                            case TargetType.Value:
                                SetValue(PoiSheet, RIdx, CIdx, CValue);
                                break;
                            //Value2をセット
                            case TargetType.Value2:
                                SetValue(PoiSheet, RIdx, CIdx, CValue);
                                break;
                            //Formulaをセット
                            case TargetType.Formula:
                                SetFormula(PoiSheet, RIdx, CIdx, CValue);
                                break;
                            //あり得ないので何もしない
                            default:
                                break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// セルの値(Value)を取得する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private object GetValue(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            object RetVal = null;
            ////行の取得
            IRow row = Sheet.GetRow(RowIndex);
            //行があれば参照する
            if (row != null)
            {
                //セルの取得
                ICell cell = row.GetCell(ColumnIndex);
                //セルがあれば参照する
                if (cell != null)
                {
                    //セルの型に応じたプロパティを参照する
                    switch (cell.CellType)
                    {
                        //文字列 => string
                        case CellType.String:
                            RetVal = cell.StringCellValue;
                            break;
                        //数値
                        case CellType.Numeric:
                            //日付フォーマットされている場合 => DateTime
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                RetVal = cell.DateCellValue;
                            }
                            //日付フォーマットでない場合 => double
                            else
                                RetVal = cell.NumericCellValue;
                            break;
                        //Boolean => bool
                        case CellType.Boolean:
                            RetVal = cell.BooleanCellValue;
                            break;
                        //式(評価結果を返す)
                        case CellType.Formula:
                            //式を評価してCellValueを生成
                            IFormulaEvaluator Evaluator = PoiBook.GetCreationHelper().CreateFormulaEvaluator();
                            CellType EvaluatedType = Evaluator.EvaluateFormulaCell(cell);
                            //評価結果ごとに処理
                            switch (EvaluatedType)
                            {
                                //数値
                                case CellType.Numeric:
                                    //日付フォーマットされている場合 => DateTime
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        try
                                        {
                                            RetVal = DateTime.FromOADate(cell.NumericCellValue);
                                        }
                                        catch (Exception e)
                                        {
                                            Logger.Error("Formula evaluation failed." + e.ToString());
                                            RetVal = e.GetType().ToString() + "[" + e.Message + "]";
                                        }
                                    }
                                    //日付フォーマットでない場合 => double
                                    else
                                    {
                                        RetVal = cell.NumericCellValue;
                                    }
                                    break;
                                //文字列 => string
                                case CellType.String:
                                    RetVal = cell.StringCellValue;
                                    break;
                                //Boolean => bool
                                case CellType.Boolean:
                                    RetVal = cell.BooleanCellValue;
                                    break;
                                //エラー => string
                                case CellType.Error:
                                    try
                                    {
                                        RetVal = cell.NumericCellValue;
                                    }
                                    catch (Exception ex)
                                    {
                                        RetVal = ex.HResult;
                                    }
                                    break;
                                //その他(あり得ない模様)
                                default:
                                    RetVal = null;
                                    break;
                            }
                            break;
                        //エラー
                        case CellType.Error:
                            try
                            {
                                RetVal = cell.NumericCellValue;
                            }
                            catch (Exception ex)
                            {
                                RetVal = ex.HResult;
                            }
                            break;
                        //空白
                        case CellType.Blank:
                            RetVal = null;
                            break;
                        //その他
                        default:
                            RetVal = null;
                            break;
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// セルの値(Value)を設定する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <param name="Value">セットする値</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private bool SetValue(ISheet Sheet, int RowIndex, int ColumnIndex, object Value)
        {
            bool RetVal = false;

            //Create初期値false
            bool Create = false;
            //nullでなければCellをCreateする
            if(Value != null)
            {
                Create = true;
                //だだし空文字ならCreateしない
                if (Value is string StringValue)
                {
                    if (StringValue.Length == 0)
                    {
                        Create = false;
                    }
                }
            }
            //CellのCreateが必要なケース
            if (Create)
            {
                //行、列が存在しない場合は生成する。
                IRow row = Sheet.GetRow(RowIndex);
                if (row == null)
                {
                    row = Sheet.CreateRow(RowIndex);
                    RetVal = true;
                    Logger.Debug(
                        "Shhet[" + Sheet.SheetName + "]:Row[" + RowIndex + "] *** Row Created. ***");
                }
                ICell cell = row.GetCell(ColumnIndex);
                if (cell == null)
                {
                    cell = row.CreateCell(ColumnIndex);
                    Logger.Debug(
                        "Shhet[" + Sheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex + "] *** Column Created. ***");
                    RetVal = true;
                }
                //DateTimeの場合
                if (Value is DateTime SafeDateTime)
                {
                    cell.SetCellValue(SafeDateTime);
                    cell.SetCellType(CellType.Numeric);
                    //自動書式設定
                    string FormatString = "yyyy/mm/dd hh:mm";
                    //時分秒がすべて０なら日付のみとみなす
                    if (SafeDateTime.Hour == 0 && SafeDateTime.Minute == 0 && SafeDateTime.Second == 0)
                    {
                        FormatString = "yyyy/mm/dd";
                    }
                    //デフォルトスタイルの場合はFormat設定(ここでCreateしたCell、またはスタイル未設定のCell)
                    if (cell.CellStyle.Index == 0)
                    {
                        ParentRange.Parent.Range[cell.Address.FormatAsString()].NumberFormat = FormatString;
                        Logger.Debug(
                            "Shhet[" + Sheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex + "] Auto format:[" + FormatString + "]");
                    }
                }
                //boolの場合
                else if (Value is bool SafeBool)
                {
                    cell.SetCellValue(SafeBool);
                    cell.SetCellType(CellType.Boolean);
                }
                //stringの場合
                else if (Value is string StringValue)
                {
                    //セル書式が文字列の場合はそのまま文字列としてセット
                    if (cell.CellStyle.GetDataFormatString() == "@"
                        || cell.CellStyle.GetDataFormatString().ToLower() == "text")
                    {
                        cell.SetCellValue(StringValue);
                        cell.SetCellType(CellType.String);
                    }
                    //セル書式が文字列でなければ型判定
                    else
                    {
                        //DateTimeにキャストできる場合はDateTime
                        if (DateTime.TryParse(StringValue, out DateTime DateTimeValue))
                        {
                            //日付フォーマット
                            DateTimePart Format = DateTimePart.None;
                            //":"で分割してみる
                            string[] Parts = StringValue.Split(':');
                            //２分割以上なら時刻を含むとみなす。
                            if (Parts.Length >= 2)
                            {
                                Format |= DateTimePart.Time;
                                //分割の先頭が長さ２を超える場合は日付が含まれるとみなす。
                                if (Parts[0].Length > 2)
                                {
                                    Format |= DateTimePart.Date;
                                }
                                //３分割以上なら秒が含まれるとみなす
                                if (Parts.Length >= 3)
                                {
                                    Format |= DateTimePart.Second;
                                }
                            }
                            //分割できない場合は日付とみなす
                            else
                            {
                                Format = DateTimePart.Date;
                            }
                            //日付を含む場合はそのままセット
                            if (Format.HasFlag(DateTimePart.Date))
                            {
                                cell.SetCellValue(DateTimeValue);
                                cell.SetCellType(CellType.Numeric);
                            }
                            //日付を含まない場合は時刻のみでキャストしてセット
                            else
                            {
                                CultureInfo Culture = CultureInfo.CurrentCulture;
                                DateTimeStyles Styles = DateTimeStyles.NoCurrentDateDefault;
                                //時刻のみでDateTimeにキャストできれば時刻のみのデータとしてセット
                                if (DateTime.TryParseExact(
                                        StringValue, (Format.HasFlag(DateTimePart.Second) ? "H:m:s" : "H:m"),
                                        Culture, Styles, out DateTime TimeValue))
                                {
                                    //日付なしDateTimeのままだとなぜかdpuble(-1)になってしまうのでOADateをセットする
                                    cell.SetCellValue(TimeValue.ToOADate());
                                    cell.SetCellType(CellType.Numeric);
                                }
                                //キャストできなければ日付時刻としてセット
                                else
                                {
                                    cell.SetCellValue(DateTimeValue);
                                    cell.SetCellType(CellType.Numeric);
                                    //フォーマット再設定
                                    Format = DateTimePart.Date | DateTimePart.Hour | DateTimePart.Minute | DateTimePart.Second;
                                }
                            }
                            //自動書式設定
                            string FormatString = string.Empty;
                            if (Format.HasFlag(DateTimePart.Date))
                            {
                                FormatString = "yyyy/mm/dd";
                            }
                            if (Format.HasFlag(DateTimePart.Time))
                            {
                                FormatString += FormatString.Length > 0 ? " " : "";
                                FormatString += "hh:mm";
                                if (Format.HasFlag(DateTimePart.Second))
                                {
                                    FormatString += ":ss";
                                }
                            }
                            //デフォルトスタイルの場合はFormat設定(ここでCreateしたCell、またはスタイル未設定のCell)
                            if (cell.CellStyle.Index == 0)
                            {
                                ParentRange.Parent.Range[cell.Address.FormatAsString()].NumberFormat = FormatString;
                                Logger.Debug(
                                    "Shhet[" + Sheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex + "] Auto format:[" + FormatString + "]");
                            }
                        }
                        //数値にキャストできる場合はdouble
                        else if (double.TryParse(StringValue, out double DoubleValue))
                        {
                            cell.SetCellValue(DoubleValue);
                            cell.SetCellType(CellType.Numeric);
                        }
                        //日付でも数値でもなければstring
                        else
                        {
                            cell.SetCellValue(StringValue);
                            cell.SetCellType(CellType.String);
                        }
                    }
                }
                //数値系の型の場合(char含む)
                else
                {
                    //文字列に変換
                    string StrValue = Value.ToString();
                    //数値にキャストできる場合はdouble
                    if (double.TryParse(StrValue, out double DoubleValue))
                    {
                        cell.SetCellValue(DoubleValue);
                        cell.SetCellType(CellType.Numeric);
                    }
                    //数値でなければstring
                    else
                    {
                        cell.SetCellValue(StrValue);
                        cell.SetCellType(CellType.String);
                    }
                }

            }
            //nullをセットする場合
            else
            {
                //セルが実在する場合のみBlankCellを設定する
                IRow row = Sheet.GetRow(RowIndex);
                if (row != null)
                {
                    ICell cell = row.GetCell(ColumnIndex);
                    if (cell != null)
                    {
                        cell.SetCellValue((string)null);
                        cell.SetCellType(CellType.Blank);
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// セルの値(Value2)を取得する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private object GetValue2(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            object RetVal = null;
            ////行の取得
            IRow row = Sheet.GetRow(RowIndex);
            //行があれば参照する
            if (row != null)
            {
                //セルの取得
                ICell cell = row.GetCell(ColumnIndex);
                //セルがあれば参照する
                if (cell != null)
                {
                    //セルの型に応じたプロパティを参照する
                    switch (cell.CellType)
                    {
                        //文字列 => string
                        case CellType.String:
                            RetVal = cell.StringCellValue;
                            break;
                        //数値
                        case CellType.Numeric:
                            RetVal = cell.NumericCellValue;
                            break;
                        //Boolean => bool
                        case CellType.Boolean:
                            RetVal = cell.BooleanCellValue;
                            break;
                        //式(評価結果を返す)
                        case CellType.Formula:
                            //式を評価してCellValueを生成
                            IFormulaEvaluator Evaluator = PoiBook.GetCreationHelper().CreateFormulaEvaluator();
                            CellType EvaluatedType = Evaluator.EvaluateFormulaCell(cell);
                            //評価結果ごとに処理
                            switch (EvaluatedType)
                            {
                                //数値
                                case CellType.Numeric:
                                    RetVal = cell.NumericCellValue;
                                    break;
                                //文字列 => string
                                case CellType.String:
                                    RetVal = cell.StringCellValue;
                                    break;
                                //Boolean => bool
                                case CellType.Boolean:
                                    RetVal = cell.BooleanCellValue;
                                    break;
                                //エラー => string
                                case CellType.Error:
                                    try
                                    {
                                        RetVal = cell.NumericCellValue;
                                    }
                                    catch (Exception ex)
                                    {
                                        RetVal = ex.HResult;
                                    }
                                    break;
                                //その他(あり得ない模様)
                                default:
                                    RetVal = null;
                                    break;
                            }
                            break;
                        //エラー
                        case CellType.Error:
                            try
                            {
                                RetVal = cell.NumericCellValue;
                            }
                            catch (Exception ex)
                            {
                                RetVal = ex.HResult;
                            }
                            break;
                        //空白
                        case CellType.Blank:
                            RetVal = null;
                            break;
                        //その他
                        default:
                            RetVal = null;
                            break;
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// セルの値(Value2)を設定する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <param name="Value">セットする値</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private bool SetValue2(ISheet Sheet, int RowIndex, int ColumnIndex, object Value)
        {
            //Valueと同じ処理をしておく
            return SetValue(Sheet, RowIndex, ColumnIndex, Value);
        }

        /// <summary>
        /// セルの値(Text)を取得する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private object GetText(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            object RetVal = string.Empty;
            ////行の取得
            IRow row = Sheet.GetRow(RowIndex);
            //行があれば参照する
            if (row != null)
            {
                //セルの取得
                ICell cell = row.GetCell(ColumnIndex);
                //セルがあれば参照する
                if (cell != null)
                {
                    //フォーマッター生成
                    DataFormatter Formatter = new DataFormatter();
                    //Fomulaセルの場合
                    if (cell.CellType == CellType.Formula)
                    {
                        //数式式を評価
                        IFormulaEvaluator evaluator = PoiBook.GetCreationHelper().CreateFormulaEvaluator();
                        CellValue cellValue = evaluator.Evaluate(cell);
                        //評価結果が数値の場合
                        if (cellValue.CellType == CellType.Numeric)
                        {
                            //FormatRawCellContentsでフォーマット
                            RetVal = Formatter.FormatRawCellContents(
                                        cellValue.NumberValue, cell.CellStyle.DataFormat, cell.CellStyle.GetDataFormatString());
                        }
                        else if (cellValue.CellType == CellType.String)
                        {
                            RetVal = cellValue.StringValue;
                        }
                        else
                        {
                            RetVal = cellValue.FormatAsString();
                        }
                    }
                    //Formula以外はともかくFormatCellValueでフォーマットする
                    else
                    {
                        RetVal = Formatter.FormatCellValue(cell);
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// セルのFormulaを取得する(Formula設定のないセルでは、Value系を文字列化した値を取得する)
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <returns></returns>
        private object GetFormula(ISheet Sheet, int RowIndex, int ColumnIndex)
        {
            object RetVal = string.Empty;
            ////行の取得
            IRow row = Sheet.GetRow(RowIndex);
            //行があれば参照する
            if (row != null)
            {
                //セルの取得
                ICell cell = row.GetCell(ColumnIndex);
                //セルがあれば参照する
                if (cell != null)
                {
                    //Frmula設定されているセルならFormulaを取得
                    if (cell.CellType == CellType.Formula)
                    {
                        RetVal = "=" + cell.CellFormula;
                    }
                    //文字列セル
                    else if (cell.CellType == CellType.String)
                    {
                        //文字があればセットする
                        if (cell.StringCellValue != null)
                        {
                            RetVal = cell.StringCellValue;
                        }
                    }
                    //数値セル
                    else if (cell.CellType == CellType.Numeric)
                    {
                        RetVal = cell.NumericCellValue.ToString();
                    }
                    //Boleanセル
                    else if (cell.CellType == CellType.Boolean)
                    {
                        //RetVal = cell.BooleanCellValue.ToString();
                        //フォーマッター生成
                        DataFormatter Formatter = new DataFormatter();
                        RetVal = Formatter.FormatCellValue(cell);

                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// セルの値を設定する
        /// </summary>
        /// <param name="Sheet">ISheetインスタンス</param>
        /// <param name="RowIndex">行インデックス</param>
        /// <param name="ColumnIndex">列インデックス</param>
        /// <param name="Value">セットする式</param>
        /// <returns>(bool)行または列を新規作成した</returns>
        private bool SetFormula(ISheet Sheet, int RowIndex, int ColumnIndex, object Value)
        {
            bool RetVal = false;
            IFormulaEvaluator Evaluator = PoiBook.GetCreationHelper().CreateFormulaEvaluator();
            //文字列化
            string Formula = string.Empty;
            if (Value != null)
            {
                //objectからstringに変換し、先頭の"="を除去
                Formula = Value.ToString().TrimStart('=');
            }
            //Formulaが実在する場合
            if (Formula.Length > 0)
            {
                //行、列が存在しない場合は生成する。
                IRow row = Sheet.GetRow(RowIndex);
                if (row == null)
                {
                    row = Sheet.CreateRow(RowIndex);
                    RetVal = true;
                    Logger.Debug(
                        "Sheet[" + Sheet.SheetName + "]:Row[" + RowIndex + "] *** Row Created. ***");
                }
                ICell cell = row.GetCell(ColumnIndex);
                if (cell == null)
                {
                    cell = row.CreateCell(ColumnIndex);
                    RetVal = true;
                    Logger.Debug(
                        "Sheet[" + Sheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex  + "] *** Column Created. ***");
                }
                //Formulaのセット
                cell.SetCellType(CellType.Formula);
                cell.SetCellFormula(Formula);
                //数式評価
                CellType　EvaluatedType = Evaluator.EvaluateFormulaCell(cell);
                //評価結果が数値の場合
                if (EvaluatedType == CellType.Numeric)
                {
                    //日付フォーマットされていない場合
                    if (!DateUtil.IsCellDateFormatted(cell))
                    {
                        //日付とみなせる場合
                        if (DateUtil.IsValidExcelDate(cell.NumericCellValue))
                        {
                            //デフォルトスタイルの場合
                            if (cell.CellStyle.Index == 0)
                            {
                                //TODAY(), NOW()から始まる式であれば書式設定を行う(この２つのみ救済するが他は数値のまま)
                                if (Formula.ToLower().StartsWith("today"))
                                {
                                    ParentRange.Parent.Range[cell.Address.FormatAsString()].NumberFormat = "yyyy/mm/dd";
                                }
                                else if (Formula.ToLower().StartsWith("now"))
                                {
                                    ParentRange.Parent.Range[cell.Address.FormatAsString()].NumberFormat = "yyyy/mm/dd hh:mm";
                                }
                                else
                                {
                                    //何もしない
                                }
                            }
                        }
                    }
                }
            }
            //nullまたは空文字の場合
            else
            {
                //セルが実在する場合のみFormulaを設定する
                IRow row = Sheet.GetRow(RowIndex);
                if (row != null)
                {
                    ICell cell = row.GetCell(ColumnIndex);
                    if (cell != null)
                    {
                        //現在Formulaセルの場合はBlankに設定
                        if (cell.CellType == CellType.Formula)
                        {
                            cell.SetCellFormula(null);
                            cell.SetCellType(CellType.Blank);
                        }
                        //現在FormulaセルでなければFormulaのクリアのみ実施
                        else
                        {
                            cell.SetCellFormula(null);
                        }
                        //数式評価
                        Evaluator.EvaluateFormulaCell(cell);
                    }
                }
            }
            return RetVal; 
        }

        /// <summary>
        /// CellValueが日付フォーマットかどうか判定
        /// </summary>
        /// <param name="cell">元のCell</param>
        /// <param name="CellValue">Formula評価したCellValue</param>
        /// <returns></returns>
        private bool IsCellValueDateFormatted(ICell Cell, CellValue Value)
        {
            bool RetVal = false;
            if (Cell != null)
            {
                if (DateUtil.IsValidExcelDate(Value.NumberValue))
                {
                    ICellStyle style = Cell.CellStyle;
                    if (style != null)
                    {
                        int i = style.DataFormat;
                        String f = style.GetDataFormatString();
                        RetVal = DateUtil.IsADateFormat(i, f);
                    }
                }
            }
            return RetVal;
        }

        #endregion
    }
}
