﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.Linq;
using System;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Developers.NpoiWrapper.Util;
using System.Runtime.InteropServices.WindowsRuntime;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using System.Xml.Schema;

namespace Developers.NpoiWrapper
{

    public enum XlCellType
    {
        xlCellTypeLastCell
    }
    /// <summary>
    /// Rangeクラス
    /// WorksheetクラスのCells, Rangeプロパティにアクセスすると本クラスのインデクサでコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Range
    {
        internal Worksheet ParentSheet { get; private set; }
        internal CellRangeAddressList RawAddressList { get; private set; }
        internal CellRangeAddressList SafeAddressList { get; private set; }

        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        /// <param name="RangeAddressList">CellRangeAddressListインスタンス</param>
        internal Range(Worksheet ParentSheet, CellRangeAddressList RangeAddressList)
        {
            this.ParentSheet = ParentSheet;
            RawAddressList = RangeAddressList;
            SafeAddressList = GetSafeCellRangeAddressList(RawAddressList);
        }

        /// <summary>
        /// インデクサー
        /// </summary>
        /// <param name="Cell1"></param>
        /// <param name="Cell2"></param>
        /// <returns></returns>
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
                    AddressList = Uitl.GetMergedAddressList(AddressList);
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
                        AddressList = Uitl.GetMergedAddressList(AddressList);
                    }
                }
                //Cellsでもstringでもなければ例外スロー
                else
                {
                    throw new ArgumentException("Type of Cell1 must be Cells or string.");
                }
                //Rangeクラスインスタンス生成
                return new Range(ParentSheet, AddressList);
            }
        }

        /// <summary>
        /// Count
        /// このRageに含まれるCellの数
        /// </summary>
        public int Count
        {
            get
            {
                int RetVal = 0;
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx);
                    RetVal += RawAddress.NumberOfCells;

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
                    RetVal += RawAddress.FormatAsString() + ",";

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
                return new Areas(ParentSheet, RawAddressList);
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
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    //生アドレス取得
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx);
                    //列を全域に拡張しリストに追加
                    RangeAddressList.AddCellRangeAddress(RawAddress.FirstRow, RawAddress.LastRow, -1, -1);
                }
                //Rangeクラスインスタンス生成
                return new Range(ParentSheet, RangeAddressList);
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
                CellRangeAddressList RangeAddressList = new CellRangeAddressList();
                for (int AIdx = 0; AIdx < RawAddressList.CountRanges(); AIdx++)
                {
                    //生アドレス取得
                    CellRangeAddress RawAddress = RawAddressList.GetCellRangeAddress(AIdx);
                    //行を全域に拡張しリストに追加
                    RangeAddressList.AddCellRangeAddress(-1, -1, RawAddress.FirstColumn, RawAddress.LastColumn);
                    //Rangeクラスインスタンス生成
                }
                return new Range(ParentSheet, RangeAddressList);
            }
        }

        /// <summary>
        /// レンジの値(書式は自動判断する)
        /// </summary>
        public dynamic Value
        {
            get
            {
                //Office.Interop.Excelにならい先頭アドレスのみ参照
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(0);
                //値リストの確保
                object[,] Values = CreateCellArray(
                    SafeAddress.LastRow - SafeAddress.FirstRow + 1, SafeAddress.LastColumn - SafeAddress.FirstColumn + 1);
                //行ループ
                for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                {
                    //行の取得(なければ生成)
                    IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
                    //列ループ
                    for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                    {
                        //列の取得(なければ生成)
                        ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
                        object CelVal;
                        //セルの型に応じたプロパティを参照する
                        switch (cell.CellType)
                        {
                            //文字列
                            case CellType.String:
                                CelVal = cell.StringCellValue;
                                break;
                            //数値
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                    CelVal = cell.DateCellValue.ToString();
                                else
                                    CelVal = cell.NumericCellValue.ToString();
                                break;
                            //Boolean
                            case CellType.Boolean:
                                CelVal = cell.BooleanCellValue.ToString();
                                break;
                            //式(評価結果を返す)
                            case CellType.Formula:
                                IFormulaEvaluator evaluator
                                    = ParentSheet.ParentBook.PoiBook.GetCreationHelper().CreateFormulaEvaluator();
                                CellValue cellValue = evaluator.Evaluate(cell);
                                if (cellValue.CellType == CellType.String)
                                    CelVal = cellValue.StringValue;
                                else
                                    CelVal = cellValue.NumberValue.ToString();
                                break;
                            //エラー
                            case CellType.Error:
                                CelVal = cell.ErrorCellValue.ToString();
                                break;
                            //空白
                            case CellType.Blank:
                                CelVal = string.Empty;
                                break;
                            //その他
                            default:
                                CelVal = string.Empty;
                                break;
                        }
                        Values[
                            RIdx - SafeAddress.FirstRow + 1,
                            CIdx - SafeAddress.FirstColumn + 1] = CelVal;
                    }
                }
                //単一セルなら配列ではなく値そのものでリターン
                if (Values.Length == 1)
                {
                    return Values[1, 1];
                }
                return Values;
            }
            set
            {
                bool PasteArray = false;
                int ValueFirstRow = 0;
                int ValueFirstColumn = 0;
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //供給された値が配列の場合
                    if (value.GetType().IsArray)
                    {
                        //２次元ならばRangeペースト処理を設定
                        if (((Array)value).Rank == 2)
                        {
                            ValueFirstRow = ((Array)value).GetLowerBound(0);
                            ValueFirstColumn = ((Array)value).GetLowerBound(1);
                            PasteArray = true;
                        }
                    }
                    //行ループ
                    for (int RIdx = 0; RIdx <= SafeAddress.LastRow - SafeAddress.FirstRow; RIdx++)
                    {
                        //行の取得(なければ生成)
                        IRow row = ParentSheet.PoiSheet.GetRow(RIdx + SafeAddress.FirstRow)
                                    ?? ParentSheet.PoiSheet.CreateRow(RIdx + SafeAddress.FirstRow);
                        //列ループ
                        for (int CIdx = 0; CIdx <= SafeAddress.LastColumn - SafeAddress.FirstColumn; CIdx++)
                        {
                            //列の取得(なければ生成)
                            ICell cell = row.GetCell(CIdx + SafeAddress.FirstColumn)
                                            ?? row.CreateCell(CIdx + SafeAddress.FirstColumn);
                            //セットする値の特定
                            object CValue = value;
                            Cell.ValueType CType = Cell.ValueType.Auto;
                            //Rangeペースト処理の場合は配列から値を取得
                            if (PasteArray)
                            {
                                CValue = value[RIdx + ValueFirstRow, CIdx + ValueFirstColumn];
                                //配列要素がCellクラスなら解読
                                if (CValue is Cell)
                                {
                                    Cell c = value[RIdx + ValueFirstRow, CIdx + ValueFirstColumn];
                                    CValue = c.Value;
                                    CType = c.Type;
                                }
                            }
                            //文字列固定の場合
                            if (CType == Cell.ValueType.String)
                            {
                                cell.SetCellValue((string)CValue);
                                cell.SetCellType(CellType.String);
                            }
                            //式固定の場合
                            else if (CType == Cell.ValueType.Formula)
                            {
                                cell.SetCellFormula((string)CValue);
                                cell.SetCellType(CellType.Formula);
                            }
                            else
                            {
                                //日付であっても数値としてセット(ユーザによる書式設定を期待する)
                                if (DateTime.TryParse(CValue.ToString(), out DateTime dtm))
                                {
                                    cell.SetCellValue((DateTime)dtm);
                                    cell.SetCellType(CellType.Numeric);
                                }
                                //数値であれば数値としてセット
                                else if (double.TryParse(CValue.ToString(), out double dbl))
                                {
                                    cell.SetCellValue((double)dbl);
                                    cell.SetCellType(CellType.Numeric);
                                }
                                //その他は文字列扱い
                                else
                                {
                                    cell.SetCellValue((string)CValue);
                                    cell.SetCellType(CellType.String);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// セルの文字列(セットのみ)
        /// </summary>
        public string Text
        {
            ///セルの値設定
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
                        IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
                        //列ループ
                        for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                        {
                            //列の取得(なければ生成)
                            ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
                            cell.SetCellValue((string)value);
                            cell.SetCellType(CellType.String);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// セルの式(セットのみ)
        /// </summary>
        public string Formula
        {
            ///セルの値設定
            set
            {
                string Formula = value;
                Formula = Formula.TrimStart('=');
                //Office.Interop.Excelにならい非連続Rangeの全てに適用
                for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
                {
                    //アドレス取得
                    CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                    //行ループ
                    for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                    {
                        //行の取得(なければ生成)
                        IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
                        //列ループ
                        for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                        {
                            //列の取得(なければ生成)
                            ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
                            cell.SetCellFormula(Formula);
                            cell.SetCellType(CellType.Formula);
                        }
                    }
                }
            }
        }

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
                    IRow row = ParentSheet.PoiSheet.GetRow(RIdx);
                    if (row != null)
                    {
                        RetVal += row.HeightInPoints;
                    }
                    else
                    {
                        //twipなので20倍してpointに編案
                        RetVal += (ParentSheet.PoiSheet.DefaultRowHeight * 20);
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
                    IRow row = ParentSheet.PoiSheet.GetRow(RIdx);
                    if (row != null)
                    {
                        ht.Add(row.HeightInPoints);
                    }
                    else
                    {
                        //twipなので20倍してpointに編案
                        ht.Add(ParentSheet.PoiSheet.DefaultRowHeight * 20);
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
                        IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
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
                    RetVal += ParentSheet.PoiSheet.GetColumnWidth(CIdx);
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
                    wd.Add(ParentSheet.PoiSheet.GetColumnWidth(CIdx));
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
                        ParentSheet.PoiSheet.SetColumnWidth(CIdx, (int)value);
                    }
                }
            }
        }

        /// <summary>
        /// セルのコメントを生成する
        /// </summary>
        /// <param name="CommentText">コメント文字列</param>
        public void AddComment(string CommentText)
        {
            //Office.Interop.Excelにならい非連続Rangeの全てに適用
            for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
            {
                //先頭アドレス取得
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                //行ループ
                for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                {
                    //行の取得(なければ生成)
                    IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
                    //列ループ
                    for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                    {
                        //列の取得(なければ生成)
                        ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
                        IDrawing drawing = ParentSheet.PoiSheet.CreateDrawingPatriarch();
                        IClientAnchor anchor = ParentSheet.ParentBook.PoiBook.GetCreationHelper().CreateClientAnchor();
                        //サイズは固定で４×３を指定
                        anchor.Col1 = cell.ColumnIndex;
                        anchor.Col2 = cell.ColumnIndex + 4;
                        anchor.Row1 = cell.RowIndex;
                        anchor.Row2 = cell.RowIndex + 3;
                        IComment comment = drawing.CreateCellComment(anchor);
                        if (ParentSheet.PoiSheet is HSSFSheet)
                        {
                            comment.String = new HSSFRichTextString(CommentText);
                        }
                        else
                        {
                            comment.String = new XSSFRichTextString(CommentText);
                        }
                        cell.CellComment = comment;
                    }
                }

            }
        }

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
                = ParentSheet.PoiSheet.GetDataValidationHelper().CreateExplicitListConstraint(ExplicitList);
            //Office.Interop.Excelにならい非連続Rangeの全てに適用(RawAddressList)
            IDataValidation Val
                = ParentSheet.PoiSheet.GetDataValidationHelper().CreateValidation(Cst, RawAddressList);
            //なぜかHSSFとXSFでは指定が逆になっている模様
            if (ParentSheet.PoiSheet is HSSFSheet)
            {
                Val.SuppressDropDownArrow = false;
            }
            else
            {
                //なぜかサプレスTRUEにすると表示される
                Val.SuppressDropDownArrow = true;
            }
            Val.ErrorStyle = 0;
            Val.ShowErrorBox = ShowErrorBox;
            Val.CreateErrorBox(ErrorBoxTitle, ErrorBoxText);
            Val.ShowPromptBox = ShowPronptBox;
            Val.CreatePromptBox(PronptBoxTitle, PronptBoxText);
            ParentSheet.PoiSheet.AddValidationData(Val);
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
                if (ParentSheet.ParentBook.CellStyles.ContainsKey(StyleName))
                {
                    //設定の取得
                    ICellStyle Style = ParentSheet.ParentBook.CellStyles[StyleName];
                    //行ループ
                    for (int RIdx = SafeAddress.FirstRow; RIdx <= SafeAddress.LastRow; RIdx++)
                    {
                        //行の取得(なければ生成)
                        IRow row = ParentSheet.PoiSheet.GetRow(RIdx) ?? ParentSheet.PoiSheet.CreateRow(RIdx);
                        //列ループ
                        for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                        {
                            //列の取得(なければ生成)
                            ICell cell = row.GetCell(CIdx) ?? row.CreateCell(CIdx);
                            //スタイルの適用
                            cell.CellStyle = Style;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 列幅の自動調整
        /// </summary>
        public void Autofit()
        {
            //Office.Interop.Excelにならい非連続Rangeの全てに適用
            for (int AIdx = 0; AIdx < SafeAddressList.CountRanges(); AIdx++)
            {
                //アドレス取得
                CellRangeAddress SafeAddress = SafeAddressList.GetCellRangeAddress(AIdx);
                //列幅自動調整ループ
                for (int CIdx = SafeAddress.FirstColumn; CIdx <= SafeAddress.LastColumn; CIdx++)
                {
                    //スッピンのAutoSizeでは独自書式(例えば通貨)の幅が少し足りない。
                    //"\"や"､"の増分が考慮されていないような感じ。
                    //ある程度救済するため、一律28%増の処理を行う
                    ParentSheet.PoiSheet.AutoSizeColumn(CIdx);
                    ParentSheet.PoiSheet.SetColumnWidth(
                        CIdx, ParentSheet.PoiSheet.GetColumnWidth(CIdx) * 128 / 100);
                    //処理対象の行が大量の場合に必要らしい
                    GC.Collect();
                }
            }
        }

        /// <summary>
        /// 指定された条件に合致するRangeを取得する
        /// </summary>
        /// <param name="Type">指定条件</param>
        /// <param name="Value">条件パラメータ</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public Range SpecialCells(XlCellType Type, object Value = null)
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
                    IRow row = ParentSheet.PoiSheet.GetRow(LastRowIndex + CIdx);
                    if (row != null)
                    {
                        //列が存在するならその列を採用
                        if(row.PhysicalNumberOfCells > 0)
                        {
                            RowIndex = LastRowIndex + CIdx;
                            ColumnIndex = row.LastCellNum - 1;
                            break;
                        }
                    }
                }
                //最終カラムのRangeでリターン
                RetVal = new Range(
                    ParentSheet, new CellRangeAddressList(RowIndex, RowIndex, ColumnIndex, ColumnIndex));
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

        /// <summary>
        /// 最大最小を考慮したアドレスを生成する
        /// とはいえ最大値の場合、メモリ不足の例外が発生する可能性が高い
        /// </summary>
        /// <param name="RangeAddress">評価対象Rangeアドレス</param>
        /// <returns>評価済Rangeアドレス</returns>
        private CellRangeAddressList GetSafeCellRangeAddressList(CellRangeAddressList RangeAddressList)
        {
            CellRangeAddressList RetVal = new CellRangeAddressList();
            for (int i = 0; i < RangeAddressList.CountRanges(); i++)
            {
                CellRangeAddress Address = RangeAddressList.GetCellRangeAddress(i);
                //-1の場合は仕様上の最大、最小値を利用する
                if (Address.FirstRow == -1)
                {
                    Address.FirstRow = 0;
                }
                if (Address.LastRow == -1)
                {
                    Address.LastRow = ParentSheet.MaxRowIndex;
                }
                if (Address.FirstColumn == -1)
                {
                    Address.FirstColumn = 0;
                }
                if (Address.LastColumn == -1)
                {
                    Address.LastColumn = ParentSheet.MaxColumnIndex;
                }
                RetVal.AddCellRangeAddress(Address);
            }
            return RetVal;
        }

        /// <summary>
        /// object二次元配列の生成(インデックスを１開始にする)
        /// </summary>
        /// <param name="RowCount">行数</param>
        /// <param name="ColumnCount">列数</param>
        /// <returns>生成された配列</returns>
        private object[,] CreateCellArray(int RowCount, int ColumnCount)
        {
            int[] LowerBoundArray = { 1, 1 };
            int[] LengthArray = { RowCount, ColumnCount };
            object[,] dataArray = (object[,])Array.CreateInstance(typeof(object), LengthArray, LowerBoundArray);
            return dataArray;
        }
    }
}
