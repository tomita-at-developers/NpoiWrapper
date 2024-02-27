using Developers.NpoiWrapper.Model.Param;
using Developers.NpoiWrapper.Utils;
using NPOI.POIFS.Properties;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Developers.NpoiWrapper.Model
{
    internal class RangeBorder
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
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        /// <param name="BordersIndex">XlBordersIndex値</param>
        public RangeBorder(ISheet PoiSheet, CellRangeAddressList SafeAddressList, XlBordersIndex? BordersIndex)
        {
            //親Range情報の保存
            this.PoiSheet = PoiSheet;
            //CellRangeAddressListの保存(安全化して保存)
            this.WholeRangeAddressList = RangeUtil.CreateSafeCellRangeAddressList(SafeAddressList, this.PoiBook.SpreadsheetVersion);
            //Border情報の保存
            this.BordersIndex = BordersIndex;
            //このBordersIndexが担当するRangeAddressの切り出し
            for (int a = 0; a < this.WholeRangeAddressList.CountRanges(); a++)
            {
                IndexdRangeAddressList.Add(
                    new BorderCellRangeAddress(this.WholeRangeAddressList.GetCellRangeAddress(a), this.PoiBook.SpreadsheetVersion)
                            .GetIndexedBorderCellRangeAddress(this.BordersIndex));
            }
        }

        #endregion

        #region "Utils"

        /// <summary>
        /// 親IWorkbook
        /// </summary>
        private IWorkbook PoiBook
        {
            get { return this.PoiSheet.Workbook; }
        }

        /// <summary>
        /// 親ISheet
        /// </summary>
        private ISheet PoiSheet { get; }

        /// <summary>
        /// 親Rangeのアドレスリスト
        /// </summary>
        private CellRangeAddressList WholeRangeAddressList { get; }

        /// <summary>
        /// XlBordersIndexに対応する部分Range
        /// </summary>
        private List<BorderCellRangeAddress> IndexdRangeAddressList { get; } = new List<BorderCellRangeAddress>();

        /// <summary>
        /// このBorderのBorderIndex
        /// このBorderがRange内のどの位置(または種類)を担当するものであるかを示す。
        /// </summary>
        private XlBordersIndex? BordersIndex { get; set; }

        /// <summary>
        /// Cellから読みだした値を集積するリスト
        /// (XlsStyles,XlsWeightsはPoiStylesから生成)
        /// </summary>
        private List<object> PoiStyles { get; set; }
        private List<object> XlsStyles { get; set; }
        private List<object> XlsWeights { get; set; }
        private List<object> ColorIndexes { get; set; }

        /// <summary>
        /// Cellから読みだした値を統合した値。
        /// 統合した結果、値が複数種あればnull, 唯一であればその値がセットされる。
        /// </summary>
        private object CommonXlsStyle { get; set; }
        private object CommonXlsWeight { get; set; }
        private object CommonColorIndex { get; set; }

        #endregion

        #region "methods"

        /// <summary>
        /// Rangeのプロパティ設定値を１つ取得する(全セルで同一値であればその値を、同一でなければnullが取得される。)
        /// </summary>
        /// <param name="Param">取得指示</param>
        /// <returns></returns>
        public object GetCommonProperty(BorderStyleParam Param)
        {
            Dictionary<string, object> Utils = GetCommonUtils(new List<BorderStyleParam> { Param });
            return Utils[Param.Name];
        }

        /// <summary>
        /// Rangeのプロパティ設定値を複数取得する(全セルで同一値であればその値を、同一でなければnullが取得される。)
        /// </summary>
        /// <param name="Params">取得指示リスト</param>
        /// <returns></returns>
        public Dictionary<string, object> GetCommonUtils(List<BorderStyleParam> Params)
        {
            //RangeからBoder情報を読み取る
            GetPoiBorders();
            //リターン値初期化
            Dictionary<string, object> RetVal = new Dictionary<string, object>();
            //要求にしたがいリターン値セット
            foreach (BorderStyleParam Param in Params)
            {
                //BordersIndex別の処理
                if (Param.Name == Utils.StyleName.XlsBorder.LineStyle)
                {
                    RetVal.Add(Param.Name, CommonXlsStyle);
                }
                else if (Param.Name == Utils.StyleName.XlsBorder.Weight)
                {
                    RetVal.Add(Param.Name, CommonXlsWeight);
                }
                else if (Param.Name == Utils.StyleName.XlsBorder.ColorIndex)
                {
                    RetVal.Add(Param.Name, CommonColorIndex);
                }
            }
            return RetVal;
        }

        /// <summary>
        /// Paramで指定された１つのUpdateを実行
        /// </summary>
        /// <param name="Param">Uupdate指示</param>
        public void UpdateProperty(BorderStyleParam Param)
        {
            UpdateUtils(new List<BorderStyleParam>{ Param });
        }

        /// <summary>
        /// Paramsで指定された複数のUpdateを実行
        /// this.IndexdRangeAddressListには必要最低限のRangeのみ格納されている。たとえばEdgeTopであれば先頭行のみ。
        /// なので比較的単純な判断で情報取得できる。
        /// </summary>
        /// <param name="Params">Uupdate指示リスト</param>
        public void UpdateUtils(List<BorderStyleParam> Params)
        {
            //デバッグログ用情報
            var StopwatchForDebugLog = new System.Diagnostics.Stopwatch();
            StopwatchForDebugLog.Start();
            int CellCountForDebugLog = 0;
            string DebugLogString;
            string ParamsForDebugLog = string.Empty;
            foreach (BorderStyleParam Param in Params) { ParamsForDebugLog += Param.GetParamsString() + ","; }
            ParamsForDebugLog = ParamsForDebugLog.TrimEnd(',');
            Logger.Debug("Start processing for Params[" + ParamsForDebugLog + "]");
            //更新履歴管理クラス生成
            Utils.CellStyleUpdateHistory History = new Utils.CellStyleUpdateHistory();
            //デフォルトスタイル取得
            ICellStyle DefaultStyle = this.PoiBook.GetCellStyleAt(0);
            //Areasループ
            foreach (BorderCellRangeAddress Address in this.IndexdRangeAddressList)
            {
                for (int RowIndex = Address.FirstRow; RowIndex <= Address.LastRow; RowIndex++)
                {
                    //列ループ
                    for (int ColumnIndex = Address.FirstColumn; ColumnIndex <= Address.LastColumn; ColumnIndex++)
                    {
                        //デバッグログ用処理カウンタ加算
                        CellCountForDebugLog += 1;
                        //現在のスタイルを取得
                        ICellStyle CurrentStyle = GetCellStyle(RowIndex, ColumnIndex, DefaultStyle);
                        short CurrentIndex = CurrentStyle.Index;
                        //Excel語のパラメータをPOI語に翻訳する。
                        List<CellStyleParam> CellParams = GetParams(Address, RowIndex, ColumnIndex, Params, CurrentStyle);
                        //同じIndexとパラメータの実施履歴がなければ変更処理を実施
                        short IndexToApply = History.Query(CurrentIndex, CellParams);
                        if (IndexToApply == Utils.CellStyleUpdateHistory.None)
                        {
                            //PoiCellStyleの生成
                            Wrapper.PoiCellStyle Style = new Wrapper.PoiCellStyle(this.PoiSheet, CurrentIndex);
                            ///変更処理実行
                            foreach (CellStyleParam p in CellParams)
                            {
                                PropertyInfo CurrentProp;
                                object CurrentObj = Style;
                                //プロパティ名を分割
                                string[] Names = p.Name.Split('.');
                                //プロパティ名ネスト追跡ループ
                                foreach (string Name in Names)
                                {
                                    CurrentProp = CurrentObj.GetType().GetProperty(Name);
                                    if (CurrentProp != null)
                                    {
                                        //最後の名前なら値をセット
                                        if (Names.Last() == Name)
                                        {
                                            CurrentProp.SetValue(CurrentObj, p.Value);
                                        }
                                        //次のオブジェクトに移動
                                        CurrentObj = CurrentProp.GetValue(CurrentObj, null);
                                    }
                                    else
                                    {
                                        throw new ArgumentException("Property name " + p.Name + " is not found.");
                                    }
                                }
                            }
                            //変更のコミット
                            IndexToApply = Style.Commit();
                            //実施履歴追記
                            History.Add(CurrentIndex, CellParams, IndexToApply);
                            //デバッグログ情報
                            DebugLogString = "PoiCelStyle.Commit()";
                        }
                        //更新履歴に存在する場合
                        else
                        {
                            //デバッグログ情報
                            DebugLogString = "History.Query()";
                        }
                        //適用すべきIndexが異なる場合はCellに適用
                        if (CurrentIndex != IndexToApply)
                        {
                            //変更の適用
                            GetOrCreateCell(RowIndex, ColumnIndex).CellStyle = this.PoiBook.GetCellStyleAt(IndexToApply);
                            DebugLogString += " + Apply";
                        }
                        //適用すべきIndexが変化しない場合はスキップ
                        else
                        {
                            DebugLogString += " + Skip";
                        }
                        Logger.Debug("Cell[" + RowIndex + ", " + ColumnIndex + "] : Index[" + CurrentIndex + "] = > Index[" + IndexToApply + "] : Ation[" + DebugLogString + "]");
                    }
                }
            }
            //処理時間測定タイマー停止＆ログ出力
            StopwatchForDebugLog.Stop();
            TimeSpan TimeSpanForDebugLog = StopwatchForDebugLog.Elapsed;
            string Index = "null";
            if (this.BordersIndex is XlBordersIndex bidx)
            {
                Index = bidx.ToString();
            }
            Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for BordersIndex[" + Index + "][" + CellCountForDebugLog + "]Cells / Params[" + ParamsForDebugLog + "]");
        }

        /// <summary>
        /// 罫線情報の読み出し
        /// this.IndexdRangeAddressListには必要最低限のRangeのみ格納されている。たとえばEdgeTopであれば先頭行のみ。
        /// なので比較的単純な判断で情報取得できる。
        /// </summary>
        private void GetPoiBorders()
        {
            //デバッグログ用情報
            var StopwatchForDebugLog = new System.Diagnostics.Stopwatch();
            StopwatchForDebugLog.Start();
            int CellCountForDebugLog = 0;
            //情報初期化
            Initialize();
            //Cell未生成時のデフォルトスタイル
            ICellStyle DefalutStyle = this.PoiBook.GetCellStyleAt(0);
            //Areasループ
            foreach (BorderCellRangeAddress Address in this.IndexdRangeAddressList)
            {
                //行ループ
                for (int r = Address.FirstRow; r <= Address.LastRow; r++)
                {
                    //列
                    for (int c = Address.FirstColumn; c <= Address.LastColumn; c++)
                    {
                        //デバッグログ用処理カウンタ加算
                        CellCountForDebugLog += 1;
                        //CellStyleの取得
                        ICellStyle Style = GetCellStyle(r, c, DefalutStyle);
                        //特定のXlBordersIndexが指定されている場合
                        if (this.BordersIndex != null)
                        {
                            //上端指定なら上罫線
                            if (this.BordersIndex == XlBordersIndex.xlEdgeTop)
                            {
                                PoiStyles.Add(Style.BorderTop);
                                ColorIndexes.Add(Style.TopBorderColor);
                            }
                            //下端指定なら下罫線
                            else if (this.BordersIndex == XlBordersIndex.xlEdgeBottom)
                            {
                                PoiStyles.Add(Style.BorderBottom);
                                ColorIndexes.Add(Style.BottomBorderColor);
                            }
                            //左端指定なら左罫線
                            else if (this.BordersIndex == XlBordersIndex.xlEdgeLeft)
                            {
                                PoiStyles.Add(Style.BorderLeft);
                                ColorIndexes.Add(Style.LeftBorderColor);
                            }
                            //右端指定なら右罫線
                            else if (this.BordersIndex == XlBordersIndex.xlEdgeRight)
                            {
                                PoiStyles.Add(Style.BorderLeft);
                                ColorIndexes.Add(Style.LeftBorderColor);
                            }
                            //内部水平指定
                            else if (this.BordersIndex == XlBordersIndex.xlInsideHorizontal)
                            {
                                //前に行があれば上罫線
                                if (Address.HasPreviousRow(r))
                                {
                                    PoiStyles.Add(Style.BorderTop);
                                    ColorIndexes.Add(Style.TopBorderColor);
                                }
                                //後に行があれば下罫線
                                if (Address.HasNextRow(r))
                                {
                                    PoiStyles.Add(Style.BorderBottom);
                                    ColorIndexes.Add(Style.BottomBorderColor);
                                }
                            }
                            //内部垂直指定
                            else if (this.BordersIndex == XlBordersIndex.xlInsideVertical)
                            {
                                //前に列があれば左罫線
                                if (Address.HasPreviousColumn(c))
                                {
                                    PoiStyles.Add(Style.BorderLeft);
                                    ColorIndexes.Add(Style.LeftBorderColor);
                                }
                                //後に列があれば右罫線
                                if (Address.HasNextColumn(c))
                                {
                                    PoiStyles.Add(Style.BorderRight);
                                    ColorIndexes.Add(Style.RightBorderColor);
                                }
                            }
                            //右下がり斜線
                            else if (this.BordersIndex == XlBordersIndex.xlDiagonalDown)
                            {
                                //右下がり斜線がある場合
                                if (Style.BorderDiagonal == BorderDiagonal.Backward || Style.BorderDiagonal == BorderDiagonal.Both)
                                {
                                    PoiStyles.Add(Style.BorderDiagonalLineStyle);
                                    ColorIndexes.Add(Style.BorderDiagonalColor);
                                }
                                //右下がり斜線がなければ「線なし・色自動」をセット
                                else
                                {
                                    PoiStyles.Add(Style.BorderDiagonalLineStyle);
                                    ColorIndexes.Add(Style.BorderDiagonalColor);
                                }
                            }
                            //右上がり斜線
                            else if (this.BordersIndex == XlBordersIndex.xlDiagonalUp)
                            {
                                //右上がり斜線がある場合
                                if (Style.BorderDiagonal == BorderDiagonal.Forward || Style.BorderDiagonal == BorderDiagonal.Both)
                                {
                                    PoiStyles.Add(Style.BorderDiagonalLineStyle);
                                    ColorIndexes.Add(Style.BorderDiagonalColor);
                                }
                                //右下がり斜線がなければ「線なし・色自動」をセット
                                else
                                {
                                    PoiStyles.Add(Style.BorderDiagonalLineStyle);
                                    ColorIndexes.Add(Style.BorderDiagonalColor);
                                }
                            }
                        }
                        //nullなら全体なので全周囲読み取り
                        else
                        {
                            PoiStyles.Add(Style.BorderTop);
                            ColorIndexes.Add(Style.TopBorderColor);
                            PoiStyles.Add(Style.BorderBottom);
                            ColorIndexes.Add(Style.BottomBorderColor);
                            PoiStyles.Add(Style.BorderLeft);
                            ColorIndexes.Add(Style.LeftBorderColor);
                            PoiStyles.Add(Style.BorderRight);
                            ColorIndexes.Add(Style.RightBorderColor);
                        }
                    }
                }
            }
            //重複除去用
            IEnumerable<object> Distinct;
            object CommonValue;
            //PoiStylesからXlsStylesを生成
            foreach (object PoiStyle in PoiStyles)
            {
                Utils.XlsBorderStyle Xls = new Utils.XlsBorderStyle((BorderStyle)PoiStyle);
                XlsStyles.Add(Xls.LineStyle);
                XlsWeights.Add(Xls.Weight);
            }
            //XlsStyles集約
            Distinct = XlsStyles.Distinct();
            CommonValue = null;
            if (Distinct.Count() == 1)
            {
                CommonValue = Distinct.First();
            }
            CommonXlsStyle = CommonValue;
            //XlsWeights集約
            Distinct = XlsWeights.Distinct();
            CommonValue = null;
            if (Distinct.Count() == 1)
            {
                CommonValue = Distinct.First();
            }
            CommonXlsWeight = CommonValue;
            //ColorIndexess集約
            Distinct = ColorIndexes.Distinct();
            CommonValue = null;
            if (Distinct.Count() == 1)
            {
                CommonValue = Distinct.First();
            }
            CommonColorIndex = CommonValue;
            //処理時間測定タイマー停止＆ログ出力
            StopwatchForDebugLog.Stop();
            TimeSpan TimeSpanForDebugLog = StopwatchForDebugLog.Elapsed;
            string Index = "null";
            if (this.BordersIndex is XlBordersIndex bidx)
            {
                Index = bidx.ToString();
            }
            Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for BordersIndex[" + Index + "][" + CellCountForDebugLog + "]Cells");
        }

        /// <summary>
        /// 情報初期化
        /// </summary>
        private void Initialize()
        {
            //各統合Dictionary作成
            PoiStyles = new List<object>();
            XlsStyles = new List<object>();
            XlsWeights = new List<object>();
            ColorIndexes = new List<object>();
            //全BorderIndex統合情報初期化
            CommonXlsStyle = null;
            CommonXlsWeight = null;
            CommonColorIndex = null;
        }

        /// <summary>
        /// PoiCellStyle更新用パラメータの生成
        /// BorderStyleItemで指定されたExcel語のパラメターをPOI語によるCellStyleItemの指定に変換する。
        /// </summary>
        /// <param name="Address">現在のCellRangeAddress</param>
        /// <param name="RowIndex">このCellの行Index</param>
        /// <param name="ColumnIndex">このCellの列Index</param>
        /// <param name="Params">BorderStyleItemリスト</param>
        /// <param name="CurrentStyle">このCellが現在持っているCellStyle</param>
        /// <returns>変換結果(CellStyleItemのリスト)</returns>
        private List<CellStyleParam> GetParams(
                    BorderCellRangeAddress Address, int RowIndex, int ColumnIndex, List<BorderStyleParam> Params, ICellStyle CurrentStyle)
        {
            List<CellStyleParam> RetVal = new List<CellStyleParam>();
            //Paramsループ
            foreach (BorderStyleParam Param in Params)
            {
                //LineStyle指定
                if (Param.Name == Utils.StyleName.XlsBorder.LineStyle)
                {
                    //特定のBordersIndex指定がある場合
                    if (BordersIndex != null)
                    {
                        //xlEdgeTop指定：Top
                        if (BordersIndex == XlBordersIndex.xlEdgeTop)
                        {
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeBottom指定：Bottom
                        else if (BordersIndex == XlBordersIndex.xlEdgeBottom)
                        {
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeLeft指定：Left
                        else if (BordersIndex == XlBordersIndex.xlEdgeLeft)
                        {
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeRight指定：Right
                        else if (BordersIndex == XlBordersIndex.xlEdgeRight)
                        {
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                        }
                        //xlInsideHorizontal指定
                        else if (BordersIndex == XlBordersIndex.xlInsideHorizontal)
                        {
                            //前に行があればTop
                            if (Address.HasPreviousRow(RowIndex))
                            {
                                RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                            }
                            //後に行があればBottom
                            if(Address.HasNextRow(RowIndex))
                            {
                                RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                            }
                        }
                        //xlInsideVertical指定
                        else if (BordersIndex == XlBordersIndex.xlInsideVertical)
                        {
                            //前に列があればLeft
                            if (Address.HasPreviousColumn(ColumnIndex))
                            {
                                RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                            }
                            //後に列があればRight
                            if (Address.HasNextColumn(ColumnIndex))
                            {
                                RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                            }
                        }
                        //xlDiagonalDown指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalDown)
                        {
                            //Diagonal系(Type:BorderDiagonal.Backward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Backward));
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Diagonal.Style, CurrentStyle, Param.Value));
                        }
                        //xlDiagonalUp指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalUp)
                        {
                            //Diagonal系(Type:BorderDiagonal.Forward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Forward));
                            RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Diagonal.Style, CurrentStyle, Param.Value));
                        }
                    }
                    //BordersIndex指定がなければ全周囲
                    else
                    {
                        RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetChangeStyleParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                    }
                }
                //Weight指定
                else if (Param.Name == Utils.StyleName.XlsBorder.Weight)
                {
                    //特定のBordersIndex指定がある場合
                    if (BordersIndex != null)
                    {
                        //xlEdgeTop指定：Top
                        if (BordersIndex == XlBordersIndex.xlEdgeTop)
                        {
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeBottom指定：Bottom
                        else if (BordersIndex == XlBordersIndex.xlEdgeBottom)
                        {
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeLeft指定：Left
                        else if (BordersIndex == XlBordersIndex.xlEdgeLeft)
                        {
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                        }
                        //xlEdgeRight指定：Right
                        else if (BordersIndex == XlBordersIndex.xlEdgeRight)
                        {
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                        }
                        //xlInsideHorizontal指定
                        else if (BordersIndex == XlBordersIndex.xlInsideHorizontal)
                        {
                            //前に行があればTop
                            if (Address.HasPreviousRow(RowIndex))
                            {
                                RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                            }
                            //後に行があればBottom
                            if (Address.HasNextRow(RowIndex))
                            {
                                RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                            }
                        }
                        //xlInsideVertical指定
                        else if (BordersIndex == XlBordersIndex.xlInsideVertical)
                        {
                            //前に列があればLeft
                            if (Address.HasPreviousColumn(ColumnIndex))
                            {
                                RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                            }
                            //後に列があればRight
                            if (Address.HasNextColumn(ColumnIndex))
                            {
                                RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                            }
                        }
                        //xlDiagonalDown指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalDown)
                        {
                            //Diagonal系(Type:BorderDiagonal.Backward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Backward));
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Diagonal.Style, CurrentStyle, Param.Value));
                        }
                        //xlDiagonalUp指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalUp)
                        {
                            //Diagonal系(Type:BorderDiagonal.Forward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Forward));
                            RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Diagonal.Style, CurrentStyle, Param.Value));
                        }
                    }
                    //BordersIndex指定がなければ全周囲
                    else
                    {
                        RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Top.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Bottom.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Left.Style, CurrentStyle, Param.Value));
                        RetVal.Add(GetAlterWeightParam(Utils.StyleName.PoiBorder.Right.Style, CurrentStyle, Param.Value));
                    }
                }
                //ColorIndex指定
                else if (Param.Name == Utils.StyleName.XlsBorder.ColorIndex)
                {
                    //特定のBordersIndex指定がある場合
                    if (BordersIndex != null)
                    {
                        //xlEdgeTop指定：Top
                        if (BordersIndex == XlBordersIndex.xlEdgeTop)
                        {
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Top.Color, Param.Value));
                        }
                        //xlEdgeBottom指定：Bottom
                        else if (BordersIndex == XlBordersIndex.xlEdgeBottom)
                        {
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Top.Color, Param.Value));
                        }
                        //xlEdgeLeft指定：Left
                        else if (BordersIndex == XlBordersIndex.xlEdgeLeft)
                        {
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Left.Color, Param.Value));
                        }
                        //xlEdgeRight指定：Right
                        else if (BordersIndex == XlBordersIndex.xlEdgeRight)
                        {
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Right.Color, Param.Value));
                        }
                        //xlInsideHorizontal指定
                        else if (BordersIndex == XlBordersIndex.xlInsideHorizontal)
                        {
                            //前に行があればTop
                            if (Address.HasPreviousRow(RowIndex))
                            {
                                RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Top.Color, Param.Value));
                            }
                            //後に行があればBottom
                            if (Address.HasNextRow(RowIndex))
                            {
                                RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Bottom.Color, Param.Value));
                            }
                        }
                        //xlInsideVertical指定
                        else if (BordersIndex == XlBordersIndex.xlInsideVertical)
                        {
                            //前に列があればLeft
                            if (Address.HasPreviousColumn(ColumnIndex))
                            {
                                RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Left.Color, Param.Value));
                            }
                            //後に列があればRight
                            if (Address.HasNextColumn(ColumnIndex))
                            {
                                RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Right.Color, Param.Value));
                            }
                        }
                        //xlDiagonalDown指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalDown)
                        {
                            //Diagonal系(Type:BorderDiagonal.Backward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Backward));
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Color, Param.Value));
                        }
                        //xlDiagonalUp指定
                        else if (BordersIndex == XlBordersIndex.xlDiagonalUp)
                        {
                            //Diagonal系(Type:BorderDiagonal.Forward)
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Type, BorderDiagonal.Forward));
                            RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Diagonal.Color, Param.Value));
                        }
                    }
                    //BordersIndex指定がなければ全周囲
                    else
                    {
                        RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Top.Color, Param.Value));
                        RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Bottom.Color, Param.Value));
                        RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Left.Color, Param.Value));
                        RetVal.Add(new CellStyleParam(Utils.StyleName.PoiBorder.Right.Color, Param.Value));
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// LineStyle更新パラメータ生成
        /// </summary>
        /// <param name="CurrentStyle"></param>
        /// <param name="NewStyle"></param>
        /// <returns></returns>
        private CellStyleParam GetChangeStyleParam(string TargetName, ICellStyle CurrentStyle, object Value)
        {
            BorderStyle CurrentBorderStyle = BorderStyle.None;
            if (TargetName == Utils.StyleName.PoiBorder.Top.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderTop;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Bottom.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderBottom;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Left.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderLeft;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Right.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderRight;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Diagonal.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderDiagonalLineStyle;
            }
            return new CellStyleParam(TargetName, ChangeStyle(CurrentBorderStyle, (XlLineStyle)Value));
        }

        /// <summary>
        /// Weight更新パラメータ生成
        /// </summary>
        /// <param name="TargetName"></param>
        /// <param name="CurrentStyle"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        private CellStyleParam GetAlterWeightParam(string TargetName, ICellStyle CurrentStyle, object Value)
        {
            BorderStyle CurrentBorderStyle = BorderStyle.None;
            if (TargetName == Utils.StyleName.PoiBorder.Top.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderTop;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Bottom.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderBottom;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Left.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderLeft;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Right.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderRight;
            }
            else if (TargetName == Utils.StyleName.PoiBorder.Diagonal.Style)
            {
                CurrentBorderStyle = CurrentStyle.BorderDiagonalLineStyle;
            }
            return new CellStyleParam(TargetName, AlterWeight(CurrentBorderStyle, (XlBorderWeight)Value));
        }

        /// <summary>
        /// 現在のBorderStyleに内在するWeightを維持しつつ、XlLineStyle指定に対応したBorderStyleを選択する。
        /// </summary>
        /// <param name="CurrentStyle"></param>
        /// <param name="NewStyle"></param>
        /// <returns></returns>
        private BorderStyle ChangeStyle(BorderStyle CurrentStyle, XlLineStyle NewStyle)
        {
            //現在のBorderStyleをXlLineStyleとXlBorderWeightに分解
            Utils.XlsBorderStyle XlsStyle = new Utils.XlsBorderStyle(CurrentStyle);
            //新しいXlLineStyleと現在のXlBorderWeightで新しいBorderStyleを生成
            Utils.PoiBorderStyle PoiStyle = new Utils.PoiBorderStyle(NewStyle, XlsStyle.Weight);
            return PoiStyle.BorderStyle;
        }

        /// <summary>
        /// 現在のBorderStyleに内在するLineStyleを維持しつつ、Weight指定に対応したBorderStyleを選択する。
        /// </summary>
        /// <param name="CurrentStyle"></param>
        /// <param name="NewStyle"></param>
        /// <returns></returns>
        private BorderStyle AlterWeight(BorderStyle CurrentStyle, XlBorderWeight NewWeight)
        {
            //現在のBorderStyleをXlLineStyleとXlBorderWeightに分解
            Utils.XlsBorderStyle XlsStyle = new Utils.XlsBorderStyle(CurrentStyle);
            //新しいXlLineStyleと現在のXlBorderWeightで新しいBorderStyleを生成
            Utils.PoiBorderStyle PoiStyle = new Utils.PoiBorderStyle(XlsStyle.LineStyle, NewWeight);
            return PoiStyle.BorderStyle;
        }

        /// <summary>
        /// 指定したセルのCellStyleを取得する
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <param name="DefaultStyle">セルが実在しない場合のデフォルト値</param>
        /// <returns></returns>
        private ICellStyle GetCellStyle(int RowIndex, int ColumnIndex, ICellStyle DefaultStyle)
        {
            ICellStyle RetVal = null;
            IRow Row = this.PoiSheet.GetRow(RowIndex);
            if (Row != null)
            {
                ICell Cell = Row.GetCell(ColumnIndex);
                if (Cell != null)
                {
                    RetVal = Cell.CellStyle;
                }
            }
            return RetVal ?? DefaultStyle;
        }

        /// <summary>
        /// 指定した位置のセルを取得する(なければ生成)
        /// </summary>
        /// <param name="RowIndex"></param>
        /// <param name="ColumnIndex"></param>
        /// <returns></returns>
        private ICell GetOrCreateCell(int RowIndex, int ColumnIndex)
        {
            IRow Row = this.PoiSheet.GetRow(RowIndex);
            if (Row == null)
            {
                Row = this.PoiSheet.CreateRow(RowIndex);
                Logger.Debug(
                    "Sheet[" + this.PoiSheet.SheetName + "]:Row[" + RowIndex + "] *** Row Created. ***");
            }
            ICell Cell = Row.GetCell(ColumnIndex);
            if (Cell == null)
            {
                Cell = Row.CreateCell(ColumnIndex);
                Logger.Debug(
                    "Sheet[" + this.PoiSheet.SheetName + "]:Cell[" + RowIndex + "][" + ColumnIndex + "] *** Column Created. ***");
            }
            return Cell;
        }

        #endregion
    }
}
