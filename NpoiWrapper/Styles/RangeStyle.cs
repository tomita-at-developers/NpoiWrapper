using Developers.NpoiWrapper.Utils;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Developers.NpoiWrapper.Styles
{
    /// <summary>
    /// RangeStyleManger
    /// Range内のCellStyleを取得・更新する
    /// </summary>
    internal class RangeStyle
    {
        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// 親IWorkbook
        /// </summary>
        public IWorkbook PoiBook
        { 
            get { return this.PoiSheet.Workbook; }
        }

        /// <summary>
        /// 親ISheet
        /// </summary>
        public ISheet PoiSheet { get; private set; }

        /// <summary>
        /// 絶対表現(RonwIndex,ColumnIndexとして直接利用可能)されたアドレスリスト
        /// </summary>
        public CellRangeAddressList SafeRangeAddressList { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiSheet">ISheetインスタンス</param>
        /// <param name="SafeAddressList">CellRangeAddressListクラスインスタンス</param>
        public RangeStyle(ISheet PoiSheet, CellRangeAddressList SafeAddressList)
        {
            //親Range情報の保存
            this.PoiSheet = PoiSheet;
            this.SafeRangeAddressList = SafeAddressList;
        }

        /// <summary>
        /// プロパティ名で指定されたICellStyleのプロパティをRangeから取得する
        /// Rangeが複数のCellを持つ場合は、全Cellで同じ値が設定されている場合はその値を返す。
        /// 全Cellで同じ値が設定されていない場合はNULLを返す。
        /// </summary>
        /// <param name="Param">取得指示</param>
        /// <returns></returns>
        public object GetCommonProperty(Properties.CellStyleParam Param)
        {
            Dictionary<string, object> Properties = GetCommonProperties(new List<Properties.CellStyleParam> { Param });
            return Properties[Param.Name];
        }

        /// <summary>
        /// プロパティ名で指定されたICellStyleのプロパティをRangeから取得する
        /// Rangeが複数のCellを持つ場合は、全Cellで同じ値が設定されている場合はその値を返す。
        /// 全Cellで同じ値が設定されていない場合はNULLを返す。
        /// </summary>
        /// <param name="Params">取得指示リスト</param>
        /// <returns></returns>
        public Dictionary<string, object> GetCommonProperties(List<Properties.CellStyleParam> Params)
        {
            //デバッグログ用情報
            var StopwatchForDebugLog = new System.Diagnostics.Stopwatch();
            StopwatchForDebugLog.Start();
            int CellCountForDebugLog = 0;
            string ParamsForDebugLog = string.Empty;
            foreach (Properties.CellStyleParam Param in Params) { ParamsForDebugLog += Param.GetParamsString() + ","; }
            ParamsForDebugLog = ParamsForDebugLog.TrimEnd(',');
            Logger.Debug("Start processing for Params[" + ParamsForDebugLog + "]");
            //リターン値、集積エリアの生成と初期化
            Dictionary<string, object> RetVal = new Dictionary<string, object>();
            Dictionary<string, List<object>> Values = new Dictionary<string, List<object>>();
            foreach (Properties.CellStyleParam Param in Params)
            {
                RetVal.Add(Param.Name, null);
                Values.Add(Param.Name, new List<object>());
            }
            //中断判定用
            bool ProcBreak = false;
            //デフォルトスタイル取得(Cell未生成時に使用)
            ICellStyle DefaultStyle = PoiBook.GetCellStyleAt(0);
            //Areasループ
            for (int a = 0; a < SafeRangeAddressList.CountRanges(); a++)
            {
                //Areasアドレス取得
                CellRangeAddress Address = SafeRangeAddressList.GetCellRangeAddress(a);
                //行ループ
                for (int RowIndex = Address.FirstRow; RowIndex <= Address.LastRow; RowIndex++)
                {
                    //列ループ
                    for (int ColumnIndex = Address.FirstColumn; ColumnIndex <= Address.LastColumn; ColumnIndex++)
                    {
                        //デバッグログ用処理カウンタ加算
                        CellCountForDebugLog += 1;
                        //PoiCellStyleの生成
                        Models.PoiCellStyle Style = new Models.PoiCellStyle(
                                this.PoiSheet, GetCellStyle(RowIndex, ColumnIndex, DefaultStyle).Index);
                        ///プロパティ取得ループ
                        foreach (Properties.CellStyleParam p in Params)
                        {
                            //プロパティ情報
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
                                    CurrentObj = CurrentProp.GetValue(CurrentObj);
                                }
                                else
                                {
                                    throw new ArgumentException("Property name " + p.Name + " is not found.");
                                }
                            }
                            //プロパティ値を集積
                            Values[p.Name].Add(CurrentObj);
                        }
                        //集積結果を集約
                        int ContinueCount = 0;
                        foreach (KeyValuePair<string, List<object>> PropValues in Values)
                        {
                            IEnumerable<object> Distinct = PropValues.Value.Distinct();
                            if (Distinct.Count() == 1)
                            {
                                ContinueCount += 1;
                            }

                        }
                        //続行不要(全プロパティで複数種発見)ならば中断指示オン
                        if (ContinueCount == 0)
                        {
                            ProcBreak = true;
                            break;
                        }
                    }
                    //中断指示があれば終了
                    if (ProcBreak)
                    {
                        break;
                    }
                }
                //中断指示があれば終了
                if (ProcBreak)
                {
                    break;
                }
            }
            //集積結果の集約
            foreach (KeyValuePair<string, List<object>> PropValues in Values)
            {
                IEnumerable<object> Distinct = PropValues.Value.Distinct();
                if (Distinct.Count() == 1)
                {
                    RetVal[PropValues.Key] = Distinct.First();
                }

            }
            //処理時間測定タイマー停止＆ログ出力
            StopwatchForDebugLog.Stop();
            TimeSpan TimeSpanForDebugLog = StopwatchForDebugLog.Elapsed;
            Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for [" + CellCountForDebugLog + "]Cells / Params[" + ParamsForDebugLog + "]");
            return RetVal;
        }

        /// <summary>
        /// Paramで指定された１つのUpdateを実行
        /// </summary>
        /// <param name="Param">Uupdate指示</param>
        public void UpdateProperty(Properties.CellStyleParam Param)
        {
            UpdateProperties(new List<Properties.CellStyleParam> { Param });
        }

        /// <summary>
        /// Paramsで指定されたUpdateを実行
        /// </summary>
        /// <param name="Params">Uupdate指示リスト</param>
        public void UpdateProperties(List<Properties.CellStyleParam> Params)
        {
            //デバッグログ用情報
            var StopwatchForDebugLog = new System.Diagnostics.Stopwatch();
            StopwatchForDebugLog.Start();
            int CellCountForDebugLog = 0;
            string DebugLogString;
            string ParamsForDebugLog = string.Empty;
            foreach (Properties.CellStyleParam Param in Params) { ParamsForDebugLog += Param.GetParamsString() + ","; }
            ParamsForDebugLog = ParamsForDebugLog.TrimEnd(',');
            Logger.Debug("Start processing for Params[" + ParamsForDebugLog + "]");
            //更新履歴管理クラス生成
            Utils.CellStyleUpdateHistory History = new Utils.CellStyleUpdateHistory();
            //デフォルトスタイル取得
            ICellStyle DefaultStyle = this.PoiBook.GetCellStyleAt(0);
            //Areaループ
            for (int a = 0; a < SafeRangeAddressList.CountRanges(); a++)
            {
                //行ループ
                CellRangeAddress Address = SafeRangeAddressList.GetCellRangeAddress(a);
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
                        //同じIndexとパラメータの実施履歴がなければ変更処理を実施
                        short IndexToApply = History.Query(CurrentIndex, Params);
                        if (IndexToApply == Utils.CellStyleUpdateHistory.None)
                        {
                            //PoiCellStyleの生成
                            Models.PoiCellStyle Style = new Models.PoiCellStyle(this.PoiSheet, CurrentIndex);
                            ///変更処理実行
                            foreach (Properties.CellStyleParam p in Params)
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
                            History.Add(CurrentIndex, Params, IndexToApply);
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
            Logger.Debug("Processing Time[" + TimeSpanForDebugLog.ToString(@"ss\.fff") + "sec] for [" + CellCountForDebugLog + "]Cells / Params[" + ParamsForDebugLog + "]");
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
            IRow Row = this.PoiSheet.GetRow(RowIndex) ?? this.PoiSheet.CreateRow(RowIndex);
            return Row.GetCell(ColumnIndex) ?? Row.CreateCell(ColumnIndex);
        }
    }
}
