using System;
using System.Collections.Generic;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Developers.NpoiWrapper.Styles.Properties;

namespace Developers.NpoiWrapper.Styles.Utils
{

    /// <summary>
    /// CellStyle更新履歴管理クラス
    /// </summary>
    internal class CellStyleUpdateHistory
    {
        #region "constats"

        public const short None = -1;

        #endregion

        #region "fields"

        private readonly List<Log> UpdateLogs = new List<Log>();

        #endregion

        #region "methods"

        /// <summary>
        /// 更新履歴の取り合わせ
        /// </summary>
        /// <param name="StyleIndex">問合せ対象のStyleIndex</param>
        /// <param name="CellStyleParamList">問合せ対象の更新パラメータリスト</param>
        /// <returns></returns>
        public short Query(short StyleIndex, List<CellStyleParam> CellStyleParamList)
        {
            short RetVal = -1;
            //同一Index、同一パラメータでログを検索
            List<Log> QueryResult = UpdateLogs.Where(
                                            log => log.TargetIndex == StyleIndex  
                                                    && log.ParamsString == GetParamsString(CellStyleParamList)).ToList();
            //ヒットするログがあればそのAppliedIndexを返す
            if (QueryResult.Count > 0)
            {
                RetVal = QueryResult[0].AppliedIndex;
            }
            return RetVal;
        }

        /// <summary>
        /// Historyログの追加
        /// </summary>
        /// <param name="StyleIndex">更新前のStyleIndex</param>
        /// <param name="CellStyleParamList">更新パラメータリスト</param>
        /// <param name="AppliedIndex">適用したStyleIndex</param>
        public void Add(short StyleIndex, List<CellStyleParam> CellStyleParamList, short AppliedIndex)
        {
            UpdateLogs.Add(new Log(StyleIndex, GetParamsString(CellStyleParamList), AppliedIndex));
        }

        /// <summary>
        /// List<CellStyleParam>を文字列に変換する
        /// </summary>
        /// <param name="Params">更新パラメータリスト</param>
        /// <returns></returns>
        private string GetParamsString(List<CellStyleParam> Params)
        {
            string RetVal = string.Empty;
            RetVal += "[";
            foreach (CellStyleParam Param in Params)
            {
                RetVal += Param.GetParamsString() + ",";
            }
            RetVal = RetVal.TrimEnd(',') + "]";
            return RetVal;
        }

        #endregion

        #region "classes"
        /// <summary>
        /// CellStyle更新履歴クラス
        /// </summary>
        public class Log
        {
            public short TargetIndex { get; set; }
            public string ParamsString { get; set; }
            public short AppliedIndex { get; set; }
            public Log(short TargetIndex, string ParamsString, short AppliedIndex)
            {
                this.TargetIndex = TargetIndex;
                this.ParamsString = ParamsString;
                this.AppliedIndex = AppliedIndex;
            }
        }

        #endregion
    }
}
