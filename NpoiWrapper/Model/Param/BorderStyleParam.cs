﻿namespace Developers.NpoiWrapper.Model.Param
{
    /// <summary>
    /// 罫線情報の参照/更新パラメータ
    /// Excxel語のパラメータ。PoiCellStyle更新時はCellStyleParamに変換される。
    /// </summary>
    internal class BorderStyleParam
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ(参照用)
        /// </summary>
        /// <param name="Name"></param>
        public BorderStyleParam(string Name)
        {
            this.Name = Name;
        }

        /// <summary>
        /// コンストラクタ(更新用)
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Value"></param>
        public BorderStyleParam(string Name, object Value)
            : this(Name)
        {
            this.Value = Value;
            this.Update = true;
        }

        #endregion

        #region "properties"

        /// <summary>
        /// 対象名
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// 適用する値
        /// </summary>
        public object Value { get; private set; } = null;
        /// <summary>
        /// 更新フラグ
        /// </summary>
        public bool Update { get; private set; } = false;

        #endregion

        #region "methods"

        /// <summary>
        /// 文字列に変換する
        /// </summary>
        /// <returns></returns>
        public string GetParamsString()
        {
            string RetVal = "[" + Name;
            if (Value != null)
            {
                RetVal += "=" + Value.ToString();
            }
            RetVal += "]";
            return RetVal;
        }

        #endregion
    }
}