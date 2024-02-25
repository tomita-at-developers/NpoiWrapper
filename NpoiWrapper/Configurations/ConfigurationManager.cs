using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations
{
    /// <summary>
    /// 設定ファイル読取
    /// </summary>
    internal class ConfigurationManager
    {
        #region "constants"

        private const string CONFIG_FILENAME = @"NpoiWrapperStyle.config";

        #endregion

        #region "fields"

        private readonly Models.Configurations _Configs = new Models.Configurations();

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// 設定ファイルを読み込んでConfigsにデシリアライズする。
        /// </summary>
        public ConfigurationManager()
        {
            //フルパスファイル名生成
            string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CONFIG_FILENAME);
            //ファイルが存在するなら読み取る
            if (System.IO.File.Exists(ConfigPath))
            {
                XmlSerializer Serializer = new XmlSerializer(typeof(Models.Configurations));
                XmlReaderSettings Settings = new XmlReaderSettings()
                {
                    CheckCharacters = false,
                };
                using (StreamReader Reader = new StreamReader(ConfigPath, Encoding.UTF8))
                using (var xmlReader = XmlReader.Create(Reader, Settings))
                {
                    _Configs = (Models.Configurations)Serializer.Deserialize(xmlReader);
                }
            }
        }

        #endregion

        #region "properties"

        /// <summary>
        /// 設定ファイル読取結果
        /// </summary>
        public Models.Configurations Configs
        { 
            get { return _Configs; }
        }

        #endregion
    }
}
