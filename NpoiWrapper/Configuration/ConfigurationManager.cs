using NPOI.SS.UserModel;
using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Collections.Generic;
using Developers.NpoiWrapper.Configuration.Model;

namespace Developers.NpoiWrapper.Configuration
{
    internal class ConfigurationManager
    {

        public Model.Configurations Configs { get; private set; }
        public ConfigurationManager()
        {
            //設定ファイルの読み込み
            XmlSerializer Serializer = new XmlSerializer(typeof(Model.Configurations));
            XmlReaderSettings Settings = new XmlReaderSettings()
            {
                CheckCharacters = false,
            };
            using (StreamReader Reader = new StreamReader(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"NpoiWrapper.config"), Encoding.UTF8))
            using (var xmlReader = XmlReader.Create(Reader, Settings))
            {
                Configs = (Model.Configurations)Serializer.Deserialize(xmlReader);
            }
        }
    }
}
