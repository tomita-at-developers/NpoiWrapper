using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations
{
    internal class ConfigurationManager
    {

        public Models.Configurations Configs { get; private set; }
        public ConfigurationManager()
        {
            //設定ファイルの読み込み
            XmlSerializer Serializer = new XmlSerializer(typeof(Models.Configurations));
            XmlReaderSettings Settings = new XmlReaderSettings()
            {
                CheckCharacters = false,
            };
            using (StreamReader Reader = new StreamReader(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"NpoiWrapper.config"), Encoding.UTF8))
            using (var xmlReader = XmlReader.Create(Reader, Settings))
            {
                Configs = (Models.Configurations)Serializer.Deserialize(xmlReader);
            }
        }
    }
}
