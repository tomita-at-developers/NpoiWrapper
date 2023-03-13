using System.Collections.Generic;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations.Models
{
    [XmlRoot("configurations")]
    public class Configurations
    {
        /// <summary>
        /// ページ設定
        /// </summary>
        [XmlArray("pageSetups")]
        [XmlArrayItem("pageSetup")]
        public List<PageSetup> PageSetup { get; set; } = new List<PageSetup>();
        /// <summary>
        /// フォント設定
        /// </summary>
        [XmlElement("font")] public Font Font { get; set; } = new Font();
        /// <summary>
        /// セルのスタイル設定
        /// </summary>
        [XmlArray("cellStyles")]
        [XmlArrayItem("cellStyle")]
        public List<CellStyle> CellStyle { get; set; } = new List<CellStyle>();
    }
}
