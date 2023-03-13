using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations.Models
{
    /// <summary>
    /// フォント設定
    /// </summary>
    public class Font
    {
        [XmlAttribute("name")] public string Name { get; set; } = "Yu Gothic UI";
        [XmlAttribute("size")] public double Size { get; set; } = 9;
    }
}
