using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configuration.Model
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
