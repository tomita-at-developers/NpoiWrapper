using NPOI.SS.UserModel;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations.Models
{
    /// <summary>
    /// ページ設定
    /// </summary>
    public class PageSetup
    {
        [XmlAttribute("name")] public string Name { get; set; } = string.Empty;
        [XmlElement("paper")] public Paper Paper { get; set; } = new Paper();
        [XmlElement("scaling")] public Scaling Scaling { get; set; } = new Scaling();
        [XmlElement("margins")] public Margins Margins { get; set; } = new Margins();
        [XmlElement("center")] public Center Center { get; set; } = new Center();
        [XmlElement("titles")] public Titles Titles { get; set; } = new Titles();
    }
    /// <summary>
    /// 用紙設定
    /// </summary>
    public class Paper
    {
        [XmlAttribute("size")] public PaperSize Size { get; set; } = PaperSize.A3;
        [XmlAttribute("landscape")] public bool Landscape { get; set; } = true;
    }
    /// <summary>
    /// 拡大縮小印刷
    /// Adjust, Fitのいずれか一方を期待するので初期値は両方ともnullとしている
    /// </summary>
    public class Scaling
    {
        [XmlElement("adjust")] public Adjust Adjust { get; set; } = null;
        [XmlElement("fit")] public Fit Fit { get; set; } = null;
    }
    /// <summary>
    /// 拡大縮小
    /// </summary>
    public class Adjust
    {
        [XmlAttribute("scale")]
        public short Scale { get; set; } = 100;
    }
    /// <summary>
    /// 次のページ数に合わせて印刷
    /// </summary>
    public class Fit
    {
        [XmlAttribute("wide")] public short Wide { get; set; } = 1;
        [XmlAttribute("tall")] public short Tall { get; set; } = 0;
    }
    /// <summary>
    /// 余白
    /// </summary>
    public class Margins
    {
        [XmlElement("header")] public Header Header { get; set; } = new Header();
        [XmlElement("footer")] public Footer Footer { get; set; } = new Footer();
        [XmlElement("body")] public Body Body { get; set; } = new Body();
    }
    /// <summary>
    /// 余白.ヘッダー
    /// </summary>
    public class Header
    {
        [XmlAttribute("value")]  public double ValueInCentimeter { get; set; } = 1.2;
        public double ValueInInch
        {
            get { return ValueInCentimeter / Constants.InchInCentimeter; }
        }
    }
    /// <summary>
    /// 余白.フッター
    /// </summary>
    public class Footer
    {
        [XmlAttribute("value")] public double ValueInCentimeter { get; set; } = 0.5;
        public double ValueInInch
        {
            get { return ValueInCentimeter / Constants.InchInCentimeter; }
        }
    }
    /// <summary>
    /// 余白.本体
    /// </summary>
    public class Body
    {
        [XmlAttribute("top")] public double TopInCentimeter { get; set; } = 2;
        [XmlAttribute("right")] public double RightInCentimeter { get; set; } = 0.5;
        [XmlAttribute("bottom")] public double BottomInCentimeter { get; set; } = 1.5;
        [XmlAttribute("left")] public double LeftInCentimeter { get; set; } = 0.5;
        public double TopInInch
        {
            get { return TopInCentimeter / Constants.InchInCentimeter; }
        }
        public double RightInInch
        {
            get { return RightInCentimeter / Constants.InchInCentimeter; }
        }
        public double BottomInInch
        {
            get { return BottomInCentimeter / Constants.InchInCentimeter; }
        }
        public double LeftInInch
        {
            get { return LeftInCentimeter / Constants.InchInCentimeter; }
        }
    }
    /// <summary>
    /// ページ中央
    /// </summary>
    public class Center
    {
        [XmlAttribute("horizontally")] public bool Horizontally { get; set; } = true;
        [XmlAttribute("vertically")] public bool Vertically { get; set; } = false;
    }
    /// <summary>
    /// タイトル行/タイトル列
    /// </summary>
    public class Titles
    {
        [XmlAttribute("row")] public string Row { get; set; } = string.Empty;
        [XmlAttribute("column")] public string Column { get; set; } = string.Empty;
    }
}
