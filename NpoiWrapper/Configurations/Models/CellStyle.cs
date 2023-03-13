using NPOI.SS.UserModel;
using System.Xml.Serialization;

namespace Developers.NpoiWrapper.Configurations.Models
{
    /// <summary>
    /// セルスタイル設定
    /// </summary>
    public class CellStyle
    {
        [XmlAttribute("name")] public string Name { get; set; } = string.Empty;
        [XmlElement("border")] public Border Border { get; set; } = new Border(); 
        [XmlElement("align")] public Align Align { get; set; } = new Align();
        [XmlElement("wrapText")] public WrapText WrapText { get; set; } = new WrapText();
        [XmlElement("dataFormat")] public DataFormat DataFormat { get; set; } = new DataFormat();
        [XmlElement("fill")]  public Fill Fill { get; set; } = new Fill();
        [XmlElement("isLocked")] public IsLocked IsLocked { get; set; } = new IsLocked();
    }
    /// <summary>
    /// 罫線
    /// </summary>
    public class Border
    {
        [XmlAttribute("top")] public NPOI.SS.UserModel.BorderStyle Top { get; set; } = BorderStyle.None;
        [XmlAttribute("right")] public NPOI.SS.UserModel.BorderStyle Right { get; set; } = BorderStyle.None;
        [XmlAttribute("bottom")] public NPOI.SS.UserModel.BorderStyle Bottom { get; set; } = BorderStyle.None;
        [XmlAttribute("left")] public NPOI.SS.UserModel.BorderStyle Left { get; set; } = BorderStyle.None;
    }
    /// <summary>
    /// 文字揃え
    /// </summary>
    public class Align
    {
        [XmlAttribute("horizontal")] public HorizontalAlignment Horizontal { get; set; } = HorizontalAlignment.Left;
        [XmlAttribute("vertical")] public VerticalAlignment Vertical { get; set; } = VerticalAlignment.Center;
    }
    /// <summary>
    /// テキストの折り返し
    /// </summary>
    public class WrapText
    {
        [XmlAttribute("value")] public bool Value { get; set; } = false;
    }
    /// <summary>
    /// 表示書式
    /// </summary>
    public class DataFormat
    {
        [XmlAttribute("value")]
        public string Value { get; set; } = string.Empty;
    }
    /// <summary>
    /// 塗りつぶし
    /// </summary>
    public class Fill
    {
        /// <summary>
        /// XMLで指定されている色(文字列)
        /// </summary>
        [XmlAttribute("color")]　public string ColorName { get; set; } = "Automatic";
        /// <summary>
        /// ColorNameからshortの色インデックスを得る。
        /// </summary>
        public short Color
        {
            //IndexdColorはenumではなくクラスプロパティなので直接のデシリアライズはできない
            //文字列のまま保持し参照時にIndexedColorの値を求める
            get
            {
                short RetVal = IndexedColors.Automatic.Index;
                IndexedColors Color = IndexedColors.ValueOf(ColorName);
                if (Color != null)
                {
                    RetVal = Color.Index;
                }
                return RetVal;
            }
        }
        /// <summary>
        /// 塗りつぶしパターン
        /// 色の指定があればSolidForeground、なければNoFill
        /// </summary>
        public FillPattern Pattern
        {
            get
            {
                FillPattern RetVal = FillPattern.NoFill;
                if (Color != IndexedColors.Automatic.Index)
                {
                    RetVal = FillPattern.SolidForeground;
                }
                return RetVal;
            }
        }
    }
    /// <summary>
    /// セルの保護(シートを保護した場合に効果を発揮)
    /// </summary>
    public class IsLocked
    {
        [XmlAttribute("value")] public bool Value { get; set; } = true;
    }
}
