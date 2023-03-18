using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper.Styles.Utils
{
    internal class ColorPallet
    {
        public short? Index { get; } = null;
        public ColorPallet(object Index)
        {
            if (Index is XlColorIndex EnumValue)
            {
                //ExcelにはAutomaticとNoneの二つがあるがPOIにはAutomaticしかない。
                //(POIのAutomaticの実装は恐らくExcelのNoneなのではないか､､､､)
                this.Index = IndexedColors.Automatic.Index;
            }
            else if (Index is short ShortValue)
            {
                this.Index = ShortValue;
            }
            else if (Index is int IntValue)
            {
                this.Index = (short)IntValue;
            }
            else
            {
                this.Index = null;
            }
        }
    }
}
