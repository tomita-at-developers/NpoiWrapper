using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper.Model.Utils
{
    /// <summary>
    /// objectで指定されたColorIndexをshort?にキャストする。
    /// </summary>
    internal class ColorPallet
    {
        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="Index">objectで指定されたColorIndex</param>
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

        #endregion

        #region "properties"

        /// <summary>
        /// objectからshortにキャストされたColorIndex値
        /// </summary>
        public short? Index { get; } = null;

        #endregion
    }
}
