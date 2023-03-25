namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlBorderWeight of Interop.Excel is shown below....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlborderweight
    //----------------------------------------------------------------------------------------------
    //public enum XlBorderWeight
    //{
    //    xlHairline = 1,
    //    xlMedium = -4138,
    //    xlThick = 4,
    //    xlThin = 2
    //}

    /// <summary>
    /// 罫線の太さ
    /// </summary>
    public enum XlBorderWeight : int
    {
        xlHairline = 1,
        xlMedium = -4138,
        xlThick = 4,
        xlThin = 2
    }
}
