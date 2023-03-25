namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlBordersIndex  of Interop.Excel is shown below....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlbordersindex
    //----------------------------------------------------------------------------------------------
    //public enum XlBordersIndex
    //{
    //    xlInsideHorizontal = 12,
    //    xlInsideVertical = 11,
    //    xlDiagonalDown = 5,
    //    xlDiagonalUp = 6,
    //    xlEdgeBottom = 9,
    //    xlEdgeLeft = 7,
    //    xlEdgeRight = 10,
    //    xlEdgeTop = 8
    //}

    /// <summary>
    /// Bordersのインデックス
    /// </summary>
    public enum XlBordersIndex : int
    {
        xlInsideHorizontal = 12,
        xlInsideVertical = 11,
        xlDiagonalDown = 5,
        xlDiagonalUp = 6,
        xlEdgeBottom = 9,
        xlEdgeLeft = 7,
        xlEdgeRight = 10,
        xlEdgeTop = 8
    }
}