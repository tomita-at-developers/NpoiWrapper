namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlAutoFilterOperator of Interop.Excel is shown below.....
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlautofilteroperator
    //  in hex 58 43 45 4C  = XCEL
    //----------------------------------------------------------------------------------------------
    //public enum XlAutoFilterOperator
    //{
    //    xlAnd = 1,
    //    xlBottom10Items = 4,
    //    xlBottom10Percent = 6,
    //    xlOr = 2,
    //    xlTop10Items = 3,
    //    xlTop10Percent = 5,
    //    xlFilterValues = 7,
    //    xlFilterCellColor = 8,
    //    xlFilterFontColor = 9,
    //    xlFilterIcon = 10,
    //    xlFilterDynamic = 11,
    //    xlFilterNoFill = 12,
    //    xlFilterAutomaticFontColor = 13,
    //    xlFilterNoIcon = 14
    //}

    /// <summary>
    /// ダミーXlAutoFilterOperator
    /// </summary>
    public enum XlAutoFilterOperator
    {
        xlAnd = 1,
        xlBottom10Items = 4,
        xlBottom10Percent = 6,
        xlOr = 2,
        xlTop10Items = 3,
        xlTop10Percent = 5,
        xlFilterValues = 7,
        xlFilterCellColor = 8,
        xlFilterFontColor = 9,
        xlFilterIcon = 10,
        xlFilterDynamic = 11,
        xlFilterNoFill = 12,
        xlFilterAutomaticFontColor = 13,
        xlFilterNoIcon = 14
    }
}
