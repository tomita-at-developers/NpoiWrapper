namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlFileFormat of Interop.Excel is shown below....
    //----------------------------------------------------------------------------------------------
    //public enum XlFileFormat
    //{
    //    //
    //    // 概要:
    //    //     Microsoft Office Excel Add-In.
    //    xlAddIn = 18,
    //    //
    //    // 概要:
    //    //     Comma separated value.
    //    xlCSV = 6,
    //    //
    //    // 概要:
    //    //     Comma separated value.
    //    xlCSVMac = 22,
    //    //
    //    // 概要:
    //    //     Comma separated value.
    //    xlCSVMSDOS = 24,
    //    //
    //    // 概要:
    //    //     Comma separated value.
    //    xlCSVWindows = 23,
    //    //
    //    // 概要:
    //    //     Dbase 2 format.
    //    xlDBF2 = 7,
    //    //
    //    // 概要:
    //    //     Dbase 3 format.
    //    xlDBF3 = 8,
    //    //
    //    // 概要:
    //    //     Dbase 4 format.
    //    xlDBF4 = 11,
    //    //
    //    // 概要:
    //    //     Data Interchange format.
    //    xlDIF = 9,
    //    //
    //    // 概要:
    //    //     Excel version 2.0.
    //    xlExcel2 = 0x10,
    //    //
    //    // 概要:
    //    //     Excel version 2.0 far east.
    //    xlExcel2FarEast = 27,
    //    //
    //    // 概要:
    //    //     Excel version 3.0.
    //    xlExcel3 = 29,
    //    //
    //    // 概要:
    //    //     Excel version 4.0.
    //    xlExcel4 = 33,
    //    //
    //    // 概要:
    //    //     Excel version 5.0.
    //    xlExcel5 = 39,
    //    //
    //    // 概要:
    //    //     Excel 95.
    //    xlExcel7 = 39,
    //    //
    //    // 概要:
    //    //     Excel version 95 and 97.
    //    xlExcel9795 = 43,
    //    //
    //    // 概要:
    //    //     Excel version 4.0. Workbook format.
    //    xlExcel4Workbook = 35,
    //    //
    //    // 概要:
    //    //     Microsoft Office Excel Add-In international format.
    //    xlIntlAddIn = 26,
    //    //
    //    // 概要:
    //    //     Deprecated format.
    //    xlIntlMacro = 25,
    //    //
    //    // 概要:
    //    //     Excel workbook format.
    //    xlWorkbookNormal = -4143,
    //    //
    //    // 概要:
    //    //     Symbolic link format.
    //    xlSYLK = 2,
    //    //
    //    // 概要:
    //    //     Excel template format.
    //    xlTemplate = 17,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format
    //    xlCurrentPlatformText = -4158,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format.
    //    xlTextMac = 19,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format.
    //    xlTextMSDOS = 21,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format.
    //    xlTextPrinter = 36,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format.
    //    xlTextWindows = 20,
    //    //
    //    // 概要:
    //    //     Deprecated format.
    //    xlWJ2WD1 = 14,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK1 = 5,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK1ALL = 0x1F,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK1FMT = 30,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK3 = 0xF,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK4 = 38,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWK3FM3 = 0x20,
    //    //
    //    // 概要:
    //    //     Lotus 1-2-3 format.
    //    xlWKS = 4,
    //    //
    //    // 概要:
    //    //     Microsoft Works 2.0 format
    //    xlWorks2FarEast = 28,
    //    //
    //    // 概要:
    //    //     Quattro Pro format.
    //    xlWQ1 = 34,
    //    //
    //    // 概要:
    //    //     Deprecated format.
    //    xlWJ3 = 40,
    //    //
    //    // 概要:
    //    //     Deprecated format.
    //    xlWJ3FJ3 = 41,
    //    //
    //    // 概要:
    //    //     Specifies a type of text format.
    //    xlUnicodeText = 42,
    //    //
    //    // 概要:
    //    //     Web page format.
    //    xlHtml = 44,
    //    //
    //    // 概要:
    //    //     MHT format.
    //    xlWebArchive = 45,
    //    //
    //    // 概要:
    //    //     Excel Spreadsheet format.
    //    xlXMLSpreadsheet = 46
    //}
    public enum XlFileFormat
    {
        //
        // 概要:
        //     Microsoft Office Excel Add-In.
        xlAddIn = 18,
        //
        // 概要:
        //     Comma separated value.
        xlCSV = 6,
        //
        // 概要:
        //     Comma separated value.
        xlCSVMac = 22,
        //
        // 概要:
        //     Comma separated value.
        xlCSVMSDOS = 24,
        //
        // 概要:
        //     Comma separated value.
        xlCSVWindows = 23,
        //
        // 概要:
        //     Dbase 2 format.
        xlDBF2 = 7,
        //
        // 概要:
        //     Dbase 3 format.
        xlDBF3 = 8,
        //
        // 概要:
        //     Dbase 4 format.
        xlDBF4 = 11,
        //
        // 概要:
        //     Data Interchange format.
        xlDIF = 9,
        //
        // 概要:
        //     Excel version 2.0.
        xlExcel2 = 0x10,
        //
        // 概要:
        //     Excel version 2.0 far east.
        xlExcel2FarEast = 27,
        //
        // 概要:
        //     Excel version 3.0.
        xlExcel3 = 29,
        //
        // 概要:
        //     Excel version 4.0.
        xlExcel4 = 33,
        //
        // 概要:
        //     Excel version 5.0.
        xlExcel5 = 39,
        //
        // 概要:
        //     Excel 95.
        xlExcel7 = 39,
        //
        // 概要:
        //     Excel version 95 and 97.
        xlExcel9795 = 43,
        //
        // 概要:
        //     Excel version 4.0. Workbook format.
        xlExcel4Workbook = 35,
        //
        // 概要:
        //     Microsoft Office Excel Add-In international format.
        xlIntlAddIn = 26,
        //
        // 概要:
        //     Deprecated format.
        xlIntlMacro = 25,
        //
        // 概要:
        //     Excel workbook format.
        xlWorkbookNormal = -4143,
        //
        // 概要:
        //     Symbolic link format.
        xlSYLK = 2,
        //
        // 概要:
        //     Excel template format.
        xlTemplate = 17,
        //
        // 概要:
        //     Specifies a type of text format
        xlCurrentPlatformText = -4158,
        //
        // 概要:
        //     Specifies a type of text format.
        xlTextMac = 19,
        //
        // 概要:
        //     Specifies a type of text format.
        xlTextMSDOS = 21,
        //
        // 概要:
        //     Specifies a type of text format.
        xlTextPrinter = 36,
        //
        // 概要:
        //     Specifies a type of text format.
        xlTextWindows = 20,
        //
        // 概要:
        //     Deprecated format.
        xlWJ2WD1 = 14,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK1 = 5,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK1ALL = 0x1F,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK1FMT = 30,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK3 = 0xF,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK4 = 38,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWK3FM3 = 0x20,
        //
        // 概要:
        //     Lotus 1-2-3 format.
        xlWKS = 4,
        //
        // 概要:
        //     Microsoft Works 2.0 format
        xlWorks2FarEast = 28,
        //
        // 概要:
        //     Quattro Pro format.
        xlWQ1 = 34,
        //
        // 概要:
        //     Deprecated format.
        xlWJ3 = 40,
        //
        // 概要:
        //     Deprecated format.
        xlWJ3FJ3 = 41,
        //
        // 概要:
        //     Specifies a type of text format.
        xlUnicodeText = 42,
        //
        // 概要:
        //     Web page format.
        xlHtml = 44,
        //
        // 概要:
        //     MHT format.
        xlWebArchive = 45,
        //
        // 概要:
        //     Excel Spreadsheet format.
        xlXMLSpreadsheet = 46
    }

}

