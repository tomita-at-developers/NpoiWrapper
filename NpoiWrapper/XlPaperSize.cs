using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    //  XlPaperSize in Interop.Excel is shown below...
    //  https://learn.microsoft.com/en-us/office/vba/api/excel.xlpapersize
    //----------------------------------------------------------------------------------------------
    //public enum XlPaperSize
    //{
    //    //
    //    // 概要:
    //    //     10 in. x 14 in.
    //    xlPaper10x14 = 0x10,
    //    //
    //    // 概要:
    //    //     11 in. x 17 in.
    //    xlPaper11x17 = 17,
    //    //
    //    // 概要:
    //    //     A3 (297 mm x 420 mm)
    //    xlPaperA3 = 8,
    //    //
    //    // 概要:
    //    //     A4 (210 mm x 297 mm)
    //    xlPaperA4 = 9,
    //    //
    //    // 概要:
    //    //     A4 Small (210 mm x 297 mm)
    //    xlPaperA4Small = 10,
    //    //
    //    // 概要:
    //    //     A5 (148 mm x 210 mm)
    //    xlPaperA5 = 11,
    //    //
    //    // 概要:
    //    //     B4 (250 mm x 354 mm)
    //    xlPaperB4 = 12,
    //    //
    //    // 概要:
    //    //     A5 (148 mm x 210 mm)
    //    xlPaperB5 = 13,
    //    //
    //    // 概要:
    //    //     C size sheet
    //    xlPaperCsheet = 24,
    //    //
    //    // 概要:
    //    //     D size sheet
    //    xlPaperDsheet = 25,
    //    //
    //    // 概要:
    //    //     Envelope #10 (4-1/8 in. x 9-1/2 in.)
    //    xlPaperEnvelope10 = 20,
    //    //
    //    // 概要:
    //    //     Envelope #11 (4-1/2 in. x 10-3/8 in.)
    //    xlPaperEnvelope11 = 21,
    //    //
    //    // 概要:
    //    //     Envelope #12 (4-1/2 in. x 11 in.)
    //    xlPaperEnvelope12 = 22,
    //    //
    //    // 概要:
    //    //     Envelope #14 (5 in. x 11-1/2 in.)
    //    xlPaperEnvelope14 = 23,
    //    //
    //    // 概要:
    //    //     Envelope #9 (3-7/8 in. x 8-7/8 in.)
    //    xlPaperEnvelope9 = 19,
    //    //
    //    // 概要:
    //    //     Envelope B4 (250 mm x 353 mm)
    //    xlPaperEnvelopeB4 = 33,
    //    //
    //    // 概要:
    //    //     Envelope B5 (176 mm x 250 mm)
    //    xlPaperEnvelopeB5 = 34,
    //    //
    //    // 概要:
    //    //     Envelope B6 (176 mm x 125 mm)
    //    xlPaperEnvelopeB6 = 35,
    //    //
    //    // 概要:
    //    //     Envelope C3 (324 mm x 458 mm)
    //    xlPaperEnvelopeC3 = 29,
    //    //
    //    // 概要:
    //    //     Envelope C4 (229 mm x 324 mm)
    //    xlPaperEnvelopeC4 = 30,
    //    //
    //    // 概要:
    //    //     Envelope C5 (162 mm x 229 mm)
    //    xlPaperEnvelopeC5 = 28,
    //    //
    //    // 概要:
    //    //     Envelope C6 (114 mm x 162 mm)
    //    xlPaperEnvelopeC6 = 0x1F,
    //    //
    //    // 概要:
    //    //     Envelope C65 (114 mm x 229 mm)
    //    xlPaperEnvelopeC65 = 0x20,
    //    //
    //    // 概要:
    //    //     Envelope DL (110 mm x 220 mm)
    //    xlPaperEnvelopeDL = 27,
    //    //
    //    // 概要:
    //    //     Envelope (110 mm x 230 mm)
    //    xlPaperEnvelopeItaly = 36,
    //    //
    //    // 概要:
    //    //     Envelope Monarch (3-7/8 in. x 7-1/2 in.)
    //    xlPaperEnvelopeMonarch = 37,
    //    //
    //    // 概要:
    //    //     Envelope (3-5/8 in. x 6-1/2 in.)
    //    xlPaperEnvelopePersonal = 38,
    //    //
    //    // 概要:
    //    //     E size sheet
    //    xlPaperEsheet = 26,
    //    //
    //    // 概要:
    //    //     Executive (7-1/2 in. x 10-1/2 in.)
    //    xlPaperExecutive = 7,
    //    //
    //    // 概要:
    //    //     German Legal Fanfold (8-1/2 in. x 13 in.)
    //    xlPaperFanfoldLegalGerman = 41,
    //    //
    //    // 概要:
    //    //     German Legal Fanfold (8-1/2 in. x 13 in.)
    //    xlPaperFanfoldStdGerman = 40,
    //    //
    //    // 概要:
    //    //     U.S. Standard Fanfold (14-7/8 in. x 11 in.)
    //    xlPaperFanfoldUS = 39,
    //    //
    //    // 概要:
    //    //     Folio (8-1/2 in. x 13 in.)
    //    xlPaperFolio = 14,
    //    //
    //    // 概要:
    //    //     Ledger (17 in. x 11 in.)
    //    xlPaperLedger = 4,
    //    //
    //    // 概要:
    //    //     Legal (8-1/2 in. x 14 in.)
    //    xlPaperLegal = 5,
    //    //
    //    // 概要:
    //    //     Letter (8-1/2 in. x 11 in.)
    //    xlPaperLetter = 1,
    //    //
    //    // 概要:
    //    //     Letter Small (8-1/2 in. x 11 in.)
    //    xlPaperLetterSmall = 2,
    //    //
    //    // 概要:
    //    //     Note (8-1/2 in. x 11 in.)
    //    xlPaperNote = 18,
    //    //
    //    // 概要:
    //    //     Quarto (215 mm x 275 mm)
    //    xlPaperQuarto = 0xF,
    //    //
    //    // 概要:
    //    //     Statement (5-1/2 in. x 8-1/2 in.)
    //    xlPaperStatement = 6,
    //    //
    //    // 概要:
    //    //     Tabloid (11 in. x 17 in.)
    //    xlPaperTabloid = 3,
    //    //
    //    // 概要:
    //    //     User-defined
    //    xlPaperUser = 0x100
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding definition in NPOI is shown below...
    //----------------------------------------------------------------------------------------------
    //public enum PaperSize : short
    //{
    //    /// <summary>
    //    /// Allow accessing the Initial value.
    //    /// </summary>
    //    PRINTER_DEFAULT_PAPERSIZE = 0,
    //    US_Letter_Small = 1,
    //    US_Tabloid = 2,
    //    US_Ledger = 3,
    //    US_Legal = 4,
    //    US_Statement = 5,
    //    US_Executive = 6,
    //    A3 = 7,
    //    A4 = 8,
    //    A4_Small = 9,
    //    A5 = 10,
    //    B4 = 11,
    //    B5 = 12,
    //    Folio = 13,
    //    Quarto = 14,
    //    TEN_BY_FOURTEEN = 15,
    //    ELEVEN_BY_SEVENTEEN = 16,
    //    US_Note = 17,
    //    US_Envelope_9 = 18,
    //    US_Envelope_10 = 19,
    //    US_Envelope_11 = 20,
    //    US_Envelope_12 = 21,
    //    US_Envelope_14 = 22,
    //    C_Size_Sheet = 23,
    //    D_Size_Sheet = 24,
    //    E_Size_Sheet = 25,
    //    Envelope_DL = 26,
    //    Envelope_C5 = 27,
    //    Envelope_C3 = 28,
    //    Envelope_C4 = 29,
    //    Envelope_C6 = 30,
    //    Envelope_MONARCH = 31,
    //    A4_EXTRA = 53,
    //    /// <summary>
    //    /// A4 Transverse - 210x297 mm 
    //    /// </summary>
    //    A4_TRANSVERSE_PAPERSIZE = 55,
    //    /// <summary>
    //    /// A4 Plus - 210x330 mm 
    //    /// </summary>
    //    A4_PLUS_PAPERSIZE = 60,
    //    /// <summary>
    //    /// US Letter Rotated 11 x 8 1/2 in
    //    /// </summary>
    //    LETTER_ROTATED_PAPERSIZE = 75,
    //    /// <summary>
    //    /// A4 Rotated - 297x210 mm */
    //    /// </summary>
    //    A4_ROTATED_PAPERSIZE = 77
    //}

    public enum XlPaperSize
    {
        //
        // 概要:
        //     10 in. x 14 in.
        xlPaper10x14 = 0x10,
        //
        // 概要:
        //     11 in. x 17 in.
        xlPaper11x17 = 17,
        //
        // 概要:
        //     A3 (297 mm x 420 mm)
        xlPaperA3 = 8,
        //
        // 概要:
        //     A4 (210 mm x 297 mm)
        xlPaperA4 = 9,
        //
        // 概要:
        //     A4 Small (210 mm x 297 mm)
        xlPaperA4Small = 10,
        //
        // 概要:
        //     A5 (148 mm x 210 mm)
        xlPaperA5 = 11,
        //
        // 概要:
        //     B4 (250 mm x 354 mm)
        xlPaperB4 = 12,
        //
        // 概要:
        //     A5 (148 mm x 210 mm)
        xlPaperB5 = 13,
        //
        // 概要:
        //     C size sheet
        xlPaperCsheet = 24,
        //
        // 概要:
        //     D size sheet
        xlPaperDsheet = 25,
        //
        // 概要:
        //     Envelope #10 (4-1/8 in. x 9-1/2 in.)
        xlPaperEnvelope10 = 20,
        //
        // 概要:
        //     Envelope #11 (4-1/2 in. x 10-3/8 in.)
        xlPaperEnvelope11 = 21,
        //
        // 概要:
        //     Envelope #12 (4-1/2 in. x 11 in.)
        xlPaperEnvelope12 = 22,
        //
        // 概要:
        //     Envelope #14 (5 in. x 11-1/2 in.)
        xlPaperEnvelope14 = 23,
        //
        // 概要:
        //     Envelope #9 (3-7/8 in. x 8-7/8 in.)
        xlPaperEnvelope9 = 19,
        //
        // 概要:
        //     Envelope B4 (250 mm x 353 mm)
        xlPaperEnvelopeB4 = 33,
        //
        // 概要:
        //     Envelope B5 (176 mm x 250 mm)
        xlPaperEnvelopeB5 = 34,
        //
        // 概要:
        //     Envelope B6 (176 mm x 125 mm)
        xlPaperEnvelopeB6 = 35,
        //
        // 概要:
        //     Envelope C3 (324 mm x 458 mm)
        xlPaperEnvelopeC3 = 29,
        //
        // 概要:
        //     Envelope C4 (229 mm x 324 mm)
        xlPaperEnvelopeC4 = 30,
        //
        // 概要:
        //     Envelope C5 (162 mm x 229 mm)
        xlPaperEnvelopeC5 = 28,
        //
        // 概要:
        //     Envelope C6 (114 mm x 162 mm)
        xlPaperEnvelopeC6 = 0x1F,
        //
        // 概要:
        //     Envelope C65 (114 mm x 229 mm)
        xlPaperEnvelopeC65 = 0x20,
        //
        // 概要:
        //     Envelope DL (110 mm x 220 mm)
        xlPaperEnvelopeDL = 27,
        //
        // 概要:
        //     Envelope (110 mm x 230 mm)
        xlPaperEnvelopeItaly = 36,
        //
        // 概要:
        //     Envelope Monarch (3-7/8 in. x 7-1/2 in.)
        xlPaperEnvelopeMonarch = 37,
        //
        // 概要:
        //     Envelope (3-5/8 in. x 6-1/2 in.)
        xlPaperEnvelopePersonal = 38,
        //
        // 概要:
        //     E size sheet
        xlPaperEsheet = 26,
        //
        // 概要:
        //     Executive (7-1/2 in. x 10-1/2 in.)
        xlPaperExecutive = 7,
        //
        // 概要:
        //     German Legal Fanfold (8-1/2 in. x 13 in.)
        xlPaperFanfoldLegalGerman = 41,
        //
        // 概要:
        //     German Legal Fanfold (8-1/2 in. x 13 in.)
        xlPaperFanfoldStdGerman = 40,
        //
        // 概要:
        //     U.S. Standard Fanfold (14-7/8 in. x 11 in.)
        xlPaperFanfoldUS = 39,
        //
        // 概要:
        //     Folio (8-1/2 in. x 13 in.)
        xlPaperFolio = 14,
        //
        // 概要:
        //     Ledger (17 in. x 11 in.)
        xlPaperLedger = 4,
        //
        // 概要:
        //     Legal (8-1/2 in. x 14 in.)
        xlPaperLegal = 5,
        //
        // 概要:
        //     Letter (8-1/2 in. x 11 in.)
        xlPaperLetter = 1,
        //
        // 概要:
        //     Letter Small (8-1/2 in. x 11 in.)
        xlPaperLetterSmall = 2,
        //
        // 概要:
        //     Note (8-1/2 in. x 11 in.)
        xlPaperNote = 18,
        //
        // 概要:
        //     Quarto (215 mm x 275 mm)
        xlPaperQuarto = 0xF,
        //
        // 概要:
        //     Statement (5-1/2 in. x 8-1/2 in.)
        xlPaperStatement = 6,
        //
        // 概要:
        //     Tabloid (11 in. x 17 in.)
        xlPaperTabloid = 3,
        //
        // 概要:
        //     User-defined
        xlPaperUser = 0x100
    }

    /// <summary>
    /// XlPaperSizeとNPOI.SS.UserModel.papseSizeの相互変換
    /// </summary>
    internal static class XlPaperSizeParser
    {
        private static readonly Dictionary<XlPaperSize, PaperSize> _Map = new Dictionary<XlPaperSize, PaperSize>()
        {
            { XlPaperSize.xlPaper10x14,             PaperSize.TEN_BY_FOURTEEN },
            { XlPaperSize.xlPaper11x17,             PaperSize.ELEVEN_BY_SEVENTEEN },
            { XlPaperSize.xlPaperA3,                PaperSize.A3 },
            { XlPaperSize.xlPaperA4,                PaperSize.A4 },
            { XlPaperSize.xlPaperA4Small,           PaperSize.A4_Small },
            { XlPaperSize.xlPaperA5,                PaperSize.A5 },
            { XlPaperSize.xlPaperB4,                PaperSize.B4 },
            { XlPaperSize.xlPaperB5,                PaperSize.B5 },
            { XlPaperSize.xlPaperCsheet,            PaperSize.C_Size_Sheet },
            { XlPaperSize.xlPaperDsheet ,           PaperSize.D_Size_Sheet },
            { XlPaperSize.xlPaperEnvelope10 ,       PaperSize.US_Envelope_10 },
            { XlPaperSize.xlPaperEnvelope11,        PaperSize.US_Envelope_11 },
            { XlPaperSize.xlPaperEnvelope12,        PaperSize.US_Envelope_12 },
            { XlPaperSize.xlPaperEnvelope14,        PaperSize.US_Envelope_14 },
            { XlPaperSize.xlPaperEnvelope9,         PaperSize.US_Envelope_9 },
            { XlPaperSize.xlPaperEnvelopeC3,        PaperSize.Envelope_C3 },
            { XlPaperSize.xlPaperEnvelopeC4,        PaperSize.Envelope_C4 },
            { XlPaperSize.xlPaperEnvelopeC5,        PaperSize.Envelope_C5 },
            { XlPaperSize.xlPaperEnvelopeC6,        PaperSize.Envelope_C6 },
            { XlPaperSize.xlPaperEnvelopeDL,        PaperSize.Envelope_DL },
            { XlPaperSize.xlPaperEnvelopeMonarch,   PaperSize.Envelope_MONARCH },
            { XlPaperSize.xlPaperEsheet,            PaperSize.E_Size_Sheet },
            { XlPaperSize.xlPaperExecutive,         PaperSize.US_Executive },
            { XlPaperSize.xlPaperFolio,             PaperSize.Folio },
            { XlPaperSize.xlPaperLedger,            PaperSize.US_Ledger },
            { XlPaperSize.xlPaperLegal,             PaperSize.US_Legal },
            { XlPaperSize.xlPaperLetter,            PaperSize.US_Letter_Small },
            { XlPaperSize.xlPaperLetterSmall,       PaperSize.US_Letter_Small },
            { XlPaperSize.xlPaperNote,              PaperSize.US_Note },
            { XlPaperSize.xlPaperQuarto,            PaperSize.Quarto },
            { XlPaperSize.xlPaperStatement,         PaperSize.US_Statement },
            { XlPaperSize.xlPaperTabloid,           PaperSize.US_Tabloid }
        };
        /// <summary>
        /// XlPaperSize値を指定してPaperSize値を取得。
        /// </summary>
        /// <param name="XlValue">XlPaperSize</param>
        /// <returns>PaperSize値</returns>
        /// <exception cref="SystemException">未定義値検知時</exception>
        public static PaperSize GetPoiValue(XlPaperSize XlValue)
        {
            if (XlValue == 0)
            {
                return OperatorType.IGNORED;
            }
            else if (_Map.ContainsKey(XlValue))
            {
                return _Map[XlValue];
            }
            else
            {
                throw new SystemException("Invalid or unsupported value is spesified as a member of XlPaperSize.");
            }
        }

    }
}






