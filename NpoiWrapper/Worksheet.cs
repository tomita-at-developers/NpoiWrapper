﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Worksheetクラス
    /// Microsoft.Office.Interop.Excel.Workbookをエミュレート
    /// WorkbookクラスのActiveSheet、SheetsクラスのAddでのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Worksheet
    {
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        public Cells Cells { get; } 
        public Range Range { get; }
        internal Workbook ParentBook { get; private set; }
        internal ISheet PoiSheet { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentWorkbook">親ブックのWorkbookクラスインスタンス</param>
        /// <param name="Sheet">自ISheet</param>
        internal Worksheet(Workbook ParentWorkbook, ISheet Sheet)
        {
            Logger.Debug("SheetName[" + Sheet.SheetName + "]");

            //親クラスの保存
            this.ParentBook = ParentWorkbook;
            //POIシートの保存
            PoiSheet = Sheet;
            //Cells, Rangeの初期値をセット
            //インデクスを省略した場合はこの値が取得される(シート全域)
            Cells = new Cells(this, new CellRangeAddressList(-1, -1, -1, -1));
            Range = new Range(this, new CellRangeAddressList(-1, -1, -1, -1));
        }

        /// <summary>
        /// シート名
        /// </summary>
        public string Name
        {
            get
            {
                return PoiSheet.SheetName;
            }
            set
            {
                ParentBook.PoiBook.SetSheetName(
                    ParentBook.PoiBook.GetSheetIndex(PoiSheet.SheetName), value);
            }
        }

        /// <summary>
        /// シートIndex
        /// </summary>
        public int Index
        {
            get
            {
                return ParentBook.PoiBook.GetSheetIndex(PoiSheet.SheetName);
            }
        }

        /// <summary>
        /// シートのコピー
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public Worksheet Copy(string SheetName)
        {
            ISheet Sheet = PoiSheet.CopySheet(SheetName);
            return new Worksheet(ParentBook, Sheet);
        }

        /// <summary>
        /// シートの削除
        /// </summary>
        public void Delete()
        {
            int SheetIndex = ParentBook.PoiBook.GetSheetIndex(PoiSheet.SheetName);
            ParentBook.PoiBook.RemoveSheetAt(SheetIndex);
        }

        /// <summary>
        /// ページ設定
        /// </summary>
        /// <param name="StyleName">NpoiWrapper.configのページ設定パターン名(省略時"default")</param>
        /// <param name="HeaderLeft">ヘッダー左文字列</param>
        /// <param name="HeaderCenter">ヘッダー中央文字列</param>
        /// <param name="HeaderRight">ヘッダー右文字列</param>
        /// <param name="FooterLeft">フッター左文字列</param>
        /// <param name="FooterCenter">フッター中央文字列</param>
        /// <param name="FooterRight">フッター右文字列</param>
        public void PageSetup(
            string StyleName = "default",
            string HeaderLeft = "",
            string HeaderCenter = "",
            string HeaderRight = "",
            string FooterLeft = "",
            string FooterCenter = "",
            string FooterRight = "")
        {
            //指定された名前の設定が存在すればそれを適用
            if (ParentBook.PageSetups.ContainsKey(StyleName))
            {
                Configurations.Models.PageSetup Setup = ParentBook.PageSetups[StyleName];
                PoiSheet.PrintSetup.Landscape = Setup.Paper.Landscape;
                PoiSheet.PrintSetup.PaperSize = (short)Setup.Paper.Size;
                PoiSheet.PrintSetup.HeaderMargin = Setup.Margins.Header.ValueInInch;
                PoiSheet.PrintSetup.FooterMargin = Setup.Margins.Footer.ValueInInch;
                if (Setup.Scaling.Fit != null)
                {
                    PoiSheet.FitToPage = true;
                    PoiSheet.PrintSetup.FitWidth = Setup.Scaling.Fit.Wide;
                    PoiSheet.PrintSetup.FitHeight = Setup.Scaling.Fit.Tall;
                }
                else if (Setup.Scaling.Adjust != null)
                {
                    PoiSheet.FitToPage = false;
                    PoiSheet.PrintSetup.Scale = Setup.Scaling.Adjust.Scale;
                }
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.TopMargin, Setup.Margins.Body.TopInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.RightMargin, Setup.Margins.Body.RightInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.BottomMargin, Setup.Margins.Body.BottomInInch);
                PoiSheet.SetMargin(NPOI.SS.UserModel.MarginType.LeftMargin, Setup.Margins.Body.LeftInInch);
                PoiSheet.HorizontallyCenter = Setup.Center.Horizontally;
                PoiSheet.VerticallyCenter = Setup.Center.Vertically;
                if (Setup.Titles.Row.Length > 0)
                {
                    PoiSheet.RepeatingRows = CellRangeAddress.ValueOf(Setup.Titles.Row);
                }
                if (Setup.Titles.Column.Length > 0)
                {
                    PoiSheet.RepeatingColumns = CellRangeAddress.ValueOf(Setup.Titles.Column);
                }
            }
            //ヘッダー/フッターの文字は引数から適用
            PoiSheet.Header.Left = HeaderLeft;
            PoiSheet.Header.Center = HeaderCenter;
            PoiSheet.Header.Right = HeaderRight;
            PoiSheet.Footer.Left = FooterLeft;
            PoiSheet.Footer.Center = FooterCenter;
            PoiSheet.Footer.Right = FooterRight;
        }

        /// <summary>
        /// シートの保護
        /// HSSFではこの操作によりオートフィルターが無効化される。
        /// </summary>
        /// <param name="Password"></param>
        public void Protect(string Password = "")
        {
            PoiSheet.ProtectSheet(Password);
            //XSSFならロック解除できるのでやっておく
            if (PoiSheet is XSSFSheet xssfSheet)
            {
                xssfSheet.LockAutoFilter(false);
                xssfSheet.LockSort(false);
            }
            else
            {
                //為す術なし
            }
        }

        /// <summary>
        /// 先頭行/先頭列の固定
        /// </summary>
        /// <param name="TopLeftCell">固定位置(A1形式)</param>
        public void CreateFreezePane(string TopLeftCell)
        {
            PoiSheet.CreateFreezePane(
                CellRangeAddress.ValueOf(TopLeftCell).FirstColumn,
                CellRangeAddress.ValueOf(TopLeftCell).FirstRow);
        }

        /// <summary>
        /// オートフィルターの設定
        /// HSSFではProtectをコールするとオートフィルタが無効化される。
        /// </summary>
        /// <param name="FilterRange"></param>
        public void AutoFilter(string FilterRange)
        {
            PoiSheet.SetAutoFilter(CellRangeAddress.ValueOf(FilterRange));
        }
    }
}
