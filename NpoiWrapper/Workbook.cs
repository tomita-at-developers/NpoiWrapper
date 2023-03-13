using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
//using Developers.NpoiWrapper.Configurations;

namespace Developers.NpoiWrapper
{
    /// <summary>
    /// Workbookクラス
    /// Microsoft.Office.Interop.Excel.Workbookをエミュレート
    /// WorkbooksクラスのAdd、Openでコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Workbook
    {
        internal IWorkbook PoiBook {get; private set; }
        public string FileName { get; private set; } = string.Empty;
        internal Dictionary<string, ICellStyle> CellStyles { get; private set; } = new Dictionary<string, ICellStyle>();
        internal Dictionary<string, Configurations.Models.PageSetup> PageSetups { get; private set; } = new Dictionary<string, Configurations.Models.PageSetup>();
        internal IFont Font { get; private set; } = null;
        public Sheets Worksheets { get; private set; }
        public Sheets Sheets { get; private set; }

        /// <summary>
        /// 新規ファイルを作成する
        /// </summary>
        internal Workbook() : this(false)
        {
        }

        /// <summary>
        /// 新規ファイルを作成する
        /// </summary>
        internal Workbook(bool Excel97_2003)
        {
            //Excel形式判断
            if (Excel97_2003)
            {
                //Excel97-2003
                PoiBook = new HSSFWorkbook();
            }
            else
            {
                //Excel2007以降
                PoiBook = new XSSFWorkbook();
            }
            //設定ファイルの内容を反映
            ApplyConfigs();
            //新規の場合はシートを一つ追加しておく
            PoiBook.CreateSheet();
            //この時点でファイル名は未定義
            FileName = string.Empty;
            //Sheets,Worksheetsの初期化
            Sheets = new Sheets(this);
            Worksheets = new Worksheets(this);
        }

        /// <summary>
        /// 既存ファイルを開く(読取専用のみサポート)
        /// </summary>
        /// <param name="FileName"></param>
        internal Workbook(string FileName)
        {
            ////ファイルを開く
            FileStream Stream = new FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            PoiBook = WorkbookFactory.Create(Stream);
            //ファイル名を保存
            this.FileName = FileName;
            //Sheets,Worksheetsの初期化
            Sheets = new Sheets(this);
            Worksheets = new Worksheets(this);
        }

        /// <summary>
        /// アクティブシートの取得
        /// ★InteropでもWorkbookクラスにある
        /// </summary>
        /// <returns>ISheetインターフェイス</returns>
        public Worksheet ActiveSheet()
        {
            return new Worksheet(this, PoiBook.GetSheetAt(PoiBook.ActiveSheetIndex));
        }

        ///// <summary>
        ///// 開いているファイルの保存
        ///// </summary>
        //public void Save()
        //{
        //    //ファイル保存
        //    using (FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write))
        //    {
        //        PoiBook.Write(fs, false);
        //    }
        //}

        /// <summary>
        /// 名前を付けて保存(保存後は閉じる)
        /// </summary>
        /// <param name="Path">フルパスファイル名</param>
        public void SaveAs(string Path)
        {
            //保存ファイル名生成＆拡張子の取得
            FileName = Path;
            string CurrentExtention = System.IO.Path.GetExtension(FileName);
            //現在のWorkbookの既定拡張子を取得
            string DefaultExtension = PoiBook.SpreadsheetVersion.DefaultExtension;
            //拡張子が既定拡張子と異なる場合は既定拡張子を適用
            if (String.Compare(CurrentExtention, DefaultExtension, true) != 0)
            {
                FileName = System.IO.Path.ChangeExtension(FileName, DefaultExtension);
            }
            //ファイル保存
            using (FileStream Stream = new FileStream(FileName, FileMode.Create, FileAccess.Write))
            {
                PoiBook.Write(Stream, false);
            }
        }

        /// <summary>
        /// 設定ファイル(NpoiWrapper.config)の設定を反映
        /// </summary>
        private void ApplyConfigs()
        {
            //設定ファイル(NpoiWrapper.config)の読込
            Configurations.ConfigurationManager CfgManager = new Configurations.ConfigurationManager();
            //フォントの生成
            if (CfgManager.Configs.Font != null)
            {
                Font = PoiBook.CreateFont();
                Font.FontName = CfgManager.Configs.Font.Name;
                Font.FontHeightInPoints = CfgManager.Configs.Font.Size;
            }
            //ページ設定リストの生成
            foreach (Configurations.Models.PageSetup ps in CfgManager.Configs.PageSetup ?? new List<Configurations.Models.PageSetup>())
            {
                PageSetups.Add(ps.Name, ps);
            }
            //セルスタイルリストの生成
            foreach (Configurations.Models.CellStyle cs in CfgManager.Configs.CellStyle ?? new List<Configurations.Models.CellStyle>())
            {
                ICellStyle pcs = PoiBook.CreateCellStyle();
                pcs.BorderTop = cs.Border.Top;
                pcs.BorderRight = cs.Border.Right;
                pcs.BorderBottom = cs.Border.Bottom;
                pcs.BorderLeft = cs.Border.Left;
                pcs.Alignment = cs.Align.Horizontal;
                pcs.VerticalAlignment = cs.Align.Vertical;
                pcs.WrapText = cs.WrapText.Value;
                pcs.IsLocked = cs.IsLocked.Value;
                if (cs.DataFormat.Value.Length > 0)
                {
                    pcs.DataFormat = PoiBook.CreateDataFormat().GetFormat(cs.DataFormat.Value);
                }
                pcs.FillForegroundColor = cs.Fill.Color;
                pcs.FillPattern = cs.Fill.Pattern;
                if (Font != null)
                {
                    pcs.SetFont(Font);
                }
                CellStyles.Add(cs.Name, pcs);
            }
        }
    }
}
