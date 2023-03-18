using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //public interface Window
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    Range ActiveCell { get; }
    //    Chart ActiveChart { get; }
    //    Pane ActivePane { get; }
    //    object ActiveSheet { get; }
    //    object Caption { get; set; }
    //    bool DisplayFormulas { get; set; }
    //    bool DisplayGridlines { get; set; }
    //    bool DisplayHeadings { get; set; }
    //    bool DisplayHorizontalScrollBar { get; set; }
    //    bool DisplayOutline { get; set; }
    //    bool _DisplayRightToLeft { get; set; }
    //    bool DisplayVerticalScrollBar { get; set; }
    //    bool DisplayWorkbookTabs { get; set; }
    //    bool DisplayZeros { get; set; }
    //    bool EnableResize { get; set; }
    //    bool FreezePanes { get; set; }
    //    int GridlineColor { get; set; }
    //    XlColorIndex GridlineColorIndex { get; set; }
    //    double Height { get; set; }
    //    int Index { get; }
    //    double Left { get; set; }
    //    string OnWindow { get; set; }
    //    Panes Panes { get; }
    //    Range RangeSelection { get; }
    //    int ScrollColumn { get; set; }
    //    int ScrollRow { get; set; }
    //    Sheets SelectedSheets { get; }
    //    object Selection { get; }
    //    bool Split { get; set; }
    //    int SplitColumn { get; set; }
    //    double SplitHorizontal { get; set; }
    //    int SplitRow { get; set; }
    //    double SplitVertical { get; set; }
    //    double TabRatio { get; set; }
    //    double Top { get; set; }
    //    XlWindowType Type { get; }
    //    double UsableHeight { get; }
    //    double UsableWidth { get; }
    //    bool Visible { get; set; }
    //    Range VisibleRange { get; }
    //    double Width { get; set; }
    //    int WindowNumber { get; }
    //    XlWindowState WindowState { get; set; }
    //    object Zoom { get; set; }
    //    XlWindowView View { get; set; }
    //    bool DisplayRightToLeft { get; set; }
    //    SheetViews SheetViews { get; }
    //    object ActiveSheetView { get; }
    //    bool DisplayRuler { get; set; }
    //    bool AutoFilterDateGrouping { get; set; }
    //    bool DisplayWhitespace { get; set; }
    //    object Activate();
    //    object ActivateNext();
    //    object ActivatePrevious();
    //    bool Close([Optional] object SaveChanges, [Optional] object Filename, [Optional] object RouteWorkbook);
    //    object LargeScroll([Optional] object Down, [Optional] object Up, [Optional] object ToRight, [Optional] object ToLeft);
    //    Window NewWindow();
    //    object _PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //    object PrintPreview([Optional] object EnableChanges);
    //    object ScrollWorkbookTabs([Optional] object Sheets, [Optional] object Position);
    //    object SmallScroll([Optional] object Down, [Optional] object Up, [Optional] object ToRight, [Optional] object ToLeft);
    //    int PointsToScreenPixelsX(int Points);
    //    int PointsToScreenPixelsY(int Points);
    //    object RangeFromPoint(int x, int y);
    //    void ScrollIntoView(int Left, int Top, int Width, int Height, [Optional] object Start);
    //    object PrintOut([Optional] object From, [Optional] object To, [Optional] object Copies, [Optional] object Preview, [Optional] object ActivePrinter, [Optional] object PrintToFile, [Optional] object Collate, [Optional] object PrToFileName);
    //}
    public class Window
    {
        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Workbook Parent { get; } // 調べてみたら親はBookだった

        public string Caption { get; }
        public Worksheet ActiveSheet { get; internal set; }
        internal Range RangeSelection { get; set; }

        internal Window(Workbook ParentBook)
        {
            this.Parent = ParentBook;
            this.Caption = this.Parent.Name;
            this.ActiveSheet = this.Parent.ActiveSheet;
            this.RangeSelection = null;
        }

        /// <summary>
        /// 表示枠の固定
        /// </summary>
        bool FreezePanes
        {
            get
            {
                ISheet Sheet;
                //ActiveSheetの明示的な設定がある場合はそのSheetを参照
                if (this.ActiveSheet != null)
                {
                    Sheet = this.ActiveSheet.PoiSheet;
                }
                //SelectedRangeがあればその親Sheetを参照
                else if (this.RangeSelection != null)
                {
                    Sheet = this.RangeSelection.Parent.PoiSheet;
                }
                //いずれもなければ親BookのActiveSheetを対象とする
                else
                {
                    Sheet = this.Parent.PoiBook.GetSheetAt(this.Parent.PoiBook.ActiveSheetIndex);
                }
                return Sheet.PaneInformation.IsFreezePane();
            }
            set
            {
                //boolなら処理する(bool以外は無視)
                if (value is bool Freeze) 
                {
                    //設定の時
                    if (Freeze)
                    {
                        //SelectedRangeがあれば処理(なければ無視)
                        //もしアドレスが複数あっても最初のアドレスのみ参照してPaneを設定
                        this.RangeSelection?.Parent.PoiSheet.CreateFreezePane(
                            this.RangeSelection.SafeAddressList.GetCellRangeAddress(0).FirstColumn,
                            this.RangeSelection.SafeAddressList.GetCellRangeAddress(0).FirstRow);
                    }
                    //解除の時
                    else
                    {
                        ISheet Sheet;
                        //ActiveSheetの明示的な設定がある場合はそのSheetを参照
                        if (this.ActiveSheet != null)
                        {
                            Sheet = this.ActiveSheet.PoiSheet;
                        }
                        //SelectedRangeがあればその親Sheetを参照
                        else if (this.RangeSelection != null)
                        {
                            Sheet = this.RangeSelection.Parent.PoiSheet;
                        }
                        //いずれもなければ親BookのActiveSheetを対象とする
                        else
                        {
                            Sheet = this.Parent.PoiBook.GetSheetAt(this.Parent.PoiBook.ActiveSheetIndex);
                        }
                        //FreezePaneの解除
                        Sheet.CreateFreezePane(0, 0);
                    }
                }
            }
        }

    }
}
