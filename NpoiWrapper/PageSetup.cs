using Developers.NpoiWrapper.Model;
using Developers.NpoiWrapper.Utils;
using log4net.Repository.Hierarchy;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // PageSetup interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface PageSetup
    //{
    // +  Application Application { get; }
    // +  XlCreator Creator { get; }
    // +  object Parent { get; }
    //    bool BlackAndWhite { set; }
    // +  double BottomMargin { get; set; }
    // +  string CenterFooter { get; set; }
    // +  string CenterHeader { get; set; }
    // +  bool CenterHorizontally { get; set; }
    // +  bool CenterVertically { get; set; }
    //    XlObjectSize ChartSize { get; set; }
    //    bool Draft { get; set; }
    //    int FirstPageNumber { get; set; }
    // +  object FitToPagesTall { get; set; }
    // +  object FitToPagesWide { get; set; }
    // +  double FooterMargin { get; set; }
    // +  double HeaderMargin { get; set; }
    // +  string LeftFooter { get; set; }
    // +  string LeftHeader { get; set; }
    // +  double LeftMargin { get; set; }
    //    XlOrder Order { get; set; }
    // +  XlPageOrientation Orientation { get; set; }
    // +  XlPaperSize PaperSize { get; set; }
    //    string PrintArea { get; set; }
    //    bool PrintGridlines { get; set; }
    //    bool PrintHeadings { get; set; }
    //    bool PrintNotes { get; set; }
    //    object PrintQuality { get; set; }
    // +  string PrintTitleColumns { get; set; }
    // +  string PrintTitleRows { get; set; }
    // +  string RightFooter { get; set; }
    // +  string RightHeader { get; set; }
    // +  double RightMargin { get; set; }
    // +  double TopMargin { get; set; }
    // +  object Zoom { get; set; }
    //    XlPrintLocation PrintComments { get; set; }
    //    XlPrintErrors PrintErrors { get; set; }
    //    Graphic CenterHeaderPicture { get; }
    //    Graphic CenterFooterPicture { get; }
    //    Graphic LeftHeaderPicture { get; }
    //    Graphic LeftFooterPicture { get; }
    //    Graphic RightHeaderPicture { get; }
    //    Graphic RightFooterPicture { get; }
    //}

    public class PageSetup
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// 以下の情報によれば、NPOIの不具合により＋１しないと正しく動作しないとのこと。<br/>
        /// https://stackoverflow.com/questions/57486314/c-sharp-application-to-print-excel-npoi-print-setup-to-set-to-legal-pager-size 
        /// </summary>
        private const short _PAPER_SIZE_ADJUSTER = 1;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンスラクタ
        /// </summary>
        /// <param name="ParentSheet">親シートクラス</param>
        internal PageSetup(Worksheet ParentSheet)
        {
            Logger.Debug("ShhetName[" + ParentSheet.Name + "]");
            this.Parent = ParentSheet;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Worksheet Parent { get; }

        /// <summary>
        /// 下余白(ポイント単位)
        /// </summary>
        public double BottomMargin 
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.BottomMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.BottomMargin, Parent.Application.PointsToInches(value));
            }
        }
        public string CenterFooter
        {
            get
            {
                return Parent.PoiSheet.Footer.Center;
            }
            set
            {
                Parent.PoiSheet.Footer.Center = value;
            }
        }
        public string CenterHeader
        {
            get
            {
                return Parent.PoiSheet.Header.Center;
            }
            set
            {
                Parent.PoiSheet.Header.Center = value;
            }
        }
        public bool CenterHorizontally
        {
            get
            {
                return Parent.PoiSheet.HorizontallyCenter;
            }
            set
            {
                Parent.PoiSheet.HorizontallyCenter = value;
            }
        }
        public bool CenterVertically
        {
            get
            {
                return Parent.PoiSheet.VerticallyCenter;
            }
            set
            {
                Parent.PoiSheet.VerticallyCenter = value;
            }
        }
        /// <summary>
        /// 「次のページ数に合わせて印刷」の縦分割数
        /// 指定しない場合はFalse
        /// </summary>
        public object FitToPagesTall
        {
            get
            {
                object RetVal;
                //ゼロの場合はFalseを返す
                if (Parent.PoiSheet.PrintSetup.FitHeight == 0)
                {
                    RetVal = false;
                }
                //ゼロ以外はその数字(ページ数)を返す
                else
                {
                    RetVal = Parent.PoiSheet.PrintSetup.FitHeight;
                }
                return RetVal;
            }
            set
            {
                //boolの場合
                if (value is bool)
                {
                    //falseの場合はゼロ
                    if ((bool)value == false)
                    {
                        Parent.PoiSheet.PrintSetup.FitHeight = 0;
                    }
                    //Trueの場合
                    else
                    {
                        throw new ArgumentNullException("FitToPagesTall");
                    }
                }
                //分割ページ数としてセット
                else
                {
                    try
                    {
                        Parent.PoiSheet.PrintSetup.FitHeight = Convert.ToInt16(value);
                    }
                    catch
                    {
                        throw new ArgumentNullException("FitToPagesTall");
                    }
                }
            }
        }
        /// <summary>
        /// 「次のページ数に合わせて印刷」の横分割数
        /// 指定しない場合はFalse
        /// </summary>
        public object FitToPagesWide
        {
            get
            {
                object RetVal;
                //ゼロの場合はFalseを返す
                if (Parent.PoiSheet.PrintSetup.FitWidth == 0)
                {
                    RetVal = false;
                }
                //ゼロ以外はその数字(ページ数)を返す
                else
                {
                    RetVal = Parent.PoiSheet.PrintSetup.FitWidth;
                }
                return RetVal;
            }
            set
            {
                //boolの場合
                if (value is bool)
                {
                    //falseの場合はゼロ
                    if ((bool)value == false)
                    {
                        Parent.PoiSheet.PrintSetup.FitWidth = 0;
                    }
                    //Trueは認めない
                    else
                    {
                        throw new ArgumentNullException("FitToPagesWide");
                    }
                }
                //分割ページ数としてセット
                else
                {
                    try
                    {
                        Parent.PoiSheet.PrintSetup.FitWidth = Convert.ToInt16(value);
                    }
                    catch
                    {
                        throw new ArgumentNullException("FitToPagesWide");
                    }
                }
            }
        }
        /// <summary>
        /// ページの下部からフッターまでの距離(ポイント単位)
        /// </summary>
        public double FooterMargin
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.FooterMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.FooterMargin, Parent.Application.PointsToInches(value));
            }
        }
        /// <summary>
        /// ページの上部からヘッダーまでの距離(ポイント単位)
        /// </summary>
        public double HeaderMargin
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.HeaderMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.HeaderMargin, Parent.Application.PointsToInches(value));
            }
        }
        public string LeftFooter
        {
            get
            {
                return Parent.PoiSheet.Footer.Left;
            }
            set
            {
                Parent.PoiSheet.Footer.Left = value;
            }
        }
        public string LeftHeader
        {
            get
            {
                return Parent.PoiSheet.Header.Left;
            }
            set
            {
                Parent.PoiSheet.Header.Left = value;
            }
        }
        /// <summary>
        /// 左余白(ポイント単位)
        /// </summary>
        public double LeftMargin
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.LeftMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.LeftMargin, Parent.Application.PointsToInches(value));
            }
        }
        public XlPageOrientation Orientation
        {
            get
            {
                XlPageOrientation RetVal = XlPageOrientation.xlPortrait;
                if (Parent.PoiSheet.PrintSetup.Landscape)
                {
                    RetVal = XlPageOrientation.xlLandscape;
                }
                return RetVal;
            }
            set
            {
                bool SetValue = false;
                if (value == XlPageOrientation.xlLandscape)
                {
                    SetValue = true;
                }
                Parent.PoiSheet.PrintSetup.Landscape = SetValue;
            }
        }
        public XlPaperSize PaperSize
        {
            get
            {
                XlPaperSize RetVal = XlPaperSize.xlPaperUser;
                if (PaperSizeParser.Xls.TryParse((short)(Parent.PoiSheet.PrintSetup.PaperSize - _PAPER_SIZE_ADJUSTER), out XlPaperSize XlValue))
                {
                    RetVal = XlValue;
                }
                return RetVal;
            }
            set
            {
                short SetValue = (short)NPOI.SS.UserModel.PaperSize.PRINTER_DEFAULT_PAPERSIZE;
                if (PaperSizeParser.Poi.TryParse(value, out short PoiValue))
                {
                    SetValue = PoiValue;
                }
                Parent.PoiSheet.PrintSetup.PaperSize = (short)(SetValue + _PAPER_SIZE_ADJUSTER);
            }
        }
        public string PrintTitleColumns
        {
            get
            {
                CellRangeAddress Adr = Parent.PoiSheet.RepeatingColumns;
                return Adr.FormatAsString();
            }
            set
            {
                Parent.PoiSheet.RepeatingColumns = CellRangeAddress.ValueOf(value);
            }
        }
        public string PrintTitleRows
        {
            get
            {
                CellRangeAddress Adr = Parent.PoiSheet.RepeatingRows;
                return Adr.FormatAsString();
            }
            set
            {
                Parent.PoiSheet.RepeatingRows = CellRangeAddress.ValueOf(value);
            }
        }
        public string RightFooter
        {
            get
            {
                return Parent.PoiSheet.Footer.Right;
            }
            set
            {
                Parent.PoiSheet.Footer.Right = value;
            }
        }
        public string RightHeader
        {
            get
            {
                return Parent.PoiSheet.Header.Right;
            }
            set
            {
                Parent.PoiSheet.Header.Right = value;
            }
        }
        /// <summary>
        /// 右余白(ポイント単位)
        /// </summary>
        public double RightMargin
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.RightMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.RightMargin, Parent.Application.PointsToInches(value));
            }
        }
        /// <summary>
        /// 上余白(ポイント単位)
        /// </summary>
        public double TopMargin
        {
            get
            {
                //インチをポイントに変換
                return Parent.Application.InchesToPoints(Parent.PoiSheet.GetMargin(MarginType.TopMargin));
            }
            set
            {
                //ポイントをインチに変換
                Parent.PoiSheet.SetMargin(MarginType.TopMargin, Parent.Application.PointsToInches(value));
            }
        }
        /// <summary>
        /// 「拡大/縮小」時のパーセンテージ
        /// 「次のページ数に合わせて印刷」の場合はFalse
        /// </summary>
        public object Zoom
        {
            get
            {
                object RetVal;
                //「次のページ数に合わせて印刷」ならFalse
                if (Parent.PoiSheet.FitToPage)
                {
                    RetVal = false;
                }
                //「次のページ数に合わせて印刷」でなければ拡大/縮小パーセンテージ
                else
                {
                    RetVal = Parent.PoiSheet.PrintSetup.Scale;
                }
                return null;
            }
            set
            {
                //boolの場合
                if(value is bool)
                {
                    //Falseの場合
                    if((bool)value == false)
                    {
                        //「次のページ数に合わせて印刷」が選択されたとみなす。
                        // なのでFitToPageをTrueにする。
                        Parent.PoiSheet.FitToPage = true;
                    }
                    //Trueは認めない
                    else
                    {
                        throw new ArgumentNullException("Zoom");
                    }
                }
                //boolでない場合(パーセンテージ)
                else
                {
                    //「次のページ数に合わせて印刷」を無効に(つまり「拡大/縮小」)
                    Parent.PoiSheet.FitToPage = false;
                    //拡大/縮小のパーセンテージとしてセット
                    try
                    {
                        Parent.PoiSheet.PrintSetup.Scale = Convert.ToInt16(value);
                    }
                    catch
                    {
                        throw new ArgumentNullException("Zoom");
                    }
                }
            }
        }

        #endregion

        #endregion

    }
}
