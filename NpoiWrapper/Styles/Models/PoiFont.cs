using Developers.NpoiWrapper.Configurations.Models;
using Developers.NpoiWrapper.Utils;
using NPOI.SS.UserModel;
using System;
using System.Reflection;

namespace Developers.NpoiWrapper.Styles.Models
{
    internal class PoiFont : IFont
    {
        // IFont interface  is shown below...
        //public interface IFont
        //{
        //    string FontName { get; set; }
        //    double FontHeight { get; set; }
        //    double FontHeightInPoints { get; set; }
        //    bool IsItalic { get; set; }
        //    bool IsStrikeout { get; set; }
        //    short Color { get; set; }
        //    FontSuperScript TypeOffset { get; set; }
        //    FontUnderlineType Underline { get; set; }
        //    short Charset { get; set; }
        //    short Index { get; }
        //    [Obsolete("deprecated POI 3.15 beta 2. Use IsBold instead.")]
        //    short Boldweight { get; set; }
        //    bool IsBold { get; set; }
        //    void CloneStyleFrom(IFont src);
        //}

        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        /// <summary>
        /// フォント
        /// ApplyStyleFrom、Commitでセット
        /// </summary>
        public IFont Font { get; private set; }
        public short InitialIndex { get; private set; } = -1;

        private bool Updated { get; set; } = false;

        [Import(true), Comparison(false), Export(false)]    public short Index { get; private set; }    //set追加

        private String _FontName = string.Empty;
        [Import(true), Comparison(true), Export(true)]      public String FontName
        {
            get { return _FontName;  }
            set
            {
                _FontName = value;
                Updated = true;
            } 
        }
        private double _FontHeight;
        [Import(false), Comparison(false), Export(false)]      public double FontHeight
        {
            get { return _FontHeight; }
            set
            {
                _FontHeight = value;
                Updated = true;
            }
        }
        private double _FontHeightInPoints;
        [Import(true), Comparison(true), Export(true)]      public double FontHeightInPoints
        {
            get { return _FontHeightInPoints; }
            set 
            {
                _FontHeightInPoints = value;
                Updated = true;
            }
        }
        private bool _IsItalic;
        [Import(true), Comparison(true), Export(true)]      public bool IsItalic
        {
            get { return _IsItalic; }
            set
            {
                _IsItalic = value;
                Updated = true;
            }
        }
        private bool _IsStrikeout;
        [Import(true), Comparison(true), Export(true)]      public bool IsStrikeout
        {
            get { return _IsStrikeout; }
            set
            {
                _IsStrikeout = value;
                Updated = true;
            }
        }
        private short _Color;
        [Import(true), Comparison(true), Export(true)]      public short Color
        {
            get { return _Color; }
            set
            {
                _Color = value;
                Updated = true;
            }
        }
        private FontSuperScript _TypeOffset;
        [Import(true), Comparison(true), Export(true)]      public FontSuperScript TypeOffset
        {
            get { return _TypeOffset; }
            set
            {
                _TypeOffset = value;
                Updated = true;
            }
        }
        private FontUnderlineType _Underline;
        [Import(true), Comparison(true), Export(true)]      public FontUnderlineType Underline
        {
            get { return _Underline; }
            set
            {
                _Underline = value;
                Updated = true;
            }
        }
        private short _Charset;
        [Import(true), Comparison(true), Export(true)]      public short Charset
        {
            get { return _Charset; }
            set
            {
                _Charset = value;
                Updated = true;
            }
        }
        //deprecated POI 3.15 beta 2. Use IsBold instead.
        private short _Boldweight;
        [Import(true), Comparison(true), Export(true)]      public short Boldweight
        {
            get { return _Boldweight; }
            set
            {
                _Boldweight = value;
                Updated = true;
            }
        }
        private bool _IsBold;
        [Import(true), Comparison(true), Export(true)]      public bool IsBold
        {
            get { return _IsBold; }
            set
            {
                _IsBold = value;
                Updated = true;
            }
        }

        public void CloneStyleFrom(IFont src)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Workbook
        /// </summary>
        private IWorkbook PoiBook { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="PoiBook">IWorkbookインスタンス</param>
        /// <param name="FontIndex">編集のベースとするフォントのIndex</param>
        public PoiFont(IWorkbook PoiBook, short FontIndex)
        {
            this.PoiBook = PoiBook;
            //指定されたフォントをインポート
            ApplyStyleFrom(this.PoiBook.GetFontAt(FontIndex));
        }

        /// <summary>
        /// 指定されたフォント定義を自プロパティにインポート
        /// </summary>
        /// <param name="Font">インポート対象フォント定義</param>
        public void ApplyStyleFrom(IFont Font)
        {
            //基点Fontオブジェクトを保存
            this.Font = Font;
            this.InitialIndex = this.Font.Index;
            //指定されたCurrentFontのプロパティを一覧
            System.Reflection.PropertyInfo[] PropList = Font.GetType().GetProperties();
            //指定されたCurrentStyleのプロパティ値を自プロパティにセット
            foreach (System.Reflection.PropertyInfo source in PropList)
            {
                //無条件に全てコピーする
                System.Reflection.PropertyInfo destination = this.GetType().GetProperty(source.Name);
                if (destination != null)
                {
                    //インポート属性のあるプロパティのみインポート
                    ImportAttribute atr = AttributeUtility.GetPropertyAttribute<ImportAttribute>(this, destination.Name);
                    if (atr != null && atr.Import)
                    {
                        destination.SetValue(this, source.GetValue(Font));
                    }
                }
            }
            //インポート直後のまっさらな状態
            this.Updated = false;
        }

        /// <summary>
        /// 変更を確定する
        /// </summary>
        public short Commit()
        {
            //上位からの変更があった場合
            if (this.Updated)
            {
                //Indexをクリア
                this.Index = -1;
                //マスター上に同一フォント情報があるかチェックし、あれば採用
                for (short FIdx = 0; FIdx < PoiBook.NumberOfFonts; FIdx++)
                {
                    PoiFont MasterFont = new PoiFont(PoiBook, FIdx);
                    if (this.StyleEquals(PoiBook.GetFontAt(FIdx)))
                    {
                        this.Font = MasterFont;
                        this.Index = MasterFont.Index;
                        Logger.Debug("CurrentFont:[Index:" + this.InitialIndex + "] => Font.[Index:" + this.Index + "] is picked up from Font-Master in this Book..");
                        break;
                    }
                }
                //マスター上になかったら新規に作成(再利用可能のチェックはしていないので、追加のみ)
                if (this.Index == -1)
                {
                    this.Font = PoiBook.CreateFont();
                    this.Index = Font.Index;
                    ApplyStyleTo(this.Font);
                    Logger.Debug("CurrentFont:[Index:" + this.InitialIndex + "] => Font.[Index:" + this.Index + "] is newly created by IWorkbook.CreateFont().");
                }
            }
            //上位からの変更がなかった場合は何もしない。
            else
            {
                Logger.Debug("CurrentFont:[Index:" + this.InitialIndex + "] == Font.[Index:" + this.Index + "] No propertiy was updated.");
            }
            //InitialIndexの更新
            this.InitialIndex = this.Font.Index;
            //Indexを返す(最新のIndex)
            return this.Font.Index;
        }

        /// <summary>
        /// 指定フォント定義がこのフォント定義と同じフォント定義かどうかを判断する
        /// </summary>
        /// <param name="TagetFontStyle">対象フォント定義</param>
        /// <returns></returns>
        public bool StyleEquals(IFont Font)
        {
            bool RetVal = true;
            //nullでなくIFontであること
            if ((Font != null) && Font is IFont)
            {
                //指定されたCellStyleのプロパティを一覧
                PropertyInfo[] TargetProps = Font.GetType().GetProperties();
                //指定されたCellStyleのプロパティ分ループ
                foreach (PropertyInfo TargetProp in TargetProps)
                {
                    //同名の自プロパティ情報取得
                    PropertyInfo MyProp = this.GetType().GetProperty(TargetProp.Name);
                    //IFontプロパティ以外は無視
                    if (MyProp != null)
                    {
                        //コンペア属性のあるプロパティのみ比較
                        ComparisonAttribute Attr = AttributeUtility.GetPropertyAttribute<ComparisonAttribute>(this, MyProp.Name);
                        if (Attr != null && Attr.Compare)
                        {
                            //プロパティ値不一致
                            if (!MyProp.GetValue(this).Equals(TargetProp.GetValue(Font)))
                            {
                                //Logger.Debug("this." + MyProp.Name + "[" + MyProp.GetValue(this) + "] != Target[" + TargetProp.GetValue(Font) + "]");
                                RetVal = false;
                                break;
                            }
                        }
                    }
                }
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたフォント定義に自プロパティをエクスポート
        /// </summary>
        /// <param name="Font">エクスポート対象フォント定義</param>
        public void ApplyStyleTo(IFont Font)
        {
            //プロパティリスト取得
            PropertyInfo[] Props = this.GetType().GetProperties();
            //指定されたCurrentStyleのプロパティ値を自プロパティにセット
            foreach (PropertyInfo source in Props)
            {
                //エクスポート属性のあるプロパティのみエクスポート
                ExportAttribute atr = AttributeUtility.GetPropertyAttribute<ExportAttribute>(this, source.Name);
                if (atr != null && atr.Export)
                {
                    //プロパティ情報を取得しプロパティに値をセット
                    //ただしICellStyeから拡張された独自プロパティは無視
                    PropertyInfo destination = Font.GetType().GetProperty(source.Name);
                    destination?.SetValue(Font, source.GetValue(this));
                }
            }
        }
    }
}
