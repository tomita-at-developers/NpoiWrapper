using Developers.NpoiWrapper.Utils;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Developers.NpoiWrapper.Model.Wrapper
{
    internal class PoiCellStyle : ICellStyle
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        private static Dictionary<string, PropertyInfo> _ImportMap = null;
        private static Dictionary<string, PropertyInfo> _CompareMap = null;
        private static Dictionary<string, PropertyInfo> _ExportMap = null;
        private static Dictionary<string, PropertyInfo> _PoiCellStyleMap = null;
        private static Dictionary<string, PropertyInfo> _ICellStyleMap = null;
        private static PropertyInfo _FillBackgroundColorColorProp = null;

        /// <summary>
        /// 書式文字列
        /// </summary>
        public string _DataFormatString = string.Empty;

        #endregion

        #region "construcrtors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentSheet">Worksheetインスタンス</param>
        /// <param name="StyleIndex">スタイルIndex</param>
        public PoiCellStyle(ISheet PoiSheet, short StyleIndex)
        {
            this.PoiSheet = PoiSheet;
            //プロパティマップの生成
            if(_PoiCellStyleMap == null)
            {
                CreateMaps();
            }
            //指定されたスタイルをインポート
            this.CellStyle = this.PoiBook.GetCellStyleAt(StyleIndex);
            ImportFrom(this.CellStyle);
        }

        #endregion

        #region "interface implementations"

        #region "mandatory properties"

        [Import(true), Comparison(false), Export(false)] public short Index { get; /*追加*/ private set; }
        [Import(true), Comparison(true), Export(true)] public bool ShrinkToFit { get; set; }
        [Import(true), Comparison(true), Export(true)] public short DataFormat { get; set; }
        [Import(true), Comparison(false), Export(false)] public short FontIndex { get; /*追加*/ private set; }
        [Import(true), Comparison(true), Export(true)] public bool IsHidden { get; set; }
        [Import(true), Comparison(true), Export(true)] public bool IsLocked { get; set; }
        [Import(true), Comparison(true), Export(true)] public HorizontalAlignment Alignment { get; set; }
        [Import(true), Comparison(true), Export(true)] public bool WrapText { get; set; }
        [Import(true), Comparison(true), Export(true)] public VerticalAlignment VerticalAlignment { get; set; }
        [Import(true), Comparison(true), Export(true)] public short Rotation { get; set; }
        [Import(true), Comparison(true), Export(true)] public short Indention { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderStyle BorderLeft { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderStyle BorderRight { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderStyle BorderTop { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderStyle BorderBottom { get; set; }
        [Import(true), Comparison(true), Export(true)] public short LeftBorderColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public short RightBorderColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public short TopBorderColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public short BottomBorderColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public FillPattern FillPattern { get; set; }
        [Import(true), Comparison(true), Export(true)] public short FillBackgroundColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public short FillForegroundColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public short BorderDiagonalColor { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderStyle BorderDiagonalLineStyle { get; set; }
        [Import(true), Comparison(true), Export(true)] public BorderDiagonal BorderDiagonal { get; set; }
        /// <summary>
        /// FillBackgroundColorと同義のIColorオブジェクト
        /// NPOIのICellStyle実装クラスでは、一方を更新するともう一方も自動更新される。
        /// 本クラスは自動更新をサポートしていないので、ColorIndex系のみ更新可能としている。
        /// IColor系の更新はサポートしていないので、Comparison, Exportはfalse指定としている。
        /// </summary>
        [Import(false), Comparison(false), Export(false)] public IColor FillBackgroundColorColor { get; }
        /// <summary>
        /// FillForegroundColorと同義のIColorオブジェクト
        /// NPOIのICellStyle実装クラスでは、一方を更新するともう一方も自動更新される。
        /// 本クラスは自動更新をサポートしていないので、ColorIndex系のみ更新可能としている。
        /// IColor系の更新はサポートしていないので、Comparison, Exportはfalse指定としている。
        /// </summary>
        [Import(false), Comparison(false), Export(false)] public IColor FillForegroundColorColor { get; }

        #endregion

        #region "mandatory methods"

        /// <summary>
        /// 必須メソッド実装
        /// DataFormatから実際の書式文字列を特定したもの
        /// </summary>
        /// <returns></returns>
        public String GetDataFormatString()
        {
            return _DataFormatString;
        }
        /// <summary>
        /// 必須メソッド実装(ダミー)
        /// フォント設定(IFont自体がIndexを持っているのでそれがFontIndexに反映される)
        /// </summary>
        /// <param name="font"></param>
        public void SetFont(IFont font)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// 必須メソッド実装(ダミー)
        /// 指定されたICellStyleをクローニングするメソッド
        /// </summary>
        /// <param name="source"></param>
        public void CloneStyleFrom(ICellStyle source)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// ダミーメソッド実装
        ///FontIndexからIFontを特定するメソッド
        /// </summary>
        /// <param name="parentWorkbook"></param>
        /// <returns></returns>
        public IFont GetFont(IWorkbook parentWorkbook)
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion

        #region "properties"

        /// <summary>
        /// 基点とするCellStyle。ApplyStyleFromまはたCommitでセット。
        /// </summary>
        public ICellStyle CellStyle { get; private set; }

        /// <summary>
        /// ICellStyleの拡張:DataFormatをGetDataFormatString()で展開して格納
        /// </summary>
        public string DataFormatString
        {
            get
            {
                return _DataFormatString;
            }
            set
            {
                //上位からの値を保存し、現在のIndexをクリア
                _DataFormatString = value;
                DataFormat = -1;
            }
        }

        /// <summary>
        /// ICellStyleの拡張:FontIndexをGetFont()で展開して格納
        /// </summary>
        public PoiFont PoiFont { get; private set; }

        /// <summary>
        /// 親ISheetクラスインスタンス
        /// </summary>
        private ISheet PoiSheet { get; }

        /// <summary>
        /// 親IWorkbook
        /// </summary>
        private IWorkbook PoiBook
        {
            get { return PoiSheet.Workbook; }
        }

        #endregion

        #region "methods"

        /// <summary>
        /// 指定されたスタイルを自プロパティにインポート
        /// 　(1) ImportAttributeでtrueが指定されているプロパティをコピーする
        /// 　(2) GetDataFormatString()によりDataFormatStringへ文字列展開する
        /// 　(3) FontIndexからその内容をPoiFontに展開する
        /// </summary>
        /// <param name="CellStyle">インポート対象スタイル</param>
        public void ImportFrom(ICellStyle CellStyle) 
        {
            if (SystemParams.UseReflection)
            {
                if (SystemParams.UseReflectionMap)
                {
                    ImportFromViaReflection1(CellStyle);
                }
                else
                {
                    ImportFromViaReflection0(CellStyle);
                }
            }
            else
            {
                ImportFromViaProperty(CellStyle);
            }
        }

        /// <summary>
        /// 指定スタイルが自クラスプロパティと同じスタイルかどうかを判断する
        /// </summary>
        /// <param name="TagetCellStyle">自クラスを同じか判定する対象スタイル</param>
        /// <returns>一致時true</returns>
        public bool Equals(ICellStyle CellStyle)
        {
            bool RetVal = false;
            if (SystemParams.UseReflection)
            {
                if (SystemParams.UseReflectionMap)
                {
                    RetVal = EqualsViaReflection1(CellStyle);
                }
                else
                {
                    RetVal = EqualsViaReflection0(CellStyle);
                }
            }
            else
            {
                RetVal = EqualsViaProperty(CellStyle);
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたスタイルに自プロパティをエクスポート
        /// 　(1) ExportAttributeでtrueが指定されているプロパティをコピーする
        /// 　(2) DataFormatはそのままコピー
        /// 　(3) Font情報はPoiFont.FontをSetFont()する
        /// </summary>
        /// <param name="CellStyle">エクスポート対象スタイル</param>
        public void ExportTo(ICellStyle CellStyle)
        {
            if (SystemParams.UseReflection)
            {
                if (SystemParams.UseReflectionMap)
                {
                    ExportToViaReflection1(CellStyle);
                }
                else
                {
                    ExportToViaReflection0(CellStyle);
                }
            }
            else
            {
                ExportToViaProperty(CellStyle);
            }
        }

        #region "Via Reflection with maps"

        /// <summary>
        /// 指定されたスタイルを自プロパティにインポート
        /// </summary>
        /// <param name="CellStyle">インポート対象スタイル</param>
        public void ImportFromViaReflection1(ICellStyle CellStyle)
        {
            //基点CellTyleを保存
            this.CellStyle = CellStyle;
            //インポートマップに従いインポート
            foreach (var map in _ImportMap)
            {
                map.Value.SetValue(this, _ICellStyleMap[map.Key].GetValue(CellStyle));
            }
            //DataFormatを文字列に展開(プロパティではなく元のFieldに展開)。Indexは保持。
            _DataFormatString = CellStyle.GetDataFormatString();
            //フォント情報を展開。展開したらFontIndexは無効化
            PoiFont = new PoiFont(PoiBook, CellStyle.FontIndex);
            FontIndex = -1;
        }

        /// <summary>
        /// 指定スタイルが自クラスプロパティと同じスタイルかどうかを判断する
        /// </summary>
        /// <param name="TagetCellStyle">自クラスを同じか判定する対象スタイル</param>
        /// <returns>一致時true</returns>
        public bool EqualsViaReflection1(ICellStyle CellStyle)
        {
            bool RetVal = true;
            //コンペアマップに従い比較
            foreach (var map in _CompareMap)
            {
                if (!map.Value.GetValue(this).Equals(_ICellStyleMap[map.Key].GetValue(CellStyle)))
                {
                    RetVal = false;
                    break;
                }
            }
            //一致していればFontも比較
            if (RetVal)
            {
                if (!PoiFont.Equals(CellStyle.GetFont(PoiBook)))
                {
                    RetVal = false;
                }
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたスタイルに自プロパティをエクスポート
        /// 　(1) ExportAttributeでtrueが指定されているプロパティをコピーする
        /// 　(2) DataFormatはそのままコピー
        /// 　(3) Font情報はPoiFont.FontをSetFont()する
        /// </summary>
        /// <param name="CellStyle">エクスポート対象スタイル</param>
        public void ExportToViaReflection1(ICellStyle CellStyle)
        {
            //エクスポートマップに従い比較
            foreach (var map in _ExportMap)
            {
                _ICellStyleMap[map.Key].SetValue(CellStyle, map.Value.GetValue(this));
            }
            //FontからFontをセットする
            CellStyle.SetFont(PoiFont.Font);
            //★★★要検討(その１)★★★
            //FillForegroundColorColor, FillBackfroundColorColorにはSetterがないので直接更新ができない。
            //FillForegroundColor, FillBackfroundColor更新時に、自動更新される模様なので、ここで個別に
            //明示的な更新を行う。念のためFillPatternも更新しておく。
            CellStyle.FillPattern = this.FillPattern;
            CellStyle.FillForegroundColor = this.FillForegroundColor;
            CellStyle.FillBackgroundColor = this.FillBackgroundColor;
            //★★★要検討(その２)★★★
            //Excelでファイルを開き、FillPatternがNoFillのセルに入ると背景色がおかしくなってしまう。
            //どうやらFillBackgroundColorColorに問題がある模様。
            //Excelで編集した場合はFillBackgroundColorColorがNULLなのでそれに倣う。
            //ただしsetterがないのでリフレクションで強制的に実施する。
            if (this.FillPattern == FillPattern.NoFill)
            {
                if (_FillBackgroundColorColorProp != null)
                {
                    _FillBackgroundColorColorProp.SetValue(CellStyle, null);
                    Logger.Debug("FillBackgroundColorColor null-cleared.");
                }
            }
        }

        #endregion

        #region "Via Reflection without maps"

        /// <summary>
        /// 指定されたスタイルを自プロパティにインポート
        /// </summary>
        /// <param name="CellStyle">インポート対象スタイル</param>
        public void ImportFromViaReflection0(ICellStyle CellStyle)
        {
            //基点CellTyleを保存
            this.CellStyle = CellStyle;
            //指定されたCellStyleのプロパティを一覧
            PropertyInfo[] Props = CellStyle.GetType().GetProperties();
            //指定されたCellStyleのプロパティ分ループ
            foreach (PropertyInfo source in Props)
            {
                //同名の自プロパティ情報取得
                PropertyInfo destination = this.GetType().GetProperty(source.Name);
                //ICellSytleプロパティ以外は無視
                if (destination != null)
                {
                    //インポート属性のあるプロパティのみインポート
                    ImportAttribute atr = AttributeUtility.GetPropertyAttribute<ImportAttribute>(this, destination.Name);
                    if (atr != null && atr.Import)
                    {
                        destination.SetValue(this, source.GetValue(CellStyle));
                    }
                }
            }
            //DataFormatを文字列に展開(プロパティではなく元のFieldに展開)。Indexは保持。
            _DataFormatString = CellStyle.GetDataFormatString();
            //フォント情報を展開。展開したらFontIndexは無効化
            PoiFont = new PoiFont(PoiBook, CellStyle.FontIndex);
            FontIndex = -1;
        }

        /// <summary>
        /// 指定スタイルが自クラスプロパティと同じスタイルかどうかを判断する
        /// </summary>
        /// <param name="TagetCellStyle">自クラスを同じか判定する対象スタイル</param>
        /// <returns>一致時true</returns>
        public bool EqualsViaReflection0(ICellStyle CellStyle)
        {
            bool RetVal = true;
            //nullでなくICellStyleであること
            if ((CellStyle != null) && CellStyle is ICellStyle)
            {
                //指定されたCellStyleのプロパティを一覧
                PropertyInfo[] TargetProps = CellStyle.GetType().GetProperties();
                //指定されたCellStyleのプロパティ分ループ
                foreach (PropertyInfo TargetProp in TargetProps)
                {
                    //同名の自プロパティ情報取得
                    PropertyInfo MyProp = this.GetType().GetProperty(TargetProp.Name);
                    //ICellSytleプロパティ以外は無視
                    if (MyProp != null)
                    {
                        //コンペア属性のあるプロパティのみ比較
                        ComparisonAttribute Attr = AttributeUtility.GetPropertyAttribute<ComparisonAttribute>(this, MyProp.Name);
                        if (Attr != null && Attr.Compare)
                        {
                            //プロパティ値不一致
                            if (!MyProp.GetValue(this).Equals(TargetProp.GetValue(CellStyle)))
                            {
                                RetVal = false;
                                break;
                            }
                        }
                    }
                }
                //一致していればFontも比較
                if (RetVal)
                {
                    if (!PoiFont.Equals(CellStyle.GetFont(PoiBook)))
                    {
                        RetVal = false;
                    }
                }
            }
            //nullまたは型違い
            else
            {
                RetVal = false;
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたスタイルに自プロパティをエクスポート
        /// 　(1) ExportAttributeでtrueが指定されているプロパティをコピーする
        /// 　(2) DataFormatはそのままコピー
        /// 　(3) Font情報はPoiFont.FontをSetFont()する
        /// </summary>
        /// <param name="CellStyle">エクスポート対象スタイル</param>
        public void ExportToViaReflection0(ICellStyle CellStyle)
        {
            PropertyInfo[] MyProps = this.GetType().GetProperties();
            //自プロパティを指定されたCellStyleのプロパティ値にセット
            foreach (PropertyInfo MyProp in MyProps)
            {
                //エクスポート属性のあるプロパティのみエクスポート
                ExportAttribute Attr = AttributeUtility.GetPropertyAttribute<ExportAttribute>(this, MyProp.Name);
                if (Attr != null && Attr.Export)
                {
                    //プロパティ情報を取得しプロパティに値をセット
                    //ただしICellStyeから拡張された独自プロパティは無視
                    PropertyInfo TargetProp = CellStyle.GetType().GetProperty(MyProp.Name);
                    TargetProp?.SetValue(CellStyle, MyProp.GetValue(this));
                    Logger.Debug(TargetProp.Name + "=" + (MyProp.GetValue(this) ?? "null"));
                }
            }
            //FontからFontをセットする
            CellStyle.SetFont(PoiFont.Font);
            //★★★要検討(その１)★★★
            //FillForegroundColorColor, FillBackfroundColorColorにはSetterがないので直接更新ができない。
            //FillForegroundColor, FillBackfroundColor更新時に、自動更新される模様なので、ここで個別に
            //明示的な更新を行う。念のためFillPatternも更新しておく。
            CellStyle.FillPattern = this.FillPattern;
            CellStyle.FillForegroundColor = this.FillForegroundColor;
            CellStyle.FillBackgroundColor = this.FillBackgroundColor;
            //★★★要検討(その２)★★★
            //Excelでファイルを開き、FillPatternがNoFillのセルに入ると背景色がおかしくなってしまう。
            //どうやらFillBackgroundColorColorに問題がある模様。
            //Excelで編集した場合はFillBackgroundColorColorがNULLなのでそれに倣う。
            //ただしsetterがないのでリフレクションで強制的に実施する。
            if (this.FillPattern == FillPattern.NoFill)
            {
                PropertyInfo PInf = CellStyle.GetType().GetProperty("FillBackgroundColorColor", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (PInf != null)
                {
                    PInf.SetValue(CellStyle, null);
                    Logger.Debug("FillBackgroundColorColor null-cleared.");
                }
            }
        }

        #endregion

        #region "Via Property"

        /// <summary>
        /// 指定されたスタイルを自プロパティにインポート
        /// </summary>
        /// <param name="CellStyle">インポート対象スタイル</param>
        public void ImportFromViaProperty(ICellStyle CellStyle)
        {
            //基点CellTyleを保存
            this.CellStyle = CellStyle;
            //インポート属性のあるプロパティのみインポート
            this.Index = CellStyle.Index;
            this.ShrinkToFit = CellStyle.ShrinkToFit;
            this.DataFormat = CellStyle.DataFormat;
            this.FontIndex = CellStyle.FontIndex;
            this.IsHidden = CellStyle.IsHidden;
            this.IsLocked = CellStyle.IsLocked;
            this.Alignment = CellStyle.Alignment;
            this.WrapText = CellStyle.WrapText;
            this.VerticalAlignment = CellStyle.VerticalAlignment;
            this.Rotation = CellStyle.Rotation;
            this.Indention = CellStyle.Indention;
            this.BorderLeft = CellStyle.BorderLeft;
            this.BorderRight= CellStyle.BorderRight;
            this.BorderTop = CellStyle.BorderTop;
            this.BorderBottom = CellStyle.BorderBottom;
            this.LeftBorderColor = CellStyle.LeftBorderColor;
            this.RightBorderColor = CellStyle.RightBorderColor;
            this.TopBorderColor = CellStyle.TopBorderColor;
            this.BottomBorderColor = CellStyle.BottomBorderColor;
            this.FillPattern = CellStyle.FillPattern;
            this.FillBackgroundColor = CellStyle.FillBackgroundColor;
            this.FillForegroundColor = CellStyle.FillForegroundColor;
            this.BorderDiagonalColor = CellStyle.BorderDiagonalColor;
            this.BorderDiagonalLineStyle = CellStyle.BorderDiagonalLineStyle;
            this.BorderDiagonal = CellStyle.BorderDiagonal;
            //this.FillBackgroundColorColor
            //this.FillForegroundColorColor
            //DataFormatを文字列に展開(プロパティではなく元のFieldに展開)。Indexは保持。
            _DataFormatString = CellStyle.GetDataFormatString();
            //フォント情報を展開。展開したらFontIndexは無効化
            PoiFont = new PoiFont(PoiBook, CellStyle.FontIndex);
            FontIndex = -1;
        }

        /// <summary>
        /// 指定スタイルが自クラスプロパティと同じスタイルかどうかを判断する
        /// </summary>
        /// <param name="TagetCellStyle">自クラスを同じか判定する対象スタイル</param>
        /// <returns>一致時true</returns>
        public bool EqualsViaProperty(ICellStyle CellStyle)
        {
            bool RetVal = true;
            //nullでなくICellStyleであること
            if ((CellStyle != null) && CellStyle is ICellStyle)
            {
                //コンペア属性のあるプロパティのみ比較
                //this.Index
                if (!this.ShrinkToFit.Equals(CellStyle.ShrinkToFit)) RetVal = false;
                if (!this.DataFormat.Equals(CellStyle.DataFormat)) RetVal = false;
                //this.FontIndex
                if (!this.IsHidden.Equals(CellStyle.IsHidden)) RetVal = false;
                if (!this.IsLocked.Equals(CellStyle.IsLocked)) RetVal = false;
                if (!this.Alignment.Equals(CellStyle.Alignment)) RetVal = false;
                if (!this.WrapText.Equals(CellStyle.WrapText)) RetVal = false;
                if (!this.VerticalAlignment.Equals(CellStyle.VerticalAlignment)) RetVal = false;
                if (!this.Rotation.Equals(CellStyle.Rotation)) RetVal = false;
                if (!this.Indention.Equals(CellStyle.Indention)) RetVal = false;
                if (!this.BorderLeft.Equals(CellStyle.BorderLeft)) RetVal = false;
                if (!this.BorderRight.Equals(CellStyle.BorderRight)) RetVal = false;
                if (!this.BorderTop.Equals(CellStyle.BorderTop)) RetVal = false;
                if (!this.BorderBottom.Equals(CellStyle.BorderBottom)) RetVal = false;
                if (!this.LeftBorderColor.Equals(CellStyle.LeftBorderColor)) RetVal = false;
                if (!this.RightBorderColor.Equals(CellStyle.RightBorderColor)) RetVal = false;
                if (!this.TopBorderColor.Equals(CellStyle.TopBorderColor)) RetVal = false;
                if (!this.BottomBorderColor.Equals(CellStyle.BottomBorderColor)) RetVal = false;
                if (!this.FillPattern.Equals(CellStyle.FillPattern)) RetVal = false;
                if (!this.FillBackgroundColor.Equals(CellStyle.FillBackgroundColor)) RetVal = false;
                if (!this.FillForegroundColor.Equals(CellStyle.FillForegroundColor)) RetVal = false;
                if (!this.BorderDiagonalColor.Equals(CellStyle.BorderDiagonalColor)) RetVal = false;
                if (!this.BorderDiagonalLineStyle.Equals(CellStyle.BorderDiagonalLineStyle)) RetVal = false;
                if (!this.BorderDiagonal.Equals(CellStyle.BorderDiagonal)) RetVal = false;
                //this.FillBackgroundColorColor
                //this.FillForegroundColorColor
                //一致していればFontも比較
                if (RetVal)
                {
                    if (!PoiFont.Equals(CellStyle.GetFont(PoiBook)))
                    {
                        RetVal = false;
                    }
                }
            }
            //nullまたは型違い
            else
            {
                RetVal = false;
            }
            return RetVal;
        }

        /// <summary>
        /// 指定されたスタイルに自プロパティをエクスポート
        /// </summary>
        /// <param name="CellStyle">エクスポート対象スタイル</param>
        public void ExportToViaProperty(ICellStyle CellStyle)
        {
            PropertyInfo[] MyProps = this.GetType().GetProperties();
            //自プロパティを指定されたCellStyleのプロパティ値にセット
            foreach (PropertyInfo MyProp in MyProps)
            {
                //エクスポート属性のあるプロパティのみエクスポート
                ExportAttribute Attr = AttributeUtility.GetPropertyAttribute<ExportAttribute>(this, MyProp.Name);
                //CellStyle.Index
                CellStyle.ShrinkToFit = this.ShrinkToFit;
                CellStyle.DataFormat = this.DataFormat;
                //Cell.StyleFontIndex
                CellStyle.IsHidden = this.IsHidden;
                CellStyle.IsLocked = this.IsLocked;
                CellStyle.Alignment = this.Alignment;
                CellStyle.WrapText = this.WrapText;
                CellStyle.VerticalAlignment = this.VerticalAlignment;
                CellStyle.Rotation = this.Rotation;
                CellStyle.Indention = this.Indention;
                CellStyle.BorderLeft = this.BorderLeft;
                CellStyle.BorderRight = this.BorderRight;
                CellStyle.BorderTop = this.BorderTop;
                CellStyle.BorderBottom = this.BorderBottom;
                CellStyle.LeftBorderColor = this.LeftBorderColor;
                CellStyle.RightBorderColor = this.RightBorderColor;
                CellStyle.TopBorderColor = this.TopBorderColor;
                CellStyle.BottomBorderColor = this.BottomBorderColor;
                CellStyle.FillPattern = this.FillPattern;
                CellStyle.FillBackgroundColor = this.FillBackgroundColor;
                CellStyle.FillForegroundColor = this.FillForegroundColor;
                CellStyle.BorderDiagonalColor = this.BorderDiagonalColor;
                CellStyle.BorderDiagonalLineStyle = this.BorderDiagonalLineStyle;
                CellStyle.BorderDiagonal = this.BorderDiagonal;
                //CellStyle.FillBackgroundColorColor
                //CellStyle.FillForegroundColorColor
            }
            //FontからFontをセットする
            CellStyle.SetFont(PoiFont.Font);
            //★★★要検討(その１)★★★
            //FillForegroundColorColor, FillBackfroundColorColorにはSetterがないので直接更新ができない。
            //FillForegroundColor, FillBackfroundColor更新時に、自動更新される模様なので、ここで個別に
            //明示的な更新を行う。念のためFillPatternも更新しておく。
            CellStyle.FillPattern = this.FillPattern;
            CellStyle.FillForegroundColor = this.FillForegroundColor;
            CellStyle.FillBackgroundColor = this.FillBackgroundColor;
            //★★★要検討(その２)★★★
            //Excelでファイルを開き、FillPatternがNoFillのセルに入ると背景色がおかしくなってしまう。
            //どうやらFillBackgroundColorColorに問題がある模様。
            //Excelで編集した場合はFillBackgroundColorColorがNULLなのでそれに倣う。
            //ただしsetterがないのでリフレクションで強制的に実施する。
            if (this.FillPattern == FillPattern.NoFill)
            {
                PropertyInfo PInf = CellStyle.GetType().GetProperty("FillBackgroundColorColor", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (PInf != null)
                {
                    PInf.SetValue(CellStyle, null);
                    Logger.Debug("FillBackgroundColorColor null-cleared.");
                }
            }
        }

        #endregion

        /// <summary>
        /// 変更を確定する
        /// </summary>
        public short Commit()
        {
            //フォントの変更をCommitする
            PoiFont.Commit();
            //DataFormatの変更をCommitする
            CommitDataFormat();
            //基点を初期化
            CellStyle = null;
            //マスター上にあるかチェックしあればそれを使う(ただしIndex=1以上
            for (short i = 1; i < PoiBook.NumCellStyles; i++)
            {
                if (this.Equals(PoiBook.GetCellStyleAt(i)))
                {
                    CellStyle = PoiBook.GetCellStyleAt(i);
                    Logger.Debug("CurrentStyle:[Index:" + this.Index + "] => Style[Index:" + i + "] is found in Book.");
                    break;
                }
            }
            //マスターになければ未使用スタイルをチェックしあればそれを使う
            if (CellStyle == null)
            {
                short AvalableIndex = -1;
                Dictionary<short, int> StyleUsages = StyleUtil.GetCellStyleUsage(PoiBook);
                //いま保持しているスタイルがこのセルのみで使用している場合
                if (StyleUsages.ContainsKey(Index) && StyleUsages[Index] == 1)
                {
                    //現状がIndex=0でなければそれを使う
                    if (Index != 0)
                    {
                        AvalableIndex = Index;
                        Logger.Debug("CurrentStyle:[Index:" + this.Index + "] => Style[Index:" + this.Index + "] Current Style is updatable.");
                    }
                }
                //いま保持しているスタイルがこのセル以外でも使用されている場合
                else
                {
                    //未使用スタイルの検索
                    foreach (KeyValuePair<short, int> Style in StyleUsages)
                    {
                        //Index=0以外のスタイルで未使用のもの
                        if (Style.Key != 0 && Style.Value == 0)
                        {
                            AvalableIndex = Style.Key;
                            Logger.Debug("CurrentStyle:[Index:" + this.Index + "] => Style[Index:" + Style.Key + "] is unused and available.");
                            break;
                        }
                    }
                }
                //再利用可能なスタイルがあったならそれを使う
                if (AvalableIndex > 0)
                {
                    //未使用スタイルの取得
                    CellStyle = PoiBook.GetCellStyleAt(AvalableIndex);
                    //今回の内容を反映
                    this.ExportTo(CellStyle);
                }
            }
            //マスターになく、未使用スタイルにもなければ新規にCreateする
            if (CellStyle == null)
            {
                //新規スタイルを生成し今回の内容を反映
                CellStyle = PoiBook.CreateCellStyle();
                //今回の内容を反映
                this.ExportTo(CellStyle);
                Logger.Debug("CurrentStyle:[Index:" + this.Index + "] => Style.[Index:" + CellStyle.Index + "] is newly created.");
            }
            //Indexを新しい値に更新する
            this.Index = CellStyle.Index;
            this.FontIndex = PoiFont.Index;
            //新しいIndexでリターン
            return this.Index;
        }

        /// <summary>
        /// DataFormatのコミット
        /// </summary>
        private void CommitDataFormat()
        {
            //フォーマットが更新されている場合はマスター上に利用可能なものがあるかチェック
            if (this.DataFormat == -1)
            {
                //ビルトインフォーマットに一致するものがあるかチェック
                string[] Builtin = BuiltinFormats.GetAll();
                for (short index = 0; index < Builtin.Length; index++)
                {
                    if (Builtin[index] == this.DataFormatString)
                    {
                        this.DataFormat = index;
                        break;
                    }
                }
                //ビルトインになければユーザ設定をチェック
                if (this.DataFormat == -1)
                {
                    //ユーザ設定のフォーマットマスター上に一致するものがあるチェック
                    SortedDictionary<short, string> formats = StyleUtil.GetNumberFormats(PoiBook);
                    foreach (KeyValuePair<short, string> format in formats)
                    {
                        //プロパティにセットされた値がマスター上にある場合はそのIndexを利用
                        if (format.Value == this.DataFormatString)
                        {
                            this.DataFormat = format.Key;
                            break;
                        }
                    }
                }
            }
            //マスター上になければ新規作成する
            if (this.DataFormat == -1)
            {
                this.DataFormat = PoiBook.CreateDataFormat().GetFormat(this.DataFormatString);
            }
        }

        /// <summary>
        /// プロパティマップの作成
        /// </summary>
        private void CreateMaps()
        {
            //初期化
            _ImportMap = new Dictionary<string, PropertyInfo>();
            _CompareMap = new Dictionary<string, PropertyInfo>();
            _ExportMap = new Dictionary<string, PropertyInfo>();
            _PoiCellStyleMap = new Dictionary<string, PropertyInfo>();
            _ICellStyleMap = new Dictionary<string, PropertyInfo>();
            //自クラスのマップ作成
            PropertyInfo[] Props = this.GetType().GetProperties();
            foreach (PropertyInfo Prop in Props)
            {
                _PoiCellStyleMap.Add(Prop.Name, Prop);
                //インポート属性
                ImportAttribute imp = AttributeUtility.GetPropertyAttribute<ImportAttribute>(this, Prop.Name);
                if (imp != null && imp.Import)
                {
                    _ImportMap.Add(Prop.Name, Prop);
                }
                //コンペア属性
                ComparisonAttribute cmp = AttributeUtility.GetPropertyAttribute<ComparisonAttribute>(this, Prop.Name);
                if (cmp != null && cmp.Compare)
                {
                    _CompareMap.Add(Prop.Name, Prop);
                }
                //エクスポート属性
                ExportAttribute exp = AttributeUtility.GetPropertyAttribute<ExportAttribute>(this, Prop.Name);
                if (exp != null && exp.Export)
                {
                    _ExportMap.Add(Prop.Name, Prop);
                }
            }
            //ICellStyleのマップ作成(仮にデフォルトスタイルで作成)
            ICellStyle CellStyle = this.PoiBook.GetCellStyleAt(0);
            PropertyInfo[] ICellProps = CellStyle.GetType().GetProperties();
            foreach (PropertyInfo Prop in ICellProps)
            {
                _ICellStyleMap.Add(Prop.Name, Prop);
            }
            //FillBackgroundColorColor
            _FillBackgroundColorColorProp = CellStyle.GetType().GetProperty("FillBackgroundColorColor", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            //生成件数のみログ
            Logger.Debug(
                "_ImportMap[" + _ImportMap.Count + "] "+
                "_CompareMap[" + _CompareMap.Count + "] "+
                "_ExportMap[" + _ExportMap.Count + "] " +
                "_PoiCellStyleMap[" + _PoiCellStyleMap.Count + "] " +
                "_ICellStyleMap[" + _ICellStyleMap.Count + "]");
        }

        #endregion
    }
}
