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
            //指定されたスタイルをインポート
            this.CellStyle = this.PoiBook.GetCellStyleAt(StyleIndex);
            ApplyStyleFrom(this.CellStyle);
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
        [Import(false), Comparison(true), Export(true)] public BorderDiagonal BorderDiagonal { get; set; }
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
        [Import(false), Comparison(true), Export(false)]
        public string DataFormatString
        {
            get
            {
                return _DataFormatString;
            }
            set
            {
                _DataFormatString = value;
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
        public void ApplyStyleFrom(ICellStyle CellStyle)
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
            //DataFormatを文字列に展開。展開したらDataFormatは無効化
            DataFormatString = CellStyle.GetDataFormatString();
            DataFormat = -1;
            //フォント情報を展開。展開したらFontIndexは無効化
            PoiFont = new PoiFont(PoiBook, CellStyle.FontIndex);
            FontIndex = -1;
        }

        /// <summary>
        /// 指定スタイルががこのスタイルと同じスタイルかどうかを判断する
        /// </summary>
        /// <param name="TagetCellStyle"></param>
        /// <returns></returns>
        public bool StyleEquals(ICellStyle CellStyle)
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
                    if (!PoiFont.StyleEquals(CellStyle.GetFont(PoiBook)))
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
                if (this.StyleEquals(PoiBook.GetCellStyleAt(i)))
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
                    this.ApplyTo(CellStyle);
                }
            }
            //マスターになく、未使用スタイルにもなければ新規にCreateする
            if (CellStyle == null)
            {
                //新規スタイルを生成し今回の内容を反映
                CellStyle = PoiBook.CreateCellStyle();
                //今回の内容を反映
                this.ApplyTo(CellStyle);
                Logger.Debug("CurrentStyle:[Index:" + this.Index + "] => Style.[Index:" + CellStyle.Index + "] is newly created.");
            }
            //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            // NPOIのバグ対応
            // Fill系のColorを指すると予期しないColorColorがセットされてしまう。
            // なのでコンパイラのアクセス制限を強引に乗り越えてnullをセットする。
            //BindingFlags Flag = BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty;
            //PropertyInfo BGColor = CellStyle.GetType().GetProperty("FillBackgroundColorColor", Flag);
            //BGColor?.SetValue(CellStyle, null);
            //PropertyInfo FGColor = CellStyle.GetType().GetProperty("FillForegroundColorColor", Flag);
            //FGColor?.SetValue(CellStyle, null);
            //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
            //Indexを新しい値に更新する
            this.Index = CellStyle.Index;
            this.FontIndex = PoiFont.Index;
            //新しいIndexでリターン
            return this.Index;
        }

        /// <summary>
        /// 指定されたスタイルに自プロパティをエクスポート
        /// 　(1) ExportAttributeでtrueが指定されているプロパティをコピーする
        /// 　(2) DataFormatはそのままコピー
        /// 　(3) Font情報はPoiFont.FontをSetFont()する
        /// </summary>
        /// <param name="CellStyle">エクスポート対象スタイル</param>
        public void ApplyTo(ICellStyle CellStyle)
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
        }

        private void CommitDataFormat()
        {
            this.DataFormat = -1;
            //マスター上に存在するかチェックしあればそれを使う
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
            //マスター上になければ新規作成する
            if (this.DataFormat == -1)
            {
                this.DataFormat = PoiBook.CreateDataFormat().GetFormat(this.DataFormatString);
            }
        }

        #endregion
    }
}
