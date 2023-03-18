using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Workbooks interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Workbooks : IEnumerable
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    int Count { get; }
    //    Workbook Item { get; }
    //    [IndexerName("_Default")]
    //    Workbook this[object Index] { get; }
    //    Workbook Add([Optional] object Template);
    //    void Close();
    //    new IEnumerator GetEnumerator();
    //    Workbook _Open(string Filename, [Optional] object UpdateLinks, [Optional] object ReadOnly, [Optional] object Format, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object IgnoreReadOnlyRecommended, [Optional] object Origin, [Optional] object Delimiter, [Optional] object Editable, [Optional] object Notify, [Optional] object Converter, [Optional] object AddToMru);
    //    void __OpenText(string Filename, [Optional] object Origin, [Optional] object StartRow, [Optional] object DataType, XlTextQualifier TextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, [Optional] object ConsecutiveDelimiter, [Optional] object Tab, [Optional] object Semicolon, [Optional] object Comma, [Optional] object Space, [Optional] object Other, [Optional] object OtherChar, [Optional] object FieldInfo, [Optional] object TextVisualLayout);
    //    void _OpenText(string Filename, [Optional] object Origin, [Optional] object StartRow, [Optional] object DataType, XlTextQualifier TextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, [Optional] object ConsecutiveDelimiter, [Optional] object Tab, [Optional] object Semicolon, [Optional] object Comma, [Optional] object Space, [Optional] object Other, [Optional] object OtherChar, [Optional] object FieldInfo, [Optional] object TextVisualLayout, [Optional] object DecimalSeparator, [Optional] object ThousandsSeparator);
    //    Workbook Open(string Filename, [Optional] object UpdateLinks, [Optional] object ReadOnly, [Optional] object Format, [Optional] object Password, [Optional] object WriteResPassword, [Optional] object IgnoreReadOnlyRecommended, [Optional] object Origin, [Optional] object Delimiter, [Optional] object Editable, [Optional] object Notify, [Optional] object Converter, [Optional] object AddToMru, [Optional] object Local, [Optional] object CorruptLoad);
    //    void OpenText(string Filename, [Optional] object Origin, [Optional] object StartRow, [Optional] object DataType, XlTextQualifier TextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, [Optional] object ConsecutiveDelimiter, [Optional] object Tab, [Optional] object Semicolon, [Optional] object Comma, [Optional] object Space, [Optional] object Other, [Optional] object OtherChar, [Optional] object FieldInfo, [Optional] object TextVisualLayout, [Optional] object DecimalSeparator, [Optional] object ThousandsSeparator, [Optional] object TrailingMinusNumbers, [Optional] object Local);
    //    Workbook OpenDatabase(string Filename, [Optional] object CommandText, [Optional] object CommandType, [Optional] object BackgroundQuery, [Optional] object ImportDataAs);
    //    void CheckOut(string Filename);
    //    bool CanCheckOut(string Filename);
    //    Workbook _OpenXML(string Filename, [Optional] object Stylesheets);
    //    Workbook OpenXML(string Filename, [Optional] object Stylesheets, [Optional] object LoadOption);
    //}

    /// <summary>
    /// Workbooksクラス
    /// Microsoft.Office.Interop.Excel.Workbooksをエミュレート
    /// NpoiWrapperクラスのプロパティとしてのみコンストラクトされる
    /// ユーザからは直接コンストラクトさせないのでコンストラクタはinternalにしている
    /// </summary>
    public class Workbooks : IEnumerable, IEnumerator
    {
        #region "fields"

        /// <summary>
        /// Workbookリスト
        /// </summary>
        private readonly List<Workbook> _Item = new List<Workbook>();

        /// <summary>
        /// Workbookに付与する通し番号
        /// </summary>
        private int BookIndex = 0;

        /// <summary>
        /// Enumrator用インデクス
        /// </summary>
        private int EnumeratorIndex = -1;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// NoiWrapperのプロパティとしてのみコンストラクトされる
        /// <param name="ParentApp">親Application</param>
        /// </summary>
        internal Workbooks(Application ParentApplication)
        {
            this.Parent = ParentApplication;
        }

        #endregion

        #region "interface implementations"

        /// <summary>
        /// GetEnumeratorの実装
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            Reset();
            return (IEnumerator)this;
        }
        /// <summary>
        /// IEnumerator.MoveNextの実装
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            bool RetVal = false;
            EnumeratorIndex += 1;
            if (EnumeratorIndex < _Item.Count)
            {
                RetVal = true;
            }
            return RetVal;
        }
        /// <summary>
        /// IEnumerator.Current実装
        /// </summary>
        public object Current { get { return _Item[EnumeratorIndex]; } }
        /// <summary>
        /// IEnumerator.Resetの実装
        /// </summary>
        public void Reset() { EnumeratorIndex = -1; }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return this.Parent; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Application Parent { get; }

        #endregion

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// Excelブックの作成
        /// </summary>
        /// <param name="Template">新しいブックの作成方法。XlWBATemplateはxlWBATWorksheetのみサポート。ファイル名が指定された場合は拡張子のみ判断。</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Add(object Template = null)
        {
            bool Excel97_2003 = false;
            if (Template != null)
            {
                if (Template is XlWBATemplate TempleteEnumValue)
                {
                    if (TempleteEnumValue != XlWBATemplate.xlWBATWorksheet)
                    {
                        throw new ArgumentException("Uusupported XlWBATemplate value is specified. (This implementation supports only xlWBATWorksheet.)");
                    }

                }
                else if (Template is string TemplateFile)
                {
                    if(TemplateFile.EndsWith("xls"))
                    {
                        Excel97_2003 = true;
                    }
                }
            }
            Workbook Book = new Workbook(this.Parent, GetNextBookIndex(), Excel97_2003);
            _Item.Add(Book);
            return Book;
        }

        /// <summary>
        /// Excelブックを開く
        /// </summary>
        /// <param name="FileNanme">フルパスファイ名</param>
        /// <param name="....">Filename以外はすべて無視します。</param>
        /// <returns>Workbookクラスインスタンス</returns>
        public Workbook Open(
            string Filename, object UpdateLinks = null, object ReadOnly = null, object Format = null,
            object Password = null, object WriteResPassword = null, object IgnoreReadOnlyRecommended = null,
            object Origin = null, object Delimiter = null, object Editable = null, object Notify = null,
            object Converter = null, object AddToMru = null, object Local = null, object CorruptLoad = null)
        {
            Workbook Book = new Workbook(this.Parent, GetNextBookIndex(), Filename);
            _Item.Add(Book);
            return Book;
        }

        #endregion

        #region "private methods"

        /// <summary>
        /// 次のBookIndex値を取得
        /// </summary>
        /// <returns></returns>
        private int GetNextBookIndex()
        {
            this.BookIndex += 1;
            return this.BookIndex;
        }

        #endregion

        #endregion

        #region "indexers"

        /// <summary>
        /// インデクサ
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        [IndexerName("_Default")]
        public Workbook this[int index] { get { return _Item[index]; } }

        #endregion
    }
}
