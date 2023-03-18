using NPOI.SS.Formula.Functions;
using NPOI.Util;
using Org.BouncyCastle.Cms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Management;

namespace Developers.NpoiWrapper
{
    //public interface Windows
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    int Count { get; }
    //    Window Item { get; }
    //    [IndexerName("_Default")]
    //    Window this[object Index] { get; }
    //    bool SyncScrollingSideBySide { get; set; }
    //    object Arrange(XlArrangeStyle ArrangeStyle = XlArrangeStyle.xlArrangeStyleTiled, [Optional] object ActiveWorkbook, [Optional] object SyncHorizontal, [Optional] object SyncVertical);
    //    IEnumerator GetEnumerator();
    //    bool CompareSideBySideWith(object WindowName);
    //    bool BreakSideBySide();
    //    void ResetPositionsSideBySide();
    //}
    public class Windows
    {
        #region "fields"

        /// <summary>
        /// Windowコレクション
        /// </summary>
        private readonly Dictionary<int, Window> _Item = new Dictionary<int, Window>();

        /// <summary>
        /// Windowコレクション用Index
        /// </summary>
        private int _ItemIndex = 0;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentObject">親オブジェクト(ApplicationまたはWorkbook</param>
        internal Windows(object ParentObject)
        {
            this.Parent = ParentObject;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application
        {
            get
            {
                Application RetVal = null;
                //親がApplicationクラスの場合は親自身をセット
                if (this.Parent is Application app)
                {
                    RetVal = app;
                }
                //親がWorkbookの場合はWorkbookが持つApplicationをセット
                else if (this.Parent is Workbook wb)
                {
                    RetVal = wb.Application;
                }
                return RetVal;
            }
        }
        public XlCreator Creator { get { return Application.Creator; } }
        public object Parent { get; }   //Applicationの場合とWorkbookの場合とがある。

        /// <summary>
        /// C#では引数を指定するプロパティが実現できなかった。
        /// なのでItemは実装しない。Indexer経由での取得のみサポート。
        /// </summary>
        // public Window Item { get; }

        /// <summary>
        /// Windowの数
        /// </summary>
        public int Count
        {
            get { return _Item.Count; }
        }

        #endregion

        #endregion

        #region "methods"

        #region "internal methods"

        /// <summary>
        /// Windowの追加
        /// </summary>
        /// <param name="Item"></param>
        internal void Add(Window Item)
        {
            int NextIndex = _Item.Count + 1;
            _Item.Add(NextIndex, Item);
        }

        /// <summary>
        /// Windowの削除
        /// </summary>
        /// <param name="Item"></param>
        internal void Remove(Window Item)
        {
            List<KeyValuePair<int, Window>> Target = _Item.Where(w => w.Value.Parent.Index == Item.Parent.Index).ToList();
            foreach (KeyValuePair<int, Window> W in Target)
            {
                _Item.Remove(W.Key);
            }
        }

        /// <summary>
        /// 指定されたWindowをActiveWindowに(Key=1に)する。
        /// 指定されたWindowが_Itemになければ追加する。
        /// </summary>
        /// <param name="Item"></param>
        internal void SetActiveWindow(Window Item)
        {
            //指定されたWindowがあれば、そのKeyを-1にする
            List<KeyValuePair<int, Window>> Target = _Item.Where(w => w.Value.Parent.Index == Item.Parent.Index).ToList();
            if(Target.Count > 0)
            {
                _Item.Add(-1, Target[0].Value);
                _Item.Remove(Target[0].Key);
            }
            //なければKey=-1で追加する。
            else
            {
                _Item.Add(-1, Item);
            }
            //ソート(Key1から振り直し)
            Sort();
        }

        /// <summary>
        /// Indexの振り直し
        /// ActiveWindow(Index=1)はSort後も先頭にくるはずなので、気にせず１から振り直す。
        /// </summary>
        internal void Sort()
        {
            Dictionary<int, Window> Backup = new Dictionary<int, Window>(_Item);
            _Item.Clear();
            ResetItemIndex();
            foreach (KeyValuePair<int, Window> w in Backup.OrderBy(item => item.Key))
            {
                this._Item.Add(GetNextItemIndex(), w.Value);
            }
        }

        /// <summary>
        /// ItemIndex：次の値を取得
        /// </summary>
        /// <returns></returns>
        internal int GetNextItemIndex()
        {
            this._ItemIndex += 1;
            return this._ItemIndex;
        }
        /// <summary>
        /// ItemIndex：リセット
        /// </summary>
        internal void ResetItemIndex()
        {
            this._ItemIndex = 0;
        }

        #endregion

        #endregion

        #region "indexers"

        [IndexerName("_Default")]
        public Window this[object Index]
        {
            get
            {
                Window RetVal = null;
                //intの場合
                if (Index is int IntegerIndex)
                {
                    RetVal = _Item[IntegerIndex];
                }
                //stringの場合
                else if (Index is string StringIndex)
                {
                    //指定された文字列のCaptionを持つレコードの検索
                    Dictionary<int, Window> Select = _Item.Where(w => w.Value.Caption == StringIndex).ToDictionary(w => w.Key, w => w.Value);
                    if (Select.Count == 1)
                    {
                        RetVal = Select[0];
                    }
                }
                return RetVal;
            }
        }

        #endregion

    }
}