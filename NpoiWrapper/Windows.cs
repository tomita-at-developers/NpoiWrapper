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
        public Application Application
        {
            get
            {
                Application RetVal = null;
                if (this.Parent is Application app)
                {
                    RetVal = app;
                }
                else if (this.Parent is Workbook wb)
                {
                    RetVal = wb.Application;
                }
                return RetVal;
            }
        }
        public XlCreator Creator { get { return Application.Creator; } }
        public object Parent { get; }   //Applicationの場合とWorkbookの場合とがある。

        private Dictionary<int, Window> _Item = new Dictionary<int, Window>();
        private int ItemIndex { get; set; } = 0;

        internal Windows(object ParentBook)
        {
            this.Parent = ParentBook;
        }

        public int Count
        {
            get { return _Item.Count; }
        }

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
            this.ItemIndex += 1;
            return this.ItemIndex;
        }
        /// <summary>
        /// ItemIndex：リセット
        /// </summary>
        internal void ResetItemIndex()
        {
            this.ItemIndex = 0;
        }
    }
}