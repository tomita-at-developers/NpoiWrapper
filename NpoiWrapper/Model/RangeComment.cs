using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;

namespace Developers.NpoiWrapper.Model
{
    internal class RangeComment
    {
        #region "fields"

        /// <summary>
        /// log4net
        /// </summary>
        private static readonly log4net.ILog Logger
            = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name);

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange"></param>
        public RangeComment(Range ParentRange)
        {
            this.Parent = ParentRange;
        }

        #endregion

        #region "Properties"

        /// <summary>
        /// このRangeの先頭セルからCommentオブジェクトを読み出す。(コメントが存在しない場合はnull)
        /// </summary>
        public Comment Comment
        {
            get
            {
                Comment RetVal = null;
                //セルが実在すること
                ICell Cell = NpoiWrapper.Utils.CellUtil.GetCell(Parent.Parent.PoiSheet, FirstCellAddress);
                if (Cell != null)
                {
                    //コメントが存在すること
                    if (Cell.CellComment != null)
                    {
                        //Cell.CellCommentからCommentクラスを生成
                        RetVal = new Comment(
                            Parent,
                            Cell.CellComment.Author,
                            Cell.CellComment.String.String,
                            Cell.CellComment.Visible,
                            FirstCellAddress);
                    }
                }
                return RetVal;
            }
        }

        /// <summary>
        /// 親Rangeインスタンス
        /// </summary>
        private Range Parent { get; }

        /// <summary>
        /// このRangeが単一セルのRangeであるか否か
        /// </summary>
        private bool HasSingleCell
        {
            get
            {   
                bool RetVal = false;
                if (Parent.SafeAddressList.CountRanges() == 1
                    && Parent.SafeAddressList.GetCellRangeAddress(0).NumberOfCells == 1)
                {
                    RetVal = true;
                }
                return RetVal;
            }
        }

        /// <summary>
        /// 親Rangeの先頭セルアドレス
        /// </summary>
        private CellAddress FirstCellAddress
        {
            get
            {
                return new CellAddress(
                    Parent.SafeAddressList.GetCellRangeAddress(0).FirstRow,
                    Parent.SafeAddressList.GetCellRangeAddress(0).FirstColumn);
            }
        }

        #endregion

        #region "methods"

        /// <summary>
        /// コメントの追加
        /// </summary>
        /// <param name="Text">コメント文字列</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public Comment AddComment(object Text = null)
        {
            Comment Comment;
            //単一セルのRangeでありコメント未設定であること。
            if (this.HasSingleCell && this.Comment == null)
            {
                //アドレス特定
                int RowIndex = this.FirstCellAddress.Row;
                int ColumnIndex = this.FirstCellAddress.Column;
                //セルの取得
                //列の取得(なければ生成)
                IRow Row = Parent.Parent.PoiSheet.GetRow(RowIndex) ?? Parent.Parent.PoiSheet.CreateRow(RowIndex);
                ICell Cell = Row.GetCell(ColumnIndex) ?? Row.CreateCell(ColumnIndex);
                //コメント文字列の解析
                string CommentText = string.Empty;
                if (Text is string SafeText)
                {
                    CommentText = SafeText;
                }
                //コメントの生成
                Comment = new Comment(Parent, System.Environment.UserName, CommentText, false, new CellAddress(Cell));
                //セルへの適用
                Comment.Apply();
            }
            //複数セル選択時、コメント既存時は例外スロー
            else
            {
                throw new ArgumentException("To add comment, Range should contain single cell with no comment.");
            }
            return Comment;
        }

        #endregion
    }
}
