using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace Developers.NpoiWrapper
{
    //----------------------------------------------------------------------------------------------
    // Comment interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Comment
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    string Author { get; }
    //    Shape Shape { get; }
    //    bool Visible { get; set; }
    //    string Text([Optional] object Text, [Optional] object Start, [Optional] object Overwrite);
    //    void Delete();
    //    Comment Next();
    //    Comment Previous();
    //}
    //----------------------------------------------------------------------------------------------
    //  Corresponding interface in NPOI IComment is shown below...
    //  ＜ 注意＞
    //  - Address, Row, Column, ClientAnchr.Row1/Col1
    //    - Address
    //      コメントを所有するセルのアドレス
    //    - Row/Column
    //      コメントを所有するセルのアドレスであり、Addressと同一値。
    //    - ClientAnchor.Row1/Col1
    //      ClientAnchor作成時に指定するRow1/Col1は、コメントを所有するセルのアドレスでなければならない。
    //      ICommentインスタンスにClientAnchorを設定する際のチェックで、ClientAnchorが示すアドレスで
    //      既にCommentを所有するセルが存在すると例外が発生する。(おそらくバグ)
    //    - ClientAnchor.Dx1/Dy1/Dx2/Dy2
    //      対応するRow/ColのTopLeftからのオフセットを指定できるが、どんなに大きな値を設定しても、
    //      対応するRow/ColのBottomRightを超えてオフセットされることはない。
    //      従って事実上、Col1<Col2,　Row1<Row2でなければならない。
    //      自由な値を設定できるものの、
    //  - Dictionary<CellAddress, IComment> Comments = ISheet.GetCellComments();
    //  　　シート内の全Commentを取得可能。ただし追加順なので、アドレスでソートする必要あり。
    //  　　また、ここで取得されるValue.ClientAnchorにアクセスすると例外発生するので利用不可。
    //  　　ClientAnchorはあくまで作成時の一時的な値であるらしく、作成後その実体は存在しない模様。
    //      Move/Resizeを考慮すればある意味当然のことかも知れない。
    //----------------------------------------------------------------------------------------------
    //public interface IComment
    //{
    //    bool Visible { get; set; }
    //    CellAddress Address { get; set; }
    //    int Row { get; set; }
    //    int Column { get; set; }
    //    string Author { get; set; }
    //    IRichTextString String { get; set; }
    //    IClientAnchor ClientAnchor { get; }
    //    void SetAddress(int row, int col);
    //}
    //----------------------------------------------------------------------------------------------
    // Shape interface in Interop.Excel is shown below...
    //----------------------------------------------------------------------------------------------
    //public interface Shape
    //{
    //    Application Application { get; }
    //    XlCreator Creator { get; }
    //    object Parent { get; }
    //    Adjustments Adjustments { get; }
    //    TextFrame TextFrame { get; }
    //    MsoAutoShapeType AutoShapeType { get; set; }
    //    CalloutFormat Callout { get; }
    //    int ConnectionSiteCount { get; }
    //    MsoTriState Connector { get; }
    //    ConnectorFormat ConnectorFormat { get; }
    //    FillFormat Fill { get; }
    //    GroupShapes GroupItems { get; }
    //    float Height { get; set; }
    //    MsoTriState HorizontalFlip { get; }
    //    float Left { get; set; }
    //    LineFormat Line { get; }
    //    MsoTriState LockAspectRatio { get; set; }
    //    string Name { get; set; }
    //    ShapeNodes Nodes { get; }
    //    float Rotation { get; set; }
    //    PictureFormat PictureFormat { get; }
    //    ShadowFormat Shadow { get; }
    //    TextEffectFormat TextEffect { get; }
    //    ThreeDFormat ThreeD { get; }
    //    float Top { get; set; }
    //    MsoShapeType Type { get; }
    //    MsoTriState VerticalFlip { get; }
    //    object Vertices { get; }
    //    MsoTriState Visible { get; set; }
    //    float Width { get; set; }
    //    int ZOrderPosition { get; }
    //    Hyperlink Hyperlink { get; }
    //    MsoBlackWhiteMode BlackWhiteMode { get; set; }
    //    object DrawingObject { get; }
    //    string OnAction { get; set; }
    //    bool Locked { get; set; }
    //    Range TopLeftCell { get; }
    //    Range BottomRightCell { get; }
    //    XlPlacement Placement { get; set; }
    //    ControlFormat ControlFormat { get; }
    //    LinkFormat LinkFormat { get; }
    //    OLEFormat OLEFormat { get; }
    //    XlFormControl FormControlType { get; }
    //    string AlternativeText { get; set; }
    //    Script Script { get; }
    //    DiagramNode DiagramNode { get; }
    //    MsoTriState HasDiagramNode { get; }
    //    Diagram Diagram { get; }
    //    MsoTriState HasDiagram { get; }
    //    MsoTriState Child { get; }
    //    Shape ParentGroup { get; }
    //    CanvasShapes CanvasItems { get; }
    //    int ID { get; }
    //    Chart Chart { get; }
    //    MsoTriState HasChart { get; }
    //    TextFrame2 TextFrame2 { get; }
    //    MsoShapeStyleIndex ShapeStyle { get; set; }
    //    MsoBackgroundStyleIndex BackgroundStyle { get; set; }
    //    SoftEdgeFormat SoftEdge { get; }
    //    GlowFormat Glow { get; }
    //    ReflectionFormat Reflection { get; }
    //    void Apply();
    //    void Delete();
    //    Shape Duplicate();
    //    void Flip(MsoFlipCmd FlipCmd);
    //    void IncrementLeft(float Increment);
    //    void IncrementRotation(float Increment);
    //    void IncrementTop(float Increment);
    //    void PickUp();
    //    void RerouteConnections();
    //    void ScaleHeight(float Factor, MsoTriState RelativeToOriginalSize, [Optional] object Scale);
    //    void ScaleWidth(float Factor, MsoTriState RelativeToOriginalSize, [Optional] object Scale);
    //    void Select([Optional] object Replace);
    //    void SetShapesDefaultProperties();
    //    ShapeRange Ungroup();
    //    void ZOrder(MsoZOrderCmd ZOrderCmd);
    //    void Copy();
    //    void Cut();
    //    void CopyPicture([Optional] object Appearance, [Optional] object Format);
    //    void CanvasCropLeft(float Increment);
    //    void CanvasCropTop(float Increment);
    //    void CanvasCropRight(float Increment);
    //    void CanvasCropBottom(float Increment);
    //}

    public class Comment
    {
        #region "constants"

        private const int DEFAULT_COMMENT_SIZE_X = 4;
        private const int DEFAULT_COMMENT_SIZE_Y = 3;
        private const int DEFAULT_TOP_LEFT_OFFSET_X_IN_POINT = 100;
        private const int DEFAULT_TOP_LEFT_OFFSET_Y_IN_POINT = 100;
        private const int DEFAULT_BOTTM_RIGHT_OFFSET_X_IN_POINT = 100;
        private const int DEFAULT_BOTTM_RIGHT_OFFSET_Y_IN_POINT = 100;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        /// <param name="Text">コメント文字列</param>
        /// <param name="Visible">コメント常時表示のときtrue</param>
        /// <param name="OwnerAddress">コメントを所有するセルのアドレス</param>
        /// <param name="SizeX">コメントの幅(このセルからいくつ右のセル左端まで広がるか)</param>
        /// <param name="SizeY">コメントの高(このセルからいくつ下のセル上端まで広がるか)</param>
        /// <param name="TopLeftOfsX">左上端のX位置オフセット(単位はEMU(1pixcel=9525emu, 1point=12700emu))</param>
        /// <param name="TopLeftOfsY">左上端のY位置オフセット(単位はEMU(1pixcel=9525emu, 1point=12700emu))</param>
        /// <param name="BottomRightOfsX">右下端のX位置オフセット(単位はEMU(1pixcel=9525emu, 1point=12700emu))</param>
        /// <param name="BottomRightOfsY">右下端のY位置オフセット(単位はEMU(1pixcel=9525emu, 1point=12700emu))</param>
        internal Comment(
                    Range ParentRange, string Author, string Text, bool Visible, CellAddress OwnerAddress, int SizeX, int SizeY,
                    int TopLeftOfsX, int TopLeftOfsY, int BottomRightOfsX, int BottomRightOfsY)
                : this(ParentRange, Author, Text, Visible, OwnerAddress, SizeX, SizeY)
        {
            this.TopLeftOftsetX = TopLeftOfsX;
            this.TopLeftOffsetY = TopLeftOfsY;
            this.BottomRightOffsetX = BottomRightOfsX;
            this.BottomRightOffsetY = BottomRightOfsY;
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        /// <param name="Text">コメント文字列</param>
        /// <param name="Visible">コメント常時表示のときtrue</param>
        /// <param name="OwnerAddress">コメントを所有するセルのアドレス</param>
        /// <param name="SizeX">コメントの幅(このセルからいくつ右のセル左端まで広がるか)</param>
        /// <param name="SizeY">コメントの高(このセルからいくつ下のセル上端まで広がるか)</param>
        internal Comment(
                    Range ParentRange, string Author, string Text, bool Visible, CellAddress OwnerAddress, int SizeX, int SizeY)
                : this(ParentRange, Author, Text, Visible, OwnerAddress)
        {
            this.BottomRightAddress
                = new CellAddress(
                    TopLeftAddress.Row + SizeY, TopLeftAddress.Column + SizeX);
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="ParentRange">親Rangeクラス</param>
        /// <param name="Author">コメント作成者</param>
        /// <param name="Text">コメント文字列</param>
        /// <param name="Visible">コメント常時表示のときtrue</param>
        /// <param name="OwnerAddress">コメントを所有するセルのアドレス</param>
        internal Comment(
                    Range ParentRange, string Author, string Text, bool Visible, CellAddress OwnerAddress)
        {
            //基本情報の保存
            this.Parent = ParentRange;
            this.Author = Author;
            this.Opened = Visible;  //ここではOpenedに保存。VisibleにsetするとApply()が走ってしまうので。
            this.CommentText = Text;
            //コメントを所有するセルのアドレス
            this.OwnerAddress = OwnerAddress;
            //TopLeftAddressはOwnerAddressと等しくなければならない。(NPOIが重複エラーを誤検知する場合があるので)
            this.TopLeftAddress = new CellAddress(OwnerAddress.Row, OwnerAddress.Column);
            //そのセルの右下端までのオフセットを意図し十分に大きな値を指定しておく。
            this.TopLeftOftsetX = DEFAULT_TOP_LEFT_OFFSET_X_IN_POINT * XSSFShape.EMU_PER_POINT;
            this.TopLeftOffsetY = DEFAULT_TOP_LEFT_OFFSET_Y_IN_POINT * XSSFShape.EMU_PER_POINT;
            //右下端セルアドレス
            this.BottomRightAddress
                        = new CellAddress(
                                TopLeftAddress.Row + DEFAULT_COMMENT_SIZE_Y,
                                TopLeftAddress.Column + DEFAULT_COMMENT_SIZE_X);
            //そのセルの最右端/最下端までの拡張を意図し十分に大きな値を指定しておく。
            this.BottomRightOffsetX = DEFAULT_BOTTM_RIGHT_OFFSET_X_IN_POINT * XSSFShape.EMU_PER_POINT;
            this.BottomRightOffsetY = DEFAULT_BOTTM_RIGHT_OFFSET_Y_IN_POINT * XSSFShape.EMU_PER_POINT;
        }

        #endregion

        #region "properties"

        #region "emulated public properties"

        public Application Application { get { return Parent.Application; } }
        public XlCreator Creator { get { return Application.Creator; } }
        public Range Parent { get; }

        /// <summary>
        /// コメントの作者
        /// </summary>
        public string Author { get; }

        /// <summary>
        /// コメントの詳細設定
        /// </summary>
        //public Shape Shape { get; }

        /// <summary>
        /// コメントの常時表示するか否か
        /// </summary>
        public bool Visible
        {   get { return Opened; }
            set
            {
                //値の保存
                Opened = value;
                //セルに適用
                Apply();
            }
        }

        #endregion

        #region "internal properties"

        /// <summary>
        /// AnchorType(未使用)
        /// </summary>
        internal AnchorType Type { get; } = AnchorType.MoveAndResize;

        /// <summary>
        /// コメント文字列
        /// </summary>
        internal string CommentText { get; set; } = string.Empty;

        /// <summary>
        /// 開かれた状態にする(public Visibleの実体)
        /// </summary>
        internal bool Opened { get; private set; }

        /// <summary>
        /// コメントを所有するセルのアドレス。ICommentのAddressおよびRow/Columnにセットする。
        /// </summary>
        internal CellAddress OwnerAddress { get; }

        /// <summary>
        /// コメントの左上端表示位置を持つセルのアドレス。Ownerアドレスと同一値。
        /// </summary>
        internal CellAddress TopLeftAddress { get; }

        /// <summary>
        /// コメントの右下端表示位置を持つセルのアドレス。
        /// </summary>
        internal CellAddress BottomRightAddress { get; }

        /// <summary>
        /// 左上端表示オフセット(X方向)。TopLeftAddress左上端からオフセット値(単位はEMU)。
        /// </summary>
        internal int TopLeftOftsetX { get; }

        /// <summary>
        /// 左上端表示オフセット(Y方向)。TopLeftAddress左上端からオフセット値(単位はEMU)。
        /// </summary>
        internal int TopLeftOffsetY { get; }

        /// <summary>
        /// 右下端表示オフセット(X方向)。BottomRightddress左上端からオフセット値(単位はEMU)。
        /// </summary>
        internal int BottomRightOffsetX { get; }

        /// <summary>
        /// 右下端表示オフセット(Y方向)。BottomRightddress左上端からオフセット値(単位はEMU)。
        /// </summary>
        internal int BottomRightOffsetY { get; }

        #endregion

        #endregion

        #region "methods"

        #region "emulated public methods"

        /// <summary>
        /// コメント文字列を更新する。
        /// </summary>
        /// <param name="Text">コメント文字列</param>
        /// <param name="Start">この文字列を挿入する位置(1開始)</param>
        /// <param name="Overwrite">既存文字列を上書きする場合true</param>
        /// <returns></returns>
        public string Text(object Text = null, object Start = null, object Overwrite = null)
        {
            string CommentString = string.Empty;
            int CommentOffset = 1;  //初期値１(先頭)
            bool CommentOverwrite = true;
            //パラメータ解析
            if (Text is string SafeString)
            {
                CommentString = SafeString;
            }
            if (Start is int SafeInt)
            {
                CommentOffset = SafeInt;
                //１開始のオフセットは最小値＝１
                if (CommentOffset < 1)
                {
                    CommentOffset = 1;
                }
                //１開始のオフセットは最大値＝長さ＋１
                if (CommentOffset > CommentText.Length + 1)
                {
                    CommentOffset = CommentText.Length + 1;
                }
            }
            if (Overwrite is bool SafeBool)
            {
                CommentOverwrite = SafeBool;
            }
            //上書き指定ならば,指定位置以降を今回文字列で置換
            if (CommentOverwrite)
            {
                this.CommentText = CommentText.Substring(0, CommentOffset - 1) + CommentString;
            }
            //上書きでなければ指定位置に今回文字列を挿入
            else
            {
                this.CommentText = CommentText.Substring(0, CommentOffset - 1) + CommentString + CommentText.Substring(CommentOffset - 1);
            }
            //セルに適用
            Apply();
            return this.CommentText;
        }

        /// <summary>
        /// このコメントを削除する。
        /// </summary>
        public void Delete()
        {
            //コメント削除
            Utils.CellUtil.GetCell(Parent.Parent.PoiSheet, OwnerAddress)?.RemoveCellComment();
        }

        /// <summary>
        /// 同一シート内で次のコメントCommentを取得する。シート内最後のコメントでのNext()はnullが通知される。
        /// </summary>
        /// <returns></returns>
        public Comment Next() { return GetComment(); }

        /// <summary>
        /// 同一シート内で前のコメントCommentを取得する。シート内先頭のコメントでのPrevoius()はnullが通知される。
        /// </summary>
        /// <returns></returns>
        public Comment Previous() { return GetComment(-1); }

        #endregion

        #region "internal methods"

        /// <summary>
        /// 現在のプロパティを実Cellに適用する。
        /// </summary>
        internal void Apply()
        {
            //列の取得(なければ生成)
            NPOI.SS.UserModel.ICell Cell = Utils.CellUtil.GetOrCreateCell(Parent.Parent.PoiSheet, OwnerAddress.Row, OwnerAddress.Column);
            //ClientAnchr生成
            IDrawing drawing = Parent.Parent.PoiSheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Parent.Parent.Parent.PoiBook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = TopLeftAddress.Column;
            anchor.Row1 = TopLeftAddress.Row;
            anchor.Col2 = BottomRightAddress.Column;
            anchor.Row2 = BottomRightAddress.Row;
            anchor.Dx1 = this.TopLeftOftsetX;
            anchor.Dy1 = this.TopLeftOffsetY;
            anchor.Dx2 = this.BottomRightOffsetX;
            anchor.Dy2 = this.BottomRightOffsetY;
            //コメント生成
            IComment Comment = drawing.CreateCellComment(anchor);
            //Authorセット
            Comment.Author = this.Author;
            //コメント文字セット(リッチテキスト変換)
            if (Parent.Parent.PoiSheet is HSSFSheet)
            {
                Comment.String = new HSSFRichTextString(this.CommentText);
            }
            else
            {
                Comment.String = new XSSFRichTextString(this.CommentText);
            }
            //表示/非表示セット
            Comment.Visible = this.Visible;
            //セルに適用
            Cell.CellComment = Comment;
        }

        internal Comment GetComment(int Offset = 1)
        {
            Comment RetVal = null;
            //オフセットを-1, +1のいずれかに限定
            int AddressOffset = Offset < 0 ? -1 : 1; 
            //シート内のコメントを取得
            Dictionary<CellAddress, IComment> Comments = Parent.Parent.PoiSheet.GetCellComments();
            //アドレスの昇順ソートしてList化
            List<CellAddress> Adresses = Comments.OrderBy(c => c.Key.Row).ThenBy(c => c.Key.Column).Select(c => c.Key).ToList();
            //自分のIndex取得
            int Index = Adresses.IndexOf(OwnerAddress);
            //そもそも自分が含まれているときのみ処理可能
            //ターゲットがリストの範囲内のときのみ取得可能
            if ( 0 <= (Index + AddressOffset) && (Index + AddressOffset) < Adresses.Count)
            {
                //対象アドレスにセルが実在すること
                ICell Cell = NpoiWrapper.Utils.CellUtil.GetCell(Parent.Parent.PoiSheet, Adresses[Index + AddressOffset]);
                if (Cell != null)
                {
                    //対象アドレスにコメントが存在すること(念のため)
                    if (Cell.CellComment != null)
                    {
                        //Cell.CellCommentからCommentクラスを生成
                        RetVal = new Comment(
                            Parent,
                            Cell.CellComment.Author,
                            Cell.CellComment.String.String,
                            Cell.CellComment.Visible,
                            Adresses[Index + AddressOffset]);
                    }
                }
            }
            return RetVal;
        }

        #endregion

        #endregion
    }
}
