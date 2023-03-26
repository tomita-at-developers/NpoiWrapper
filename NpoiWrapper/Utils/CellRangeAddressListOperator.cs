using NPOI.SS.Util;

namespace Developers.NpoiWrapper.Utils
{
    internal class CellRangeAddressListOperator
    {
        #region "fields"

        /// <summary>
        /// Baseアドレスリスト
        /// </summary>
        CellRangeAddressList _BaseList;
        /// <summary>
        /// Targetアドレスリスト
        /// </summary>
        CellRangeAddressList _TargetList;

        #endregion

        #region "constructors"

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="BaseAddressList">Baseアドレスリスト</param>
        /// <param name="TargetAddressList">Targetアドレスリスト</param>
        public CellRangeAddressListOperator(CellRangeAddressList BaseAddressList, CellRangeAddressList TargetAddressList)
        {
            //入力情報保存
            this._BaseList = BaseAddressList;
            this._TargetList = TargetAddressList;
            //アドレスリストの解析
            Analyze();
        }

        #endregion

        #region "properties"

        /// <summary>
        /// Baseアドレスリスト
        /// </summary>
        public CellRangeAddressList Base { get { return _BaseList; } }
        public string BaseString { get { return FormatAsString(_BaseList); } }

        /// <summary>
        /// Targetアドレスリスト
        /// </summary>
        public CellRangeAddressList Target { get { return _TargetList; } }
        public string TargetString { get { return FormatAsString(_TargetList); } }

        /// <summary>
        /// Baseアドレスリストにしか含まれない部分のCellRangeAddressList
        /// </summary>
        public CellRangeAddressList BaseRemainder { get; } = new CellRangeAddressList();
        public string BaseRemainderString { get { return FormatAsString(BaseRemainder); } }

        /// <summary>
        /// BaseとTargetで重なり合う部分のCellRangeAddressList
        /// </summary>
        public CellRangeAddressList Overlapping { get; } = new CellRangeAddressList();
        public string OverlappingString { get { return FormatAsString(Overlapping); } }

        /// <summary>
        /// Targetアドレスリストにしか含まれない部分のCellRangeAddressList
        /// </summary>
        public CellRangeAddressList TargetRemainder { get; } = new CellRangeAddressList();
        public string TargetRemainderString { get { return FormatAsString(TargetRemainder); } }

        #endregion

        #region "methods"

        /// <summary>
        /// アドレスリストの比較解析
        /// </summary>
        private void Analyze()
        {
            //Targetループ
            for (int t = 0; t < _TargetList.CountRanges(); t++)
            {
                CellRangeAddress Target = _TargetList.GetCellRangeAddress(t);
                //Baseループ
                for (int b = 0; b < _BaseList.CountRanges(); b++)
                {
                    int FirstRow, LastRow, FirstColumn, LastColumn;
                    CellRangeAddress Base = _BaseList.GetCellRangeAddress(b);
                    //Oerlappingを求める
                    GetOverlapping(Base.FirstRow, Base.LastRow, Target.FirstRow, Target.LastRow, out FirstRow, out LastRow);
                    GetOverlapping(Base.FirstColumn, Base.LastColumn, Target.FirstColumn, Target.LastColumn, out FirstColumn, out LastColumn);
                    if (FirstRow >= 0 && LastRow >= 0 && FirstColumn >= 0 && LastColumn >= 0)
                    {
                        Overlapping.AddCellRangeAddress(FirstRow, FirstColumn, LastRow, LastColumn);
                    }
                    //BaseRemainderを求める
                    //固有行(先頭側)の取得
                    GetRemainderLower(Base.FirstRow, Base.LastRow, Target.FirstRow, Target.LastRow, out FirstRow, out LastRow);
                    //固有行があるならば自カラムすべてが目的のRange
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        BaseRemainder.AddCellRangeAddress(FirstRow, Base.FirstColumn, LastRow, Base.LastColumn);
                    }
                    //重なり合う行の取得
                    GetOverlapping(Base.FirstRow, Base.LastRow, Target.FirstRow, Target.LastRow, out FirstRow, out LastRow);
                    //重なり合う行があるならば固有レンジを求める
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        //固有列(先頭側)の取得
                        GetRemainderLower(Base.FirstColumn, Base.LastColumn, Target.FirstColumn, Target.LastColumn, out FirstColumn, out LastColumn);
                        if (FirstColumn >= 0 && LastColumn >= 0)
                        {
                            BaseRemainder.AddCellRangeAddress(FirstRow, FirstColumn, LastRow, LastColumn);
                        }
                        //固有列(末尾側)の取得
                        GetRemainderUpper(Base.FirstColumn, Base.LastColumn, Target.FirstColumn, Target.LastColumn, out FirstColumn, out LastColumn);
                        if (FirstColumn >= 0 && LastColumn >= 0)
                        {
                            BaseRemainder.AddCellRangeAddress(FirstRow, FirstColumn, LastRow, LastColumn);
                        }
                    }
                    //固有行(末尾側)の取得
                    GetRemainderUpper(Base.FirstRow, Base.LastRow, Target.FirstRow, Target.LastRow, out FirstRow, out LastRow);
                    //固有行があるならば自カラムすべてが目的のRange
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        BaseRemainder.AddCellRangeAddress(FirstRow, Base.FirstColumn, LastRow, Base.LastColumn);
                    }
                    //TargetRemainderを求める
                    //固有行(先頭側)の取得
                    GetRemainderLower(Target.FirstRow, Target.LastRow, Base.FirstRow, Base.LastRow, out FirstRow, out LastRow);
                    //固有行があるならば自カラムすべてが目的のRange
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        TargetRemainder.AddCellRangeAddress(FirstRow, Target.FirstColumn, LastRow, Target.LastColumn);
                    }
                    //重なり合う行の取得
                    GetOverlapping(Target.FirstRow, Target.LastRow, Base.FirstRow, Base.LastRow, out FirstRow, out LastRow);
                    //重なり合う行があるならば固有レンジを求める
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        //固有列(先頭側)の取得
                        GetRemainderLower(Target.FirstColumn, Target.LastColumn, Base.FirstColumn, Base.LastColumn, out FirstColumn, out LastColumn);
                        if (FirstColumn >= 0 && LastColumn >= 0)
                        {
                            TargetRemainder.AddCellRangeAddress(FirstRow, FirstColumn, LastRow, LastColumn);
                        }
                        //固有列(末尾側)の取得
                        GetRemainderUpper(Target.FirstColumn, Target.LastColumn, Base.FirstColumn, Base.LastColumn, out FirstColumn, out LastColumn);
                        if (FirstColumn >= 0 && LastColumn >= 0)
                        {
                            TargetRemainder.AddCellRangeAddress(FirstRow, FirstColumn, LastRow, LastColumn);
                        }
                    }
                    //固有行(末尾側)の取得
                    GetRemainderUpper(Target.FirstRow, Target.LastRow, Base.FirstRow, Base.LastRow, out FirstRow, out LastRow);
                    //固有行があるならば自カラムすべてが目的のRange
                    if (FirstRow >= 0 && LastRow >= 0)
                    {
                        TargetRemainder.AddCellRangeAddress(FirstRow, Target.FirstColumn, LastRow, Target.LastColumn);
                    }
                }
            }
        }

        /// <summary>
        /// 一次元でBaseとTargetの重なり合う部分を求める
        /// </summary>
        /// <param name="BaseFirst">Bsseの始点</param>
        /// <param name="BaseLast">Baseの終点</param>
        /// <param name="TargetFirst">Targetの始点</param>
        /// <param name="TargetLast">Targetの終点</param>
        /// <param name="First">ヒットした領域の始点(なければ-1)</param>
        /// <param name="Last">ヒットした領域の終点(なければ-1)</param>
        private void GetOverlapping(int BaseFirst, int BaseLast, int TargetFirst, int TargetLast, out int First, out int Last)
        {
            First = -1;
            Last = -1;
            //Target先行
            if (TargetFirst <= BaseFirst)
            {
                //Baseの始点はTargetの中にある
                if (BaseFirst <= TargetLast)
                {
                    //Baseの始点が始点
                    First = BaseFirst;
                    //先に終る方が終点
                    if (BaseLast <= TargetLast)
                    {
                        Last = BaseLast;
                    }
                    else
                    {
                        Last = TargetLast;
                    }
                }
            }
            //Base先行
            else
            {
                //Targetの始点はBaseの中にある
                if (TargetFirst <= BaseLast)
                {
                    //Targetの始点が始点
                    First = TargetFirst;
                    //先に終る方が終点
                    if (BaseLast <= TargetLast)
                    {
                        Last = BaseLast;
                    }
                    else
                    {
                        Last = TargetLast;
                    }
                }
            }
        }

        /// <summary>
        /// 一次元でBaseにのみ含まれる部分を求める。BaseがTargetより先に開始する場合にヒットする。
        /// </summary>
        /// <param name="BaseFirst">Bsseの始点</param>
        /// <param name="BaseLast">Baseの終点</param>
        /// <param name="TargetFirst">Targetの始点</param>
        /// <param name="TargetLast">Targetの終点</param>
        /// <param name="First">ヒットした領域の始点(なければ-1)</param>
        /// <param name="Last">ヒットした領域の終点(なければ-1)</param>
        private void GetRemainderLower(int BaseFirst, int BaseLast, int TargetFirst, int TargetLast, out int First, out int Last)
        {
            First = -1;
            Last = -1;
            //Base先行の時のみ
            if (BaseFirst < TargetFirst)
            {
                First = BaseFirst;
                //Baseの始点はTargetの中にある
                if (TargetFirst <= BaseLast)
                {
                    Last = TargetFirst - 1;
                }
                else
                {
                    Last = BaseLast;
                }
            }
        }

        /// <summary>
        /// 一次元でBaseにのみ含まれる部分を求める。BaseがTargetより後に終了する場合にヒットする。
        /// </summary>
        /// <param name="BaseFirst">Bsseの始点</param>
        /// <param name="BaseLast">Baseの終点</param>
        /// <param name="TargetFirst">Targetの始点</param>
        /// <param name="TargetLast">Targetの終点</param>
        /// <param name="First">ヒットした領域の始点(なければ-1)</param>
        /// <param name="Last">ヒットした領域の終点(なければ-1)</param>
        private void GetRemainderUpper(int BaseFirst, int BaseLast, int TargetFirst, int TargetLast, out int First, out int Last)
        {
            First = -1;
            Last = -1;
            //Target先行終了時のみ
            if (TargetLast < BaseLast)
            {
                Last = BaseLast;
                //Baseの始点はTargetの中にある
                if (BaseFirst <= TargetLast)
                {
                    First = TargetLast + 1;
                }
                else
                {
                    First = BaseFirst;
                }
            }
        }

        /// <summary>
        /// A1形式のアドレスリストを取得する
        /// </summary>
        /// <param name="Address"></param>
        /// <returns></returns>
        private string FormatAsString(CellRangeAddressList Address)
        {
            string RetVal = string.Empty;
            for (int a = 0; a < Address.CountRanges(); a++)
            {
                RetVal += Address.GetCellRangeAddress(a).FormatAsString() + " ";
            }
            return RetVal.Trim();
        }

    }

    #endregion
}
