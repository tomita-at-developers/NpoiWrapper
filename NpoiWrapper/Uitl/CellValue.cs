using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace Developers.NpoiWrapper.Util
{
    public class Cell
    {
        public enum ValueType
        {
            Auto,
            String,
            Formula
        } 

        public ValueType Type { get; set; } = ValueType.Auto;
        public object Value { get; set; } = null;
    }
}
