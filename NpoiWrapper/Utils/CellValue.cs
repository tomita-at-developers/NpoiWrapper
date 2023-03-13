namespace Developers.NpoiWrapper.Utils
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
