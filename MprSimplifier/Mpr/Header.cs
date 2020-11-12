namespace MprSimplifier.Mpr
{
    class Header
    {
        public string Data { get; private set; }

        public Header(string data)
        {
            Data = data.Trim();
        }
    }
}
