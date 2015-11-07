namespace RtfDomParser
{
    public class RTFRawLayerInfo
    {
        private int _ucValue;

        public RTFRawLayerInfo()
        {
            UcValueCount = 0;
        }

        public int UcValue
        {
            get { return _ucValue; }
            set
            {
                _ucValue = value;
                UcValueCount = 0;
            }
        }

        public int UcValueCount { get; set; }

        public bool CheckUcValueCount()
        {
            UcValueCount--;
            return UcValueCount < 0;
        }

        public RTFRawLayerInfo Clone()
        {
            return (RTFRawLayerInfo) MemberwiseClone();
        }
    }
}