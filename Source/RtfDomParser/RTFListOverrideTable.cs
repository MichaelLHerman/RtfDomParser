using System.Collections.Generic;

namespace RtfDomParser
{
    public class RTFListOverrideTable : List<RTFListOverride>
    {
        public RTFListOverride GetById(int id)
        {
            foreach (var item in this)
            {
                if (item.Id == id)
                {
                    return item;
                }
            }
            return null;
        }
    }

    public class RTFListOverride
    {
        private int _id = 1;

        public RTFListOverride()
        {
            ListOverriedCount = 0;
            ListId = 0;
        }

        public int ListId { get; set; }

        public int ListOverriedCount { get; set; }

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public override string ToString()
        {
            return "ID:" + Id + " ListID:" + ListId + " Count:" + ListOverriedCount;
        }
    }
}