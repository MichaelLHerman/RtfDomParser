using System;
using System.Collections.Generic;

namespace RtfDomParser
{
    public class RTFListTable : List<RTFList>
    {
        public RTFList GetById(int id)
        {
            foreach (var list in this)
            {
                if (list.ListId == id)
                {
                    return list;
                }
            }
            return null;
        }
    }

    public class RTFList
    {
        private LevelNumberType _levelNfc = LevelNumberType.None;

        private int _levelStartAt = 1;

        public RTFList()
        {
            LevelText = null;
            FontName = null;
            LevelFollow = 0;
            LevelJc = 0;
            ListStyleName = null;
            ListName = null;
            ListHybrid = false;
            ListSimple = false;
            ListTemplateId = 0;
            ListId = 0;
        }

        public int ListId { get; set; }

        public int ListTemplateId { get; set; }

        public bool ListSimple { get; set; }

        public bool ListHybrid { get; set; }

        public string ListName { get; set; }

        public string ListStyleName { get; set; }

        public int LevelStartAt
        {
            get { return _levelStartAt; }
            set { _levelStartAt = value; }
        }

        public LevelNumberType LevelNfc
        {
            get { return _levelNfc; }
            set { _levelNfc = value; }
        }

        public int LevelJc { get; set; }

        public int LevelFollow { get; set; }

        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; }

        public string LevelText { get; set; }

        public override string ToString()
        {
            if (LevelNfc == LevelNumberType.Bullet)
            {
                var text = "ID:" + ListId + "   Bullet:";
                if (string.IsNullOrEmpty(LevelText) == false)
                {
                    text = text + "(" + Convert.ToString((short) LevelText[0]) + ")";
                }
                return text;
            }
            return "ID:" + ListId + " " + LevelNfc + " Start:" + LevelStartAt;
        }
    }
}