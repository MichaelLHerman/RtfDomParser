/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System.ComponentModel;

namespace RtfDomParser
{
    public class RTFDomHeader : RTFDomElement
    {
        private HeaderFooterStyle _style = HeaderFooterStyle.AllPages;

        [DefaultValue(HeaderFooterStyle.AllPages)]
        public HeaderFooterStyle Style
        {
            get { return _style; }
            set { _style = value; }
        }

        public bool HasContentElement
        {
            get { return RTFUtil.HasContentElement(this); }
        }

        public override string ToString()
        {
            return "Header " + Style;
        }
    }

    public class RTFDomFooter : RTFDomElement
    {
        private HeaderFooterStyle _style = HeaderFooterStyle.AllPages;

        [DefaultValue(HeaderFooterStyle.AllPages)]
        public HeaderFooterStyle Style
        {
            get { return _style; }
            set { _style = value; }
        }


        public bool HasContentElement
        {
            get { return RTFUtil.HasContentElement(this); }
        }

        public override string ToString()
        {
            return "Footer " + Style;
        }
    }

    public enum HeaderFooterStyle
    {
        AllPages,
        LeftPages,
        RightPages,
        FirstPage
    }
}