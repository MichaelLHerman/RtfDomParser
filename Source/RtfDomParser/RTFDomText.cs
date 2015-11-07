/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.ComponentModel;
using System.Text;

namespace RtfDomParser
{
    public class RTFDomText : RTFDomElement
    {
        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomText()
        {
            Text = null;
            // text element can not contains any child element
            Locked = true;
        }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        /// <summary>
        /// text
        /// </summary>
        [DefaultValue(null)]
        public string Text { get; set; }

        public override string InnerText
        {
            get { return Text; }
        }

        public override string ToString()
        {
            var str = new StringBuilder();
            str.Append("Text");
            if (Format != null)
            {
                if (Format.Hidden)
                {
                    str.Append("(Hidden)");
                }
            }
            str.Append(":" + Text);
            return str.ToString();
        }
    }
}