/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System;
using System.Text;

namespace RtfDomParser
{
    /// <summary>
    /// paragraph element
    /// </summary>
    public class RTFDomParagraph : RTFDomElement
    {
        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        internal bool TemplateGenerated = false;

        public bool IsTemplateGenerated
        {
            get { return TemplateGenerated; }
        }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        public override string InnerText
        {
            get { return base.InnerText + Environment.NewLine; }
        }

        public override string ToString()
        {
            var str = new StringBuilder();
            str.Append("Paragraph");
            if (Format != null)
            {
                str.Append("(" + Format.Align + ")");
                if (Format.ListId >= 0)
                {
                    str.Append("ListID:" + Format.ListId);
                }
                //if (this.Format.NumberedList)
                //{
                //    str.Append("(NumberedList)");
                //}
                //else if (this.Format.BulletedList)
                //{
                //    str.Append("(BulletedList)");
                //}
            }

            return str.ToString();
        }
    }
}