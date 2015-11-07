/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

namespace RtfDomParser
{
    public class RTFDomTempContainer : RTFDomElement
    {
        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        public override string ToString()
        {
            return "TempContainer";
        }
    }
}