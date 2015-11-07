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
    public class RTFDomPageBreak : RTFDomElement
    {
        public RTFDomPageBreak()
        {
            Locked = true;
        }

        public override string InnerText
        {
            get { return ""; }
        }

        public override string ToString()
        {
            return "page";
        }
    }
}