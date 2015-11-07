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
    public class RTFDomTableColumn : RTFDomElement
    {
        public RTFDomTableColumn()
        {
            Width = 0;
        }

        /// <summary>
        /// width
        /// </summary>
        [DefaultValue(0)]
        public int Width { get; set; }

        public override string ToString()
        {
            return "Column";
        }
    }
}