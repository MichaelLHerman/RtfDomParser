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
    /// <summary>
    /// rtf bookmark
    /// </summary>

    public class RTFDomBookmark : RTFDomElement
    {
        /// <summary>
        /// name
        /// </summary>
        [System.ComponentModel.DefaultValue( null )]
        public string Name { get; set; }

        public override string ToString()
        {
            return "BookMark:" + Name;
        }
    }
}
