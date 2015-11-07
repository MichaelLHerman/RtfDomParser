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
    /// <summary>
    /// RTF element container
    /// </summary>
    public class RTFDomElementContainer : RTFDomElement
    {
        public RTFDomElementContainer()
        {
            Name = null;
        }

        /// <summary>
        /// name
        /// </summary>
        [DefaultValue(null)]
        public string Name { get; set; }

        public override string ToString()
        {
            return "Container : " + Name;
        }
    }
}