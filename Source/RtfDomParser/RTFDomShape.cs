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
    /// shape element
    /// </summary>
    public class RTFDomShape : RTFDomElement
    {
        private StringAttributeCollection _myExtAttrbutes = new StringAttributeCollection();

        public RTFDomShape()
        {
        }

        /// <summary>
        /// left position
        /// </summary>
        [DefaultValue(0)]
        public int Left { get; set; }

        /// <summary>
        /// top position
        /// </summary>
        [DefaultValue(0)]
        public int Top { get; set; }

        [DefaultValue(0)]
        public int Width { get; set; }

        /// <summary>
        /// height
        /// </summary>
        [DefaultValue(0)]
        public int Height { get; set; }

        /// <summary>
        /// Z index
        /// </summary>
        [DefaultValue(0)]
        public int ZIndex { get; set; }

        /// <summary>
        /// shape id
        /// </summary>
        [DefaultValue(0)]
        public int ShapeId { get; set; }

        /// <summary>
        /// ext attribute
        /// </summary>
        public StringAttributeCollection ExtAttrbutes
        {
            get { return _myExtAttrbutes; }
            set { _myExtAttrbutes = value; }
        }

        public override string ToString()
        {
            return "Shape:Left:" + Left + " Top:" + Top + " Width:" + Width + " Height:" + Height;
        }
    }
}