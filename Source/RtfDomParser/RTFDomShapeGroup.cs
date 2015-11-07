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
    public class RTFDomShapeGroup : RTFDomElement
    {
        private StringAttributeCollection _myExtAttrbutes = new StringAttributeCollection();

        /// <summary>
        /// extern attributes
        /// </summary>
        public StringAttributeCollection ExtAttrbutes
        {
            get { return _myExtAttrbutes; }
            set { _myExtAttrbutes = value; }
        }

        public override string ToString()
        {
            return "ShapeGroup";
        }
    }
}