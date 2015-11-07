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
    public class RTFDomShapeGroup : RTFDomElement
    {
        public RTFDomShapeGroup()
        {
        }

        private StringAttributeCollection myExtAttrbutes = new StringAttributeCollection();
        /// <summary>
        /// extern attributes
        /// </summary>
        public StringAttributeCollection ExtAttrbutes
        {
            get
            {
                return myExtAttrbutes;
            }
            set
            {
                myExtAttrbutes = value;
            }
        }

        public override string ToString()
        {
            return "ShapeGroup";
        }
    }
}
