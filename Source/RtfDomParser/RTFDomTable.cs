/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.Collections.Generic;

namespace RtfDomParser
{
    public class RTFDomTable : RTFDomElement
    {
        private List<RTFDomElement> _myColumns = new List<RTFDomElement>();

        /// <summary>
        /// column list
        /// </summary>
        public List<RTFDomElement> Columns
        {
            get { return _myColumns; }
            set { _myColumns = value; }
        }

        public override string ToString()
        {
            return "Table(Rows:" + Elements.Count + " Columns:" + _myColumns.Count + ")";
        }
    }
}