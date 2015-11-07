/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System;
using System.Collections.Generic;
using System.Text;

namespace RtfDomParser
{
    public class RTFDomTable : RTFDomElement
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomTable()
        {
        }

        private List<RTFDomElement> myColumns = new List<RTFDomElement>();
        /// <summary>
        /// column list
        /// </summary>
        public List<RTFDomElement> Columns
        {
            get
            {
                return myColumns; 
            }
            set
            {
                myColumns = value; 
            }
        }

        public override string ToString()
        {
            return "Table(Rows:" + this.Elements.Count + " Columns:" + myColumns.Count + ")";
        }
    }
}
