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
    /// <summary>
    /// rtf attribute
    /// </summary>

    public class RTFAttribute
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFAttribute()
        {
        }

        private string strName = null;
        /// <summary>
        /// attribute's name
        /// </summary>
        [System.ComponentModel.DefaultValue( null)]
        public string Name
        {
            get
            {
                return strName; 
            }
            set
            {
                strName = value; 
            }
        }

        private int intValue = int.MinValue ;
        /// <summary>
        /// value
        /// </summary>
        [System.ComponentModel.DefaultValue( int.MinValue )]
        public int Value
        {
            get
            {
                return intValue; 
            }
            set
            {
                intValue = value; 
            }
        }
        public override string ToString()
        {
            return strName + "=" + intValue;
        }
    }

    public class RTFAttributeList : List<RTFAttribute>
    {

        public RTFAttributeList()
        {
        }


        public int this[string name]
        {
            get
            {
                foreach (RTFAttribute a in this)
                {
                    if (a.Name == name)
                        return a.Value;
                }
                return int.MinValue;
            }
            set
            {
                foreach (RTFAttribute a in this)
                {
                    if (a.Name == name)
                    {
                        a.Value = value;
                        return;
                    }
                }
                RTFAttribute item = new RTFAttribute();
                item.Name = name;
                item.Value = value;
                Add(item);
            }
        }


        public void Add(string name, int v)
        {
            RTFAttribute item = new RTFAttribute();
            item.Name = name;
            item.Value = v;
            Add(item);
        }

        public void Remove(string name)
        {
            for (int iCount = this.Count - 1; iCount >= 0; iCount--)
            {
                RTFAttribute item = this[iCount];
                if (item.Name == name)
                {
                    RemoveAt(iCount);
                }
            }
        }

        public bool Contains(string name)
        {
            foreach (RTFAttribute a in this)
            {
                if (a.Name == name)
                    return true;
            }
            return false;
        }

        public RTFAttributeList Clone()
        {
            RTFAttributeList list = new RTFAttributeList();
            foreach (RTFAttribute item in this)
            {
                RTFAttribute newItem = new RTFAttribute();
                newItem.Name = item.Name;
                newItem.Value = item.Value;
                list.Add(newItem);
            }
            return list;
        }
    }
}
