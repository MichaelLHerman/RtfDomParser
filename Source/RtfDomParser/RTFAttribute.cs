/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace RtfDomParser
{
    /// <summary>
    /// rtf attribute
    /// </summary>
    public class RTFAttribute
    {
        private int _intValue = int.MinValue;

        public RTFAttribute()
        {
            Name = null;
        }

        /// <summary>
        /// attribute's name
        /// </summary>
        [DefaultValue(null)]
        public string Name { get; set; }

        /// <summary>
        /// value
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int Value
        {
            get { return _intValue; }
            set { _intValue = value; }
        }

        public override string ToString()
        {
            return Name + "=" + _intValue;
        }
    }

    public class RTFAttributeList : List<RTFAttribute>
    {
        public int this[string name]
        {
            get
            {
                foreach (var a in this)
                {
                    if (a.Name == name)
                        return a.Value;
                }
                return int.MinValue;
            }
            set
            {
                foreach (var a in this)
                {
                    if (a.Name == name)
                    {
                        a.Value = value;
                        return;
                    }
                }
                var item = new RTFAttribute();
                item.Name = name;
                item.Value = value;
                Add(item);
            }
        }


        public void Add(string name, int v)
        {
            var item = new RTFAttribute();
            item.Name = name;
            item.Value = v;
            Add(item);
        }

        public void Remove(string name)
        {
            for (var iCount = Count - 1; iCount >= 0; iCount--)
            {
                var item = this[iCount];
                if (item.Name == name)
                {
                    RemoveAt(iCount);
                }
            }
        }

        public bool Contains(string name)
        {
            return this.Any(a => a.Name == name);
        }

        public RTFAttributeList Clone()
        {
            var list = new RTFAttributeList();
            foreach (var item in this)
            {
                var newItem = new RTFAttribute();
                newItem.Name = item.Name;
                newItem.Value = item.Value;
                list.Add(newItem);
            }
            return list;
        }
    }
}