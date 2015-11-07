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

namespace RtfDomParser
{
    /// <summary>
    /// string attribute
    /// </summary>
    public class StringAttribute
    {
        public StringAttribute()
        {
            Value = null;
            Name = null;
        }

        /// <summary>
        /// name
        /// </summary>
        [DefaultValue(null)]
        public string Name { get; set; }

        /// <summary>
        /// value
        /// </summary>
        [DefaultValue(null)]
        public string Value { get; set; }

        public override string ToString()
        {
            return Name + "=" + Value;
        }
    }


    public class StringAttributeCollection : List<StringAttribute>
    {
        public string this[string name]
        {
            get
            {
                foreach (var attr in this)
                {
                    if (attr.Name == name)
                    {
                        return attr.Value;
                    }
                }
                return null;
            }
            set
            {
                foreach (var item in this)
                {
                    if (item.Name == name)
                    {
                        if (value == null)
                            Remove(item);
                        else
                            item.Value = value;
                        return;
                    }
                }
                if (value != null)
                {
                    var newItem = new StringAttribute();
                    newItem.Name = name;
                    newItem.Value = value;
                    Add(newItem);
                }
            }
        }
    }
}