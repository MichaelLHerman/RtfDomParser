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
using System.Diagnostics;
using System.Text;

namespace RtfDomParser
{
    /// <summary>
    /// RTF node list , this source code evolution from other software.
    /// </summary>
    [DebuggerDisplay("Count={ Count }")]
    public class RTFNodeList : List<RTFNode>
    {
        /// <summary>
        /// get node special keyword
        /// </summary>
        public RTFNode this[string keyWord]
        {
            get
            {
                foreach (var node in this)
                {
                    if (node.Keyword == keyWord)
                        return node;
                }
                return null;
            }
        }

        /// <summary>
        /// get node special type
        /// </summary>
        public RTFNode this[Type t]
        {
            get
            {
                foreach (var node in this)
                {
                    if (t.Equals(node.GetType()))
                        return node;
                }
                return null;
            }
        }

        public string Text
        {
            get
            {
                var myStr = new StringBuilder();
                foreach (var node in this)
                {
                    if (node.Type == RTFNodeType.Text)
                    {
                        myStr.Append(node.Keyword);
                    }
                    else if (node is RTFNodeGroup)
                    {
                        var txt = node.Nodes.Text;
                        if (txt != null)
                            myStr.Append(txt);
                    }
                }
                return myStr.ToString();
            }
        }

        /// <summary>
        /// get node's parameter value special keyword, if can not find , return default value.
        /// </summary>
        /// <param name="key">keyword</param>
        /// <param name="defaultValue">default value</param>
        /// <returns>parameter value</returns>
        public int GetParameter(string key, int defaultValue)
        {
            foreach (var node in this)
            {
                if (node.Keyword == key && node.HasParameter)
                    return node.Parameter;
            }
            return defaultValue;
        }

        /// <summary>
        /// detect whether exist node special keyword in this list
        /// </summary>
        /// <param name="key">keyword</param>
        /// <returns>exist or not</returns>
        public bool ContainsKey(string key)
        {
            return this[key] != null;
        }
    }
}