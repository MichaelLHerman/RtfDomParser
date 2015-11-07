/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.Collections.Generic;
using System.Diagnostics;

namespace RtfDomParser
{
    /// <summary>
    /// rtf color table
    /// </summary>
    [DebuggerDisplay("Count={Count}")]
    public class RTFColorTable
    {
        private bool _bolCheckValueExistWhenAdd = true;

        private readonly List<Color> _myItems = new List<Color>();

        /// <summary>
        /// get color at special index
        /// </summary>
        public Color this[int index]
        {
            get { return _myItems[index]; }
        }

        /// <summary>
        /// check color value exist when add color to list
        /// </summary>
        public bool CheckValueExistWhenAdd
        {
            get { return _bolCheckValueExistWhenAdd; }
            set { _bolCheckValueExistWhenAdd = value; }
        }

        public int Count
        {
            get { return _myItems.Count; }
        }

        /// <summary>
        /// get color at special index , if index out of range , return default color
        /// </summary>
        /// <param name="index">index</param>
        /// <param name="defaultValue">default value</param>
        /// <returns>color value</returns>
        public Color GetColor(int index, Color defaultValue)
        {
            index --;
            if (index >= 0 && index < _myItems.Count)
            {
                return _myItems[index];
            }
            return defaultValue;
        }

        /// <summary>
        /// add color to list
        /// </summary>
        /// <param name="c">new color value</param>
        public void Add(Color c)
        {
            if (c.IsEmpty)
                return;
            if (c.A == 0)
                return;

            if (c.A != 255)
            {
                c = Color.FromArgb(255, c);
            }

            if (_bolCheckValueExistWhenAdd)
            {
                if (IndexOf(c) < 0)
                {
                    _myItems.Add(c);
                }
            }
            else
            {
                _myItems.Add(c);
            }
        }

        /// <summary>
        /// delete special color
        /// </summary>
        /// <param name="c">color value</param>
        public void Remove(Color c)
        {
            var index = IndexOf(c);
            if (index >= 0)
                _myItems.RemoveAt(index);
        }

        /// <summary>
        /// get color index
        /// </summary>
        /// <param name="c">color</param>
        /// <returns>index , if not found , return -1</returns>
        public int IndexOf(Color c)
        {
            if (c.A == 0)
            {
                return -1;
            }
            if (c.A != 255)
            {
                c = Color.FromArgb(255, c);
            }
            for (var iCount = 0; iCount < _myItems.Count; iCount++)
            {
                var color = _myItems[iCount];
                if (color.ToArgb() == c.ToArgb())
                {
                    return iCount;
                }
            }
            return -1;
        }

        public void Clear()
        {
            _myItems.Clear();
        }

        public void Write(RTFWriter writer)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword(RTFConsts.Colortbl);
            writer.WriteRaw(";");
            for (var iCount = 0; iCount < _myItems.Count; iCount ++)
            {
                var c = _myItems[iCount];
                writer.WriteKeyword("red" + c.R);
                writer.WriteKeyword("green" + c.G);
                writer.WriteKeyword("blue" + c.B);
                writer.WriteRaw(";");
            }
            writer.WriteEndGroup();
        }

        public RTFColorTable Clone()
        {
            var table = new RTFColorTable();
            for (var iCount = 0; iCount < _myItems.Count; iCount++)
            {
                var c = _myItems[iCount];
                table._myItems.Add(c);
            }
            return table;
        }
    }
}