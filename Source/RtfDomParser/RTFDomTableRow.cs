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
    /// table row
    /// </summary>
    public class RTFDomTableRow : RTFDomElement
    {
        private int _intLevel = 1;

        private int _intPaddingBottom = int.MinValue;


        private int _intPaddingLeft = int.MinValue;


        private int _intPaddingRight = int.MinValue;

        private int _intPaddingTop = int.MinValue;

        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        public RTFDomTableRow()
        {
            Width = 0;
            Height = 0;
            Header = false;
            IsLastRow = false;
            RowIndex = 0;
        }

        internal List<object> CellSettings { get; set; }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        /// <summary>
        /// document level
        /// </summary>
        [DefaultValue(1)]
        public int Level
        {
            get { return _intLevel; }
            set { _intLevel = value; }
        }

        /// <summary>
        /// row index
        /// </summary>
        [DefaultValue(0)]
        internal int RowIndex { get; set; }

        /// <summary>
        /// is the last row
        /// </summary>
        [DefaultValue(false)]
        public bool IsLastRow { get; set; }

        /// <summary>
        /// is header row
        /// </summary>
        [DefaultValue(false)]
        public bool Header { get; set; }

        /// <summary>
        /// height
        /// </summary>
        [DefaultValue(0)]
        public int Height { get; set; }

        /// <summary>
        /// padding left
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int PaddingLeft
        {
            get { return _intPaddingLeft; }
            set { _intPaddingLeft = value; }
        }

        /// <summary>
        /// top padding
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int PaddingTop
        {
            get { return _intPaddingTop; }
            set { _intPaddingTop = value; }
        }

        /// <summary>
        /// right padding
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int PaddingRight
        {
            get { return _intPaddingRight; }
            set { _intPaddingRight = value; }
        }

        /// <summary>
        /// bottom padding
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int PaddingBottom
        {
            get { return _intPaddingBottom; }
            set { _intPaddingBottom = value; }
        }

        /// <summary>
        /// width
        /// </summary>
        [DefaultValue(0)]
        public int Width { get; set; }

        public override string ToString()
        {
            return "Row " + RowIndex;
        }
    }
}