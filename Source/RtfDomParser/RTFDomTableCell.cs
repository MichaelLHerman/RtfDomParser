/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.ComponentModel;
using System.Xml.Serialization;

namespace RtfDomParser
{
    public class RTFDomTableCell : RTFDomElement
    {
        private bool _bolMultiline = true;

        private int _intColSpan = 1;

        //private Color intBorderColor = Color.Black;
        ///// <summary>
        ///// border color
        ///// </summary>
        //[System.ComponentModel.DefaultValue(typeof(Color), "Black")]
        //public Color BorderColor
        //{
        //    get
        //    {
        //        return intBorderColor;
        //    }
        //    set
        //    {
        //        intBorderColor = value;
        //    }
        //}

        //private Color intBackColor = Color.Transparent;
        ///// <summary>
        ///// back color
        ///// </summary>
        //[System.ComponentModel.DefaultValue(typeof(Color), "Transparent")]
        //public Color BackColor
        //{
        //    get
        //    {
        //        return intBackColor;
        //    }
        //    set
        //    {
        //        intBackColor = value;
        //    }
        //}


        private int _intPaddingBottom = int.MinValue;

        private int _intPaddingLeft = int.MinValue;

        private int _intPaddingRight = int.MinValue;

        private int _intPaddingTop = int.MinValue;

        private int _intRowSpan = 1;

        //private bool bolLeftBorder = false;
        //[System.ComponentModel.DefaultValue(false)]
        //public bool LeftBorder
        //{
        //    get
        //    {
        //        return bolLeftBorder;
        //    }
        //    set
        //    {
        //        bolLeftBorder = value;
        //    }
        //}

        //private bool bolTopBorder = false;
        //[System.ComponentModel.DefaultValue(false)]
        //public bool TopBorder
        //{
        //    get
        //    {
        //        return bolTopBorder;
        //    }
        //    set
        //    {
        //        bolTopBorder = value;
        //    }
        //}


        //private bool bolRightBorder = false;
        //[System.ComponentModel.DefaultValue(false)]
        //public bool RightBorder
        //{
        //    get
        //    {
        //        return bolRightBorder;
        //    }
        //    set
        //    {
        //        bolRightBorder = value;
        //    }
        //}

        //private bool bolBottomBorder = false;

        //[System.ComponentModel.DefaultValue(false)]
        //public bool BottomBorder
        //{
        //    get
        //    {
        //        return bolBottomBorder;
        //    }
        //    set
        //    {
        //        bolBottomBorder = value;
        //    }
        //}

        private RTFVerticalAlignment _intVerticalAlignment = RTFVerticalAlignment.Top;

        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomTableCell()
        {
            OverrideCell = null;
            Height = 0;
            Width = 0;
            Left = 0;
            _myFormat.BorderWidth = 1;
        }

        /// <summary>
        /// row span
        /// </summary>
        [DefaultValue(1)]
        public int RowSpan
        {
            get { return _intRowSpan; }
            set { _intRowSpan = value; }
        }

        /// <summary>
        /// col span
        /// </summary>
        [DefaultValue(1)]
        public int ColSpan
        {
            get { return _intColSpan; }
            set { _intColSpan = value; }
        }

        /// <summary>
        /// left padding
        /// </summary>
        [DefaultValue(int.MinValue)]
        public int PaddingLeft
        {
            get { return _intPaddingLeft; }
            set { _intPaddingLeft = value; }
        }

        /// <summary>
        /// left padding in fact
        /// </summary>
        [XmlIgnore]
        public int RuntimePaddingLeft
        {
            get
            {
                if (_intPaddingLeft != int.MinValue)
                {
                    return _intPaddingLeft;
                }
                if (Parent != null)
                {
                    var p = ((RTFDomTableRow) Parent).PaddingLeft;
                    if (p != int.MinValue)
                    {
                        return p;
                    }
                }
                return 0;
            }
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
        /// top padding in fact
        /// </summary>
        [XmlIgnore]
        public int RuntimePaddingTop
        {
            get
            {
                if (_intPaddingTop != int.MinValue)
                {
                    return _intPaddingTop;
                }
                if (Parent != null)
                {
                    var p = ((RTFDomTableRow) Parent).PaddingTop;
                    if (p != int.MinValue)
                    {
                        return p;
                    }
                }
                return 0;
            }
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
        /// right padding in fact
        /// </summary>
        [XmlIgnore]
        public int RuntimePaddingRight
        {
            get
            {
                if (_intPaddingRight != int.MinValue)
                {
                    return _intPaddingRight;
                }
                if (Parent != null)
                {
                    var p = ((RTFDomTableRow) Parent).PaddingRight;
                    if (p != int.MinValue)
                    {
                        return p;
                    }
                }
                return 0;
            }
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
        /// bottom padding in fact
        /// </summary>
        [XmlIgnore]
        public int RuntimePaddingBottom
        {
            get
            {
                if (_intPaddingBottom != int.MinValue)
                {
                    return _intPaddingBottom;
                }
                if (Parent != null)
                {
                    var p = ((RTFDomTableRow) Parent).PaddingBottom;
                    if (p != int.MinValue)
                    {
                        return p;
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// vertial alignment
        /// </summary>
        [DefaultValue(RTFVerticalAlignment.Top)]
        public RTFVerticalAlignment VerticalAlignment
        {
            get { return _intVerticalAlignment; }
            set { _intVerticalAlignment = value; }
        }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        /// <summary>
        /// allow multiline
        /// </summary>
        [DefaultValue(false)]
        public bool Multiline
        {
            get { return _bolMultiline; }
            set { _bolMultiline = value; }
        }

        /// <summary>
        /// left position
        /// </summary>
        [DefaultValue(0)]
        public int Left { get; set; }

        /// <summary>
        /// width
        /// </summary>
        [DefaultValue(0)]
        public int Width { get; set; }

        /// <summary>
        /// height
        /// </summary>
        [DefaultValue(0)]
        public int Height { get; set; }

        /// <summary>
        /// this cell merged by another cell which this property specify
        /// </summary>
        [XmlIgnore]
        public RTFDomTableCell OverrideCell { get; set; }

        public override string ToString()
        {
            if (OverrideCell == null)
            {
                if (_intRowSpan != 1 || _intColSpan != 1)
                {
                    return "Cell: RowSpan:" + _intRowSpan + " ColSpan:" + _intColSpan + " Width:" + Width;
                }
                return "Cell:Width:" + Width;
            }
            return "Cell:Overrided";
        }
    }
}