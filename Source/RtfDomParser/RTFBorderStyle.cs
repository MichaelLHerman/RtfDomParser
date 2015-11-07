/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System.ComponentModel;

namespace RtfDomParser
{
    public class RTFBorderStyle
    {
        private Color _color = Color.Black;

        private DashStyle _style = DashStyle.Solid;

        public RTFBorderStyle()
        {
            Thickness = false;
            Bottom = false;
            Right = false;
            Top = false;
            Left = false;
        }

        [DefaultValue(false)]
        public bool Left { get; set; }

        [DefaultValue(false)]
        public bool Top { get; set; }

        [DefaultValue(false)]
        public bool Right { get; set; }

        [DefaultValue(false)]
        public bool Bottom { get; set; }

        [DefaultValue(DashStyle.Solid)]
        public DashStyle Style
        {
            get { return _style; }
            set { _style = value; }
        }

        [DefaultValue(typeof (Color), "Black")]
        public Color Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public bool Thickness { get; set; }


        public RTFBorderStyle Clone()
        {
            var b = new RTFBorderStyle
            {
                Bottom = Bottom,
                _color = _color,
                Left = Left,
                Right = Right,
                _style = _style,
                Top = Top,
                Thickness = Thickness
            };
            return b;
        }

        public bool EqualsValue(RTFBorderStyle b)
        {
            if (b == this)
            {
                return true;
            }
            if (b == null)
            {
                return false;
            }
            return b.Bottom == Bottom && b._color == _color && b.Left == Left && b.Right == Right && b._style == _style && b.Top == Top && b.Thickness == Thickness;
        }
    }
}