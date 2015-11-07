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
    /// <summary>
    /// RTF Document format information
    /// </summary>
    public class DocumentFormatInfo
    {
        //private bool bolBulletedList = false;
        ///// <summary>
        ///// list in bulleted style
        ///// </summary>
        //[DefaultValue(false)]
        //public bool BulletedList
        //{
        //    get
        //    {
        //        return bolBulletedList;
        //    }
        //    set
        //    {
        //        bolBulletedList = value;
        //    }
        //}

        //private bool bolNumberedList = false;
        ///// <summary>
        ///// list in numbered style
        ///// </summary>
        //[DefaultValue(false)]
        //public bool NumberedList
        //{
        //    get
        //    {
        //        return bolNumberedList;
        //    }
        //    set
        //    {
        //        bolNumberedList = value;
        //    }
        //}

        private bool _bolNoWwrap = true;

        private bool _bolSubscript;

        private bool _bolSuperscript;

        private DashStyle _borderStyle = DashStyle.Solid;

        private float _fFontSize = 12f;

        private RTFAlignment _intAlign = RTFAlignment.Left;

        private Color _intBackColor = Color.Empty;

        private Color _intBorderColor
            = Color.Black;

        private int _intStandTabWidth = 100;

        private Color _intTextColor = Color.Black;

        private int _listId = -1;


        private string _strFontName = "Microsoft Sans Serif";

        //public Color BorderColor = Color.Empty;

        /// <summary>
        /// nest level in native rtf document
        /// </summary>
        public int NativeLevel;

        internal bool ReadText = true;

        public DocumentFormatInfo()
        {
            Link = null;
            Hidden = false;
            Strikeout = false;
            Underline = false;
            Italic = false;
            Bold = false;
            PageBreak = false;
            SpacingAfter = 0;
            SpacingBefore = 0;
            MultipleLineSpacing = false;
            LineSpacing = 0;
            Spacing = 0;
            LeftIndent = 0;
            ParagraphFirstLineIndent = 0;
            Multiline = false;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderWidth = 0;
            BottomBorder = false;
            RightBorder = false;
            TopBorder = false;
            LeftBorder = false;
            Parent = null;
        }

        /// <summary>
        /// If this instance is create by Clone , return the parent instance
        /// </summary>
        public DocumentFormatInfo Parent { get; private set; }

        /// <summary>
        /// Display left border line
        /// </summary>
        [DefaultValue(false)]
        public bool LeftBorder { get; set; }

        /// <summary>
        /// Display top border line
        /// </summary>
        [DefaultValue(false)]
        public bool TopBorder { get; set; }

        /// <summary>
        /// Display right border line
        /// </summary>
        [DefaultValue(false)]
        public bool RightBorder { get; set; }

        [DefaultValue(false)]
        public bool BottomBorder { get; set; }

        /// <summary>
        /// Border line color
        /// </summary>
        [DefaultValue(typeof (Color), "Black")]
        public Color BorderColor
        {
            get { return _intBorderColor; }
            set { _intBorderColor = value; }
        }

        /// <summary>
        /// Border line color
        /// </summary>
        [DefaultValue(0)]
        public int BorderWidth { get; set; }

        [DefaultValue(DashStyle.Solid)]
        public DashStyle BorderStyle
        {
            get { return _borderStyle; }
            set { _borderStyle = value; }
        }

        [DefaultValue(false)]
        public bool BorderThickness { get; set; }

        [DefaultValue(0)]
        public int BorderSpacing { get; set; }

        /// <summary>
        /// Word wrap
        /// </summary>
        [DefaultValue(false)]
        public bool Multiline { get; set; }

        /// <summary>
        /// Standard tab width
        /// </summary>
        [DefaultValue(100)]
        public int StandTabWidth
        {
            get { return _intStandTabWidth; }
            set { _intStandTabWidth = value; }
        }

        /// <summary>
        /// indent of first line in a paragraph
        /// </summary>
        [DefaultValue(0)]
        public int ParagraphFirstLineIndent { get; set; }

        /// <summary>
        /// Indent of wholly paragraph
        /// </summary>
        [DefaultValue(0)]
        public int LeftIndent { get; set; }

        /// <summary>
        /// character spacing
        /// </summary>
        [DefaultValue(0)]
        public int Spacing { get; set; }

        /// <summary>
        /// line spacing
        /// </summary>
        [DefaultValue(0)]
        public int LineSpacing { get; set; }

        /// <summary>
        /// Current line spacing is multiple extractly line spacing.
        /// </summary>
        [DefaultValue(false)]
        public bool MultipleLineSpacing { get; set; }

        /// <summary>
        /// Spacing before paragrah
        /// </summary>
        [DefaultValue(0)]
        public int SpacingBefore { get; set; }

        /// <summary>
        /// Spacing after paragraph
        /// </summary>
        [DefaultValue(0)]
        public int SpacingAfter { get; set; }

        /// <summary>
        /// text alignment
        /// </summary>
        [DefaultValue(RTFAlignment.Left)]
        public RTFAlignment Align
        {
            get { return _intAlign; }
            set { _intAlign = value; }
        }

        [DefaultValue(false)]
        public bool PageBreak { get; set; }

        public Font Font
        {
            set
            {
                if (value != null)
                {
                    FontName = value.Name;
                    FontSize = value.Size;
                    Bold = value.Bold;
                    Italic = value.Italic;
                    Underline = value.Underline;
                    Strikeout = value.Strikeout;
                }
            }
        }

        /// <summary>
        /// font name
        /// </summary>
        public string FontName
        {
            get { return _strFontName; }
            set { _strFontName = value; }
        }

        /// <summary>
        /// font size
        /// </summary>
        [DefaultValue(12f)]
        public float FontSize
        {
            get { return _fFontSize; }
            set { _fFontSize = value; }
        }

        /// <summary>
        /// bold style
        /// </summary>
        [DefaultValue(false)]
        public bool Bold { get; set; }

        /// <summary>
        /// italic style
        /// </summary>
        [DefaultValue(false)]
        public bool Italic { get; set; }

        /// <summary>
        /// underline style
        /// </summary>
        [DefaultValue(false)]
        public bool Underline { get; set; }

        /// <summary>
        /// strickout style
        /// </summary>
        [DefaultValue(false)]
        public bool Strikeout { get; set; }

        /// <summary>
        /// Hidden text
        /// </summary>
        [DefaultValue(false)]
        public bool Hidden { get; set; }

        /// <summary>
        /// text color
        /// </summary>
        [DefaultValue(typeof (Color), "Black")]
        public Color TextColor
        {
            get { return _intTextColor; }
            set { _intTextColor = value; }
        }

        /// <summary>
        /// back color
        /// </summary>
        [DefaultValue(typeof (Color), "Empty")]
        public Color BackColor
        {
            get { return _intBackColor; }
            set { _intBackColor = value; }
        }

        /// <summary>
        /// link
        /// </summary>
        [DefaultValue(null)]
        public string Link { get; set; }

        /// <summary>
        /// superscript
        /// </summary>
        [DefaultValue(false)]
        public bool Superscript
        {
            get { return _bolSuperscript; }
            set
            {
                _bolSuperscript = value;
                if (_bolSuperscript)
                {
                    _bolSubscript = false;
                }
            }
        }

        /// <summary>
        /// subscript
        /// </summary>
        [DefaultValue(false)]
        public bool Subscript
        {
            get { return _bolSubscript; }
            set
            {
                _bolSubscript = value;
                if (_bolSubscript)
                {
                    _bolSuperscript = false;
                }
            }
        }

        /// <summary>
        /// list overried id 
        /// </summary>
        public int ListId
        {
            get { return _listId; }
            set { _listId = value; }
        }

        /// <summary>
        /// no wrap in word
        /// </summary>
        [DefaultValue(true)]
        public bool NoWwrap
        {
            get { return _bolNoWwrap; }
            set { _bolNoWwrap = value; }
        }

        public void SetAlign(StringAlignment align)
        {
            if (align == StringAlignment.Center)
            {
                Align = RTFAlignment.Center;
            }
            else if (align == StringAlignment.Far)
            {
                Align = RTFAlignment.Right;
            }
            else
            {
                Align = RTFAlignment.Left;
            }
        }

        public bool EqualsSettings(DocumentFormatInfo format)
        {
            if (format == this)
                return true;
            if (format == null)
                return false;
            if (Align != format.Align)
                return false;
            if (BackColor != format.BackColor)
                return false;
            if (Bold != format.Bold)
                return false;
            if (BorderColor != format.BorderColor)
                return false;
            if (LeftBorder != format.LeftBorder)
                return false;
            if (TopBorder != format.TopBorder)
                return false;
            if (RightBorder != format.RightBorder)
                return false;
            if (BottomBorder != format.BottomBorder)
                return false;
            if (BorderStyle != format.BorderStyle)
                return false;
            if (BorderThickness != format.BorderThickness)
                return false;
            if (BorderSpacing != format.BorderSpacing)
                return false;
            if (ListId != format.ListId)
            {
                return false;
            }
            if (FontName != format.FontName)
                return false;
            if (FontSize != format.FontSize)
                return false;
            if (Italic != format.Italic)
                return false;
            if (Hidden != format.Hidden)
                return false;
            if (LeftIndent != format.LeftIndent)
                return false;
            if (LineSpacing != format.LineSpacing)
                return false;
            if (Link != format.Link)
                return false;
            if (Multiline != format.Multiline)
                return false;
            if (NoWwrap != format.NoWwrap)
                return false;
            if (ParagraphFirstLineIndent != format.ParagraphFirstLineIndent)
                return false;
            if (Spacing != format.Spacing)
                return false;
            if (StandTabWidth != format.StandTabWidth)
                return false;
            if (Strikeout != format.Strikeout)
                return false;
            if (Subscript != format.Subscript)
                return false;
            if (Superscript != format.Superscript)
                return false;
            if (TextColor != format.TextColor)
                return false;
            if (Underline != format.Underline)
                return false;
            if (ReadText != format.ReadText)
                return false;
            return true;
        }

        /// <summary>
        /// close instance
        /// </summary>
        /// <returns>new instance</returns>
        public DocumentFormatInfo Clone()
        {
            return (DocumentFormatInfo) MemberwiseClone();

            //DocumentFormatInfo format = new DocumentFormatInfo();
            //format.ParagraphFirstLineIndent = this.ParagraphFirstLineIndent;
            //format.LeftIndent = this.LeftIndent;
            //format.Spacing = this.Spacing;
            //format.LineSpacing = this.LineSpacing;
            //format.Align = this.Align;
            //format.FontName = this.FontName;
            //format.FontSize = this.FontSize;
            //format.Bold = this.Bold;
            //format.Italic = this.Italic;
            //format.Underline = this.Underline;
            //format.Strikeout = this.Strikeout;
            //format.TextColor = this.TextColor;
            //format.BackColor = this.BackColor;
            //format.Hidden = this.Hidden;
            //format.Link = this.Link;
            //format.Superscript = this.Superscript;
            //format.Subscript = this.Subscript;
            //format.BulletedList = this.BulletedList;
            //format.NumberedList = this.NumberedList;
            //format.StandTabWidth = this.StandTabWidth;
            //format.Multiline = this.Multiline;
            //format.NoWwrap = this.NoWwrap;
            //format.myParent = this.myParent;
            //format.LeftBorder = this.LeftBorder;
            //format.TopBorder = this.TopBorder;
            //format.RightBorder = this.RightBorder;
            //format.BottomBorder = this.BottomBorder;
            //format.BorderColor = this.BorderColor;
            //format.BorderStyle = this.BorderStyle;
            //format.BorderThickness = this.BorderThickness;
            //format.BorderSpacing = this.BorderSpacing;
            //format.ReadText = this.ReadText;
            //format.NativeLevel = this.NativeLevel;
            //return format;
        }


        public void ResetText()
        {
            FontName = "Microsoft Sans Serif";
            FontSize = 12;
            Bold = false;
            Italic = false;
            Underline = false;
            Strikeout = false;
            TextColor = Color.Black;
            BackColor = Color.Empty;
            //this.Link = null ;
            Subscript = false;
            Superscript = false;
            Multiline = true;
            Hidden = false;
            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderStyle = DashStyle.Solid;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black;
        }

        public void ResetParagraph()
        {
            ParagraphFirstLineIndent = 0;
            Align = 0;
            ListId = -1;
            LeftIndent = 0;
            LineSpacing = 0;
            PageBreak = false;
            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderStyle = DashStyle.Solid;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black;
            MultipleLineSpacing = false;
            SpacingBefore = 0;
            SpacingAfter = 0;
            //this.LeftBorder = false;
            //this.TopBorder = false;
            //this.RightBorder = false;
            //this.BottomBorder = false;
            //this.BorderColor = Color.Transparent;
        }

        public void Reset()
        {
            ParagraphFirstLineIndent = 0;
            LeftIndent = 0;
            LeftIndent = 0;
            Spacing = 0;
            LineSpacing = 0;
            MultipleLineSpacing = false;
            SpacingBefore = 0;
            SpacingAfter = 0;
            Align = 0;
            FontName = "Microsoft Sans Serif";
            FontSize = 12;
            Bold = false;
            Italic = false;
            Underline = false;
            Strikeout = false;
            TextColor = Color.Black;
            BackColor = Color.Empty;
            Link = null;
            Subscript = false;
            Superscript = false;
            ListId = -1;
            Multiline = true;
            NoWwrap = true;


            LeftBorder = false;
            TopBorder = false;
            RightBorder = false;
            BottomBorder = false;
            BorderSpacing = 0;
            BorderThickness = false;
            BorderColor = Color.Black;

            ReadText = true;
            NativeLevel = 0;
            Hidden = false;
        }
    }
}



//  Rtf Dom Parser

//  Copyright (c) 2010 sinosoft , written by yuans.
//  http://www.sinoreport.net

//  This program is free software; you can redistribute it and/or
//  modify it under the terms of the GNU General Public License
//  as published by the Free Software Foundation; either version 2
//  of the License, or (at your option) any later version.

//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.

//  You should have received a copy of the GNU General Public License
//  along with this program; if not, write to the Free Software
//  Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

//****************************************************************************/

//using System;
//using System.ComponentModel ;

//namespace RtfDomParser
//{
//    /// <summary>
//    /// RTF Document format information
//    /// </summary>
//
//    public class DocumentFormatInfo
//    {
//        /// <summary>
//        /// Initialize instance
//        /// </summary>
//        public DocumentFormatInfo()
//        {
//        }

//        private DocumentFormatInfo myParent = null;
//        /// <summary>
//        /// If this instance is create by Clone , return the parent instance
//        /// </summary>
//
//        public DocumentFormatInfo Parent
//        {
//            get
//            {
//                return myParent ;
//            }
//        }

//        private RTFBorderStyle _Border = new RTFBorderStyle();
//        public RTFBorderStyle Border
//        {
//            get { return _Border; }
//            set { _Border = value; }
//        }

//        private RTFBorderStyle _ParagraphBorder = new RTFBorderStyle();
//        public RTFBorderStyle ParagraphBorder
//        {
//            get { return _ParagraphBorder; }
//            set { _ParagraphBorder = value; }
//        }

//        private int _BorderSpacing = 0;
//        /// <summary>
//        /// space in twips between borders and the paragraph
//        /// </summary>
//        [DefaultValue( 0 )]
//        public int BorderSpacing
//        {
//            get
//            {
//                return _BorderSpacing; 
//            }
//            set
//            {
//                _BorderSpacing = value; 
//            }
//        }

//        private bool bolMultiline = false;
//        /// <summary>
//        /// Word wrap
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Multiline
//        {
//            get
//            {
//                return bolMultiline; 
//            }
//            set
//            {
//                bolMultiline = value; 
//            }
//        }

//        private int intStandTabWidth = 100;
//        /// <summary>
//        /// Standard tab width
//        /// </summary>
//        [DefaultValue(100)]
//        public int StandTabWidth
//        {
//            get
//            {
//                return intStandTabWidth; 
//            }
//            set
//            {
//                intStandTabWidth = value; 
//            }
//        }

//        private int intParagraphFirstLineIndent = 0;
//        /// <summary>
//        /// indent of first line in a paragraph
//        /// </summary>
//        [DefaultValue(0)]
//        public int ParagraphFirstLineIndent
//        {
//            get
//            {
//                return intParagraphFirstLineIndent; 
//            }
//            set
//            {
//                intParagraphFirstLineIndent = value; 
//            }
//        }

//        private int intLeftIndent = 0;
//        /// <summary>
//        /// Indent of wholly paragraph
//        /// </summary>
//        [DefaultValue(0)]
//        public int LeftIndent
//        {
//            get
//            {
//                return intLeftIndent; 
//            }
//            set
//            {
//                intLeftIndent = value; 
//            }
//        }

//        private int intSpacing = 0;
//        /// <summary>
//        /// character spacing
//        /// </summary>
//        [DefaultValue(0)]
//        public int Spacing
//        {
//            get
//            {
//                return intSpacing; 
//            }
//            set
//            {
//                intSpacing = value; 
//            }
//        }

//        private int intLineSpacing = 0;
//        /// <summary>
//        /// line spacing
//        /// </summary>
//        [DefaultValue(0)]
//        public int LineSpacing
//        {
//            get
//            {
//                return intLineSpacing; 
//            }
//            set
//            {
//                intLineSpacing = value; 
//            }
//        }

//        private RTFAlignment intAlign = RTFAlignment.Left;
//        /// <summary>
//        /// text alignment
//        /// </summary>
//        [DefaultValue(RTFAlignment.Left)]
//        public RTFAlignment Align
//        {
//            get
//            {
//                return intAlign; 
//            }
//            set
//            {
//                intAlign = value; 
//            }
//        }

//        /// <summary>
//        /// nest level in native rtf document
//        /// </summary>
//        public int NativeLevel = 0;

//        public void SetAlign( System.Drawing.StringAlignment align )
//        {
//            if (align == System.Drawing.StringAlignment.Center)
//            {
//                this.Align = RTFAlignment.Center;
//            }
//            else if (align == System.Drawing.StringAlignment.Far)
//            {
//                this.Align = RTFAlignment.Right;
//            }
//            else
//            {
//                this.Align = RTFAlignment.Left;
//            }
//        }

//
//        public System.Drawing.Font Font
//        {
//            set
//            {
//                if( value != null )
//                {
//                    FontName = value.Name ;
//                    FontSize = value.Size ;
//                    Bold = value.Bold ;
//                    Italic = value.Italic ;
//                    Underline = value.Underline ;
//                    Strikeout = value.Strikeout ;
//                }
//            }
//        }

//        private string strFontName = System.Windows.Forms.Control.DefaultFont.Name;
//        /// <summary>
//        /// font name
//        /// </summary>
//        public string FontName
//        {
//            get
//            {
//                return strFontName; 
//            }
//            set
//            {
//                strFontName = value; 
//            }
//        }

//        private float fFontSize = 12f;
//        /// <summary>
//        /// font size
//        /// </summary>
//        [DefaultValue(12f)]
//        public float FontSize
//        {
//            get
//            {
//                return fFontSize; 
//            }
//            set
//            {
//                fFontSize = value; 
//            }
//        }


//        private bool bolBold = false;
//        /// <summary>
//        /// bold style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Bold
//        {
//            get
//            {
//                return bolBold; 
//            }
//            set
//            {
//                bolBold = value; 
//            }
//        }

//        private bool bolItalic = false;
//        /// <summary>
//        /// italic style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Italic
//        {
//            get
//            {
//                return bolItalic; 
//            }
//            set
//            {
//                bolItalic = value; 
//            }
//        }

//        private bool bolUnderline = false;
//        /// <summary>
//        /// underline style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Underline
//        {
//            get
//            {
//                return bolUnderline; 
//            }
//            set
//            {
//                bolUnderline = value; 
//            }
//        }

//        private bool bolStrikeout = false;
//        /// <summary>
//        /// strickout style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Strikeout
//        {
//            get
//            {
//                return bolStrikeout; 
//            }
//            set
//            {
//                bolStrikeout = value; 
//            }
//        }

//        private bool _Hidden = false;
//        /// <summary>
//        /// Hidden text
//        /// </summary>
//        [DefaultValue( false )]
//        public bool Hidden
//        {
//            get 
//            {
//                return _Hidden; 
//            }
//            set
//            {
//                _Hidden = value; 
//            }
//        }

//        private Color intTextColor = Color.Black;
//        /// <summary>
//        /// text color
//        /// </summary>
//        [DefaultValue(typeof(Color), "Black")]
//        public Color TextColor
//        {
//            get
//            {
//                return intTextColor; 
//            }
//            set
//            {
//                intTextColor = value; 
//            }
//        }

//        private Color intBackColor = Color.Empty;
//        /// <summary>
//        /// back color
//        /// </summary>
//        [DefaultValue(typeof(Color), "Empty")]
//        public Color BackColor
//        {
//            get
//            {
//                return intBackColor; 
//            }
//            set
//            {
//                intBackColor = value; 
//            }
//        }

//        ///// <summary>
//        ///// ±ß¿òÏßÑÕÉ«
//        ///// </summary>
//        //public Color BorderColor = Color.Empty;
//        private string strLink = null;
//        /// <summary>
//        /// link
//        /// </summary>
//        [DefaultValue(null)]
//        public string Link
//        {
//            get
//            {
//                return strLink; 
//            }
//            set
//            {
//                strLink = value; 
//            }
//        }

//        private bool bolSuperscript = false;
//        /// <summary>
//        /// superscript
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Superscript
//        {
//            get
//            {
//                return bolSuperscript; 
//            }
//            set
//            {
//                bolSuperscript = value;
//                if (bolSuperscript)
//                {
//                    bolSubscript = false;
//                }
//            }
//        }

//        private bool bolSubscript = false;
//        /// <summary>
//        /// subscript
//        /// </summary>
//        [DefaultValue(false)]
//        public bool Subscript
//        {
//            get
//            {
//                return bolSubscript; 
//            }
//            set
//            {
//                bolSubscript = value;
//                if (bolSubscript)
//                {
//                    bolSuperscript = false;
//                }
//            }
//        }

//        private bool bolBulletedList = false;
//        /// <summary>
//        /// list in bulleted style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool BulletedList
//        {
//            get 
//            {
//                return bolBulletedList; 
//            }
//            set
//            {
//                bolBulletedList = value; 
//            }
//        }

//        private bool bolNumberedList = false;
//        /// <summary>
//        /// list in numbered style
//        /// </summary>
//        [DefaultValue(false)]
//        public bool NumberedList
//        {
//            get
//            {
//                return bolNumberedList; 
//            }
//            set
//            {
//                bolNumberedList = value; 
//            }
//        }

//        private bool bolNoWwrap = true;
//        /// <summary>
//        /// no wrap in word
//        /// </summary>
//        [DefaultValue(true)]
//        public bool NoWwrap
//        {
//            get
//            {
//                return bolNoWwrap; 
//            }
//            set
//            {
//                bolNoWwrap = value; 
//            }
//        }

//        internal bool ReadText = true;

//        public bool EqualsSettings(DocumentFormatInfo format )
//        {
//            if (format == this)
//                return true;
//            if (format == null)
//                return false;
//            if (this.Align != format.Align)
//                return false;
//            if (this.BackColor != format.BackColor)
//                return false;
//            if (this.Bold != format.Bold)
//                return false;
//            if (this._Border.EqualsValue(format._Border) == false)
//            {
//                return false;
//            }
//            if (this._ParagraphBorder.EqualsValue(format._ParagraphBorder) == false)
//            {
//                return false;
//            }
//            if (this._BorderSpacing != format._BorderSpacing)
//            {
//                return false;
//            }
//            if (this.BulletedList != format.BulletedList)
//                return false;
//            if (this.FontName != format.FontName)
//                return false;
//            if (this.FontSize != format.FontSize)
//                return false;
//            if (this.Italic != format.Italic)
//                return false;
//            if (this.Hidden != format.Hidden)
//                return false;
//            if (this.LeftIndent != format.LeftIndent)
//                return false;
//            if (this.LineSpacing != format.LineSpacing)
//                return false;
//            if (this.Link != format.Link)
//                return false;
//            if (this.Multiline != format.Multiline)
//                return false;
//            if (this.NoWwrap != format.NoWwrap)
//                return false;
//            if (this.NumberedList != format.NumberedList)
//                return false;
//            if (this.ParagraphFirstLineIndent != format.ParagraphFirstLineIndent)
//                return false;
//            if (this.Spacing != format.Spacing)
//                return false;
//            if (this.StandTabWidth != format.StandTabWidth)
//                return false;
//            if (this.Strikeout != format.Strikeout)
//                return false;
//            if (this.Subscript != format.Subscript)
//                return false;
//            if (this.Superscript != format.Superscript)
//                return false;
//            if (this.TextColor != format.TextColor)
//                return false;
//            if (this.Underline != format.Underline)
//                return false;
//            if (this.ReadText != format.ReadText)
//                return false;
//            return true;
//        }

//        /// <summary>
//        /// close instance
//        /// </summary>
//        /// <returns>new instance</returns>
//        public DocumentFormatInfo Clone()
//        {
//            DocumentFormatInfo format = new DocumentFormatInfo();
//            format.ParagraphFirstLineIndent = this.ParagraphFirstLineIndent;
//            format.LeftIndent = this.LeftIndent;
//            format.Spacing = this.Spacing;
//            format.LineSpacing = this.LineSpacing;
//            format.Align = this.Align;
//            format.FontName = this.FontName;
//            format.FontSize = this.FontSize;
//            format.Bold = this.Bold;
//            format.Italic = this.Italic;
//            format.Underline = this.Underline;
//            format.Strikeout = this.Strikeout;
//            format.TextColor = this.TextColor;
//            format.BackColor = this.BackColor;
//            format.Hidden = this.Hidden;
//            format.Link = this.Link;
//            format.Superscript = this.Superscript;
//            format.Subscript = this.Subscript;
//            format.BulletedList = this.BulletedList;
//            format.NumberedList = this.NumberedList;
//            format.StandTabWidth = this.StandTabWidth;
//            format.Multiline = this.Multiline;
//            format.NoWwrap = this.NoWwrap;
//            format.myParent = this.myParent;
//            format._Border = this._Border.Clone();
//            format._ParagraphBorder = this._ParagraphBorder.Clone();
//            format._BorderSpacing = this._BorderSpacing;
//            format.ReadText = this.ReadText;
//            format.NativeLevel = this.NativeLevel;
//            return format;
//        }


//        public void ResetText()
//        {
//            this.FontName = System.Windows.Forms.Control.DefaultFont.Name;
//            this.FontSize = 12;
//            this.Bold = false;
//            this.Italic = false;
//            this.Underline = false;
//            this.Strikeout = false;
//            this.TextColor = Color.Black;
//            this.BackColor = Color.Empty;
//            //this.Link = null ;
//            this.Subscript = false;
//            this.Superscript = false;
//            this.Multiline = true;
//            this.Hidden = false;
//            this.Border = new RTFBorderStyle();
//        }

//        public void ResetParagraph()
//        {
//            this.ParagraphFirstLineIndent = 0;
//            this.Align = 0;
//            this.BulletedList = false;
//            this.NumberedList = false;
//            this.LeftIndent = 0;
//            this.ParagraphBorder = new RTFBorderStyle();
//            this._BorderSpacing = 0;
//            //this.LeftBorder = false;
//            //this.TopBorder = false;
//            //this.RightBorder = false;
//            //this.BottomBorder = false;
//            //this.BorderColor = Color.Transparent;
//        }

//        public void Reset()
//        {
//            this.ParagraphFirstLineIndent = 0;
//            this.LeftIndent = 0;
//            this.LeftIndent = 0;
//            this.Spacing = 0;
//            this.LineSpacing = 0;
//            this.Align = 0;
//            this.FontName = System.Windows.Forms.Control.DefaultFont.Name;
//            this.FontSize = 12;
//            this.Bold = false;
//            this.Italic = false;
//            this.Underline = false;
//            this.Strikeout = false;
//            this.TextColor = Color.Black;
//            this.BackColor = Color.Empty;
//            this.Link = null;
//            this.Subscript = false;
//            this.Superscript = false;
//            this.BulletedList = false;
//            this.NumberedList = false;
//            this.Multiline = true;
//            this.NoWwrap = true;
//            this.Border = new RTFBorderStyle();
//            this.ParagraphBorder = new RTFBorderStyle();
//            this._BorderSpacing = 0;
//            this.ReadText = true;
//            this.NativeLevel = 0;
//            this.Hidden = false;
//        }

//    }
//}