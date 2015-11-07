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
using System.IO;
using System.Text;

//using DCSoft.Drawing ;
//using DCSoft.Printing ;

namespace RtfDomParser
{
    /// <summary>
    /// RTF document writer
    /// </summary>
    public class RTFDocumentWriter
    {
        private bool _bolCollectionInfo = true;

        private bool _bolFirstParagraph = true;

        private bool _debugMode = true;

        private RTFListOverrideTable _listOverrideTable = new RTFListOverrideTable();

        private RTFListTable _listTable = new RTFListTable();

        /// <summary>
        /// rtf color table
        /// </summary>
        private readonly RTFColorTable _myColorTable = new RTFColorTable();

        /// <summary>
        /// rtf font table
        /// </summary>
        private readonly RTFFontTable _myFontTable = new RTFFontTable();

        /// <summary>
        /// document information
        /// </summary>
        private readonly Dictionary<string, object> _myInfo = new Dictionary<string, object>();

        private DocumentFormatInfo _myLastParagraphInfo;

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDocumentWriter()
        {
            Writer = null;
            _myColorTable.CheckValueExistWhenAdd = true;
        }

        public RTFDocumentWriter(TextWriter writer)
        {
            Writer = null;
            _myColorTable.CheckValueExistWhenAdd = true;
            Open(writer);
        }

        public RTFDocumentWriter(Stream stream)
        {
            Writer = null;
            _myColorTable.CheckValueExistWhenAdd = true;
            var writer = new StreamWriter(
                stream,
                Encoding.GetEncoding("us-ascii"));
            Open(writer);
        }

        /// <summary>
        /// base writer
        /// </summary>
        public RTFWriter Writer { get; set; }

        public Dictionary<string, object> Info
        {
            get { return _myInfo; }
        }

        /// <summary>
        /// rtf font table
        /// </summary>
        public RTFFontTable FontTable
        {
            get { return _myFontTable; }
        }

        public RTFListTable ListTable
        {
            get { return _listTable; }
            set { _listTable = value; }
        }

        public RTFListOverrideTable ListOverrideTable
        {
            get { return _listOverrideTable; }
            set { _listOverrideTable = value; }
        }

        /// <summary>
        /// rtf color table
        /// </summary>
        public RTFColorTable ColorTable
        {
            get { return _myColorTable; }
        }

        /// <summary>
        /// system collectiong document's information , maby generating
        /// font table and color table , not writting content.
        /// </summary>
        public bool CollectionInfo
        {
            get { return _bolCollectionInfo; }
            set { _bolCollectionInfo = value; }
        }

        public int GroupLevel
        {
            get { return Writer.GroupLevel; }
        }

        public bool DebugMode
        {
            get { return _debugMode; }
            set { _debugMode = value; }
        }

        public virtual bool Open(TextWriter writer)
        {
            Writer = new RTFWriter(writer);
            Writer.Encoding = Encoding.UTF8;
            Writer.Indent = false;
            return true;
        }

        public virtual bool Open(Stream strFileName)
        {
            Writer = new RTFWriter(strFileName);
            Writer.Encoding = Encoding.UTF8;
            Writer.Indent = false;
            return true;
        }

        public virtual void Close()
        {
            Writer.Close();
        }

        public void WriteStartGroup()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteStartGroup();
            }
        }

        public void WriteEndGroup()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// write rtf keyword
        /// </summary>
        /// <param name="keyword">keyword</param>
        public void WriteKeyword(string keyword)
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteKeyword(keyword);
            }
        }

        public void WriteKeyword(string keyWord, bool ext)
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteKeyword(keyWord, ext);
            }
        }

        public void WriteRaw(string txt)
        {
            if (_bolCollectionInfo == false)
            {
                if (txt != null)
                {
                    Writer.WriteRaw(txt);
                }
            }
        }

        public void WriteBorderLineDashStyle(DashStyle style)
        {
            if (_bolCollectionInfo == false)
            {
                if (style == DashStyle.Dot)
                {
                    WriteKeyword("brdrdot");
                }
                else if (style == DashStyle.DashDot)
                {
                    WriteKeyword("brdrdashd");
                }
                else if (style == DashStyle.DashDotDot)
                {
                    WriteKeyword("brdrdashdd");
                }
                else if (style == DashStyle.Dash)
                {
                    WriteKeyword("brdrdash");
                }
                else
                {
                    WriteKeyword("brdrs");
                }
            }
        }

        /// <summary>
        /// start write document
        /// </summary>
        public void WriteStartDocument()
        {
            _myLastParagraphInfo = null;
            _bolFirstParagraph = true;
            if (_bolCollectionInfo)
            {
                _myInfo.Clear();
                _myFontTable.Clear();
                _myColorTable.Clear();
                _myFontTable.Add("Microsoft Sans Serif");
            }
            else
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword(RTFConsts.RTF);
                Writer.WriteKeyword("ansi");
                Writer.WriteKeyword("ansicpg" + Writer.CodePageNumber);
                // write document information
                if (_myInfo.Count > 0)
                {
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("info");
                    foreach (var strKey in _myInfo.Keys)
                    {
                        Writer.WriteStartGroup();

                        var v = _myInfo[strKey];
                        if (v is string)
                        {
                            Writer.WriteKeyword(strKey);
                            Writer.WriteText((string) v);
                        }
                        else if (v is int)
                        {
                            Writer.WriteKeyword(strKey + v);
                        }
                        else if (v is DateTime)
                        {
                            var dtm = (DateTime) v;
                            Writer.WriteKeyword(strKey);
                            Writer.WriteKeyword("yr" + dtm.Year);
                            Writer.WriteKeyword("mo" + dtm.Month);
                            Writer.WriteKeyword("dy" + dtm.Day);
                            Writer.WriteKeyword("hr" + dtm.Hour);
                            Writer.WriteKeyword("min" + dtm.Minute);
                            Writer.WriteKeyword("sec" + dtm.Second);
                        }
                        else
                        {
                            Writer.WriteKeyword(strKey);
                        }

                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }
                // writing font table
                Writer.WriteStartGroup();
                Writer.WriteKeyword(RTFConsts.Fonttbl);
                for (var iCount = 0; iCount < _myFontTable.Count; iCount ++)
                {
                    //string f = myFontTable[ iCount ] ;
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("f" + iCount);
                    var f = _myFontTable[iCount];
                    Writer.WriteText(f.Name);
                    if (f.Charset != 1)
                    {
                        Writer.WriteKeyword("fcharset" + f.Charset);
                    }
                    Writer.WriteEndGroup();
                }
                Writer.WriteEndGroup();

                // write color table
                Writer.WriteStartGroup();
                Writer.WriteKeyword(RTFConsts.Colortbl);
                Writer.WriteRaw(";");
                for (var iCount = 0; iCount < _myColorTable.Count; iCount ++)
                {
                    var c = _myColorTable[iCount];
                    Writer.WriteKeyword("red" + c.R);
                    Writer.WriteKeyword("green" + c.G);
                    Writer.WriteKeyword("blue" + c.B);
                    Writer.WriteRaw(";");
                }
                Writer.WriteEndGroup();

                // write list table
                if (ListTable != null && ListTable.Count > 0)
                {
                    if (DebugMode)
                    {
                        Writer.WriteRaw(Environment.NewLine);
                    }
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("listtable", true);
                    foreach (var list in ListTable)
                    {
                        if (DebugMode)
                        {
                            Writer.WriteRaw(Environment.NewLine);
                        }
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("list");
                        Writer.WriteKeyword("listtemplateid" + list.ListTemplateId);
                        if (list.ListHybrid)
                        {
                            Writer.WriteKeyword("listhybrid");
                        }
                        if (DebugMode)
                        {
                            Writer.WriteRaw(Environment.NewLine);
                        }
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("listlevel");
                        Writer.WriteKeyword("levelfollow" + list.LevelFollow);
                        Writer.WriteKeyword("leveljc" + list.LevelJc);
                        Writer.WriteKeyword("levelstartat" + list.LevelStartAt);
                        Writer.WriteKeyword("levelnfc" + Convert.ToInt32(list.LevelNfc));
                        Writer.WriteKeyword("levelnfcn" + Convert.ToInt32(list.LevelNfc));
                        Writer.WriteKeyword("leveljc" + list.LevelJc);

                        //if (list.LevelNfc == LevelNumberType.Bullet)
                        {
                            if (string.IsNullOrEmpty(list.LevelText) == false)
                            {
                                Writer.WriteStartGroup();
                                Writer.WriteKeyword("leveltext");
                                Writer.WriteKeyword("'0" + list.LevelText.Length);
                                if (list.LevelNfc == LevelNumberType.Bullet)
                                {
                                    Writer.WriteUnicodeText(list.LevelText);
                                }
                                else
                                {
                                    Writer.WriteText(list.LevelText, false);
                                }
                                //myWriter.WriteStartGroup();
                                //myWriter.WriteKeyword("uc1");
                                //int v = (int)list.LevelText[0];
                                //short uv = (short)v;
                                //myWriter.WriteKeyword("u" + uv);
                                //myWriter.WriteRaw(" ?");
                                //myWriter.WriteEndGroup();
                                //myWriter.WriteRaw(";");
                                Writer.WriteEndGroup();
                                if (list.LevelNfc == LevelNumberType.Bullet)
                                {
                                    var f = FontTable["Wingdings"];
                                    if (f != null)
                                    {
                                        Writer.WriteKeyword("f" + f.Index);
                                    }
                                }
                                else
                                {
                                    Writer.WriteStartGroup();
                                    Writer.WriteKeyword("levelnumbers");
                                    Writer.WriteKeyword("'01");
                                    Writer.WriteEndGroup();
                                }
                            }
                        }
                        Writer.WriteEndGroup();

                        Writer.WriteKeyword("listid" + list.ListId);
                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }

                // write list overried table
                if (ListOverrideTable != null && ListOverrideTable.Count > 0)
                {
                    if (DebugMode)
                    {
                        Writer.WriteRaw(Environment.NewLine);
                    }
                    Writer.WriteStartGroup();
                    Writer.WriteKeyword("listoverridetable");
                    foreach (var lo in ListOverrideTable)
                    {
                        if (DebugMode)
                        {
                            Writer.WriteRaw(Environment.NewLine);
                        }
                        Writer.WriteStartGroup();
                        Writer.WriteKeyword("listoverride");
                        Writer.WriteKeyword("listid" + lo.ListId);
                        Writer.WriteKeyword("listoverridecount" + lo.ListOverriedCount);
                        Writer.WriteKeyword("ls" + lo.Id);
                        Writer.WriteEndGroup();
                    }
                    Writer.WriteEndGroup();
                }

                if (DebugMode)
                {
                    Writer.WriteRaw(Environment.NewLine);
                }
                Writer.WriteKeyword("viewkind1");
            }
        }

        /// <summary>
        /// end write document
        /// </summary>
        public void WriteEndDocument()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteEndGroup();
            }
            Writer.Flush();
        }

        /// <summary>
        /// start write header
        /// </summary>
        public void WriteStartHeader()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("header");
            }
        }

        /// <summary>
        /// end write header
        /// </summary>
        public void WriteEndHeader()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// start write footer
        /// </summary>
        public void WriteStartFooter()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("footer");
            }
        }

        /// <summary>
        /// end write end footer
        /// </summary>
        public void WriteEndFooter()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteEndGroup();
            }
        }

        public void WriteStartParagraph()
        {
            WriteStartParagraph(new DocumentFormatInfo());
        }

        /// <summary>
        /// write write paragraph
        /// </summary>
        /// <param name="info">format</param>
        public void WriteStartParagraph(DocumentFormatInfo info)
        {
            if (_bolCollectionInfo)
            {
                //myFontTable.Add("Wingdings");
            }
            else
            {
                if (_bolFirstParagraph)
                {
                    _bolFirstParagraph = false;
                    Writer.WriteRaw(Environment.NewLine);
                    //myWriter.WriteKeyword("par");
                }
                else
                {
                    Writer.WriteKeyword("par");
                }
                if (info.ListId >= 0)
                {
                    Writer.WriteKeyword("pard");
                    Writer.WriteKeyword("ls" + info.ListId);
                }
                //if( lo != null && listInfo != null )
                //{
                //    myWriter.WriteKeyword("pard");
                //    myWriter.WriteKeyword("ls" , lo.ListID );
                //    if( listInfo.LevelNfc info.NumberedList )
                //    {
                //        if( myLastParagraphInfo == null 
                //            || myLastParagraphInfo.NumberedList != info.NumberedList )
                //        {
                //            myWriter.WriteKeyword("pard");
                //            myWriter.WriteStartGroup();
                //            myWriter.WriteKeyword("pn" , true );
                //            myWriter.WriteKeyword("pnlvlbody");
                //            myWriter.WriteKeyword("pnindent400");
                //            myWriter.WriteKeyword("pnstart1");
                //            myWriter.WriteKeyword("pndec");
                //            myWriter.WriteEndGroup();
                //        }
                //    }
                //    else
                //    {
                //        if( myLastParagraphInfo == null
                //            || myLastParagraphInfo.BulletedList != info.BulletedList )
                //        {
                //            myWriter.WriteKeyword("pard");
                //            myWriter.WriteStartGroup();
                //            myWriter.WriteKeyword("pn" , true );
                //            myWriter.WriteKeyword("pnlvlblt");
                //            myWriter.WriteKeyword("pnindent400");
                //            myWriter.WriteKeyword("pnf" + myFontTable.IndexOf( "Wingdings" ));
                //            myWriter.WriteStartGroup();
                //            myWriter.WriteKeyword("pntxtb");
                //            myWriter.WriteText("l");
                //            //myWriter.WriteKeyword("'B7");
                //            myWriter.WriteEndGroup();
                //            myWriter.WriteEndGroup();
                //        }
                //    }
                //    myWriter.WriteKeyword("fi-400");
                //}
                //else
                {
                    if (_myLastParagraphInfo != null)
                    {
                        if (_myLastParagraphInfo.ListId >= 0)
                        {
                            Writer.WriteKeyword("pard");
                        }
                    }
                }

                switch (info.Align)
                {
                    case RTFAlignment.Left:
                        Writer.WriteKeyword("ql");
                        break;
                    case RTFAlignment.Center:
                        Writer.WriteKeyword("qc");
                        break;
                    case RTFAlignment.Right:
                        Writer.WriteKeyword("qr");
                        break;
                    case RTFAlignment.Justify:
                        Writer.WriteKeyword("qj");
                        break;
                }
                //
                //				if( info.LeftAlign )
                //					myWriter.WriteKeyword("ql");
                //				if( info.CenterAlign )
                //					myWriter.WriteKeyword("qc");
                //				else if( info.RigthAlign )
                //					myWriter.WriteKeyword("qr");

                //if( info.NumberedList == false && info.BulletedList == false )
                {
                    if (info.ParagraphFirstLineIndent != 0)
                    {
                        Writer.WriteKeyword("fi" + Convert.ToInt32(
                            info.ParagraphFirstLineIndent*400/info.StandTabWidth));
                    }
                    else
                    {
                        Writer.WriteKeyword("fi0");
                    }
                }
                //if( info.NumberedList == false && info.BulletedList == false )
                {
                    if (info.LeftIndent != 0)
                    {
                        Writer.WriteKeyword("li" + Convert.ToInt32(
                            info.LeftIndent*400/info.StandTabWidth));
                    }
                    else
                    {
                        Writer.WriteKeyword("li0");
                    }
                }
                Writer.WriteKeyword("plain");
            }
            _myLastParagraphInfo = info;
        }

        /// <summary>
        /// end write paragraph
        /// </summary>
        public void WriteEndParagraph()
        {
        }

        /// <summary>
        /// write plain text
        /// </summary>
        /// <param name="strText">text</param>
        public void WriteText(string strText)
        {
            if (strText != null && _bolCollectionInfo == false)
            {
                Writer.WriteText(strText);
            }
        }

        /// <summary>
        /// write font format
        /// </summary>
        /// <param name="font">font</param>
        public void WriteFont(Font font)
        {
            if (font == null)
            {
                throw new ArgumentNullException("font");
            }
            if (_bolCollectionInfo)
            {
                _myFontTable.Add(font.Name);
            }
            else
            {
                var index = _myFontTable.IndexOf(font.Name);
                if (index >= 0)
                {
                    Writer.WriteKeyword("f" + index);
                }
                if (font.Bold)
                {
                    Writer.WriteKeyword("b");
                }
                if (font.Italic)
                {
                    Writer.WriteKeyword("i");
                }
                if (font.Underline)
                {
                    Writer.WriteKeyword("ul");
                }
                if (font.Strikeout)
                {
                    Writer.WriteKeyword("strike");
                }
                Writer.WriteKeyword("fs" + Convert.ToInt32(font.Size*2));
            }
        }

        /// <summary>
        /// start write formatted text
        /// </summary>
        /// <param name="info">format</param>
        /// <remarks>
        /// This function must assort with WriteEndString strict
        /// </remarks>
        public void WriteStartString(DocumentFormatInfo info)
        {
            if (_bolCollectionInfo)
            {
                _myFontTable.Add(info.FontName);
                _myColorTable.Add(info.TextColor);
                _myColorTable.Add(info.BackColor);
                if (info.BorderColor.A != 0)
                {
                    _myColorTable.Add(info.BorderColor);
                }
                return;
            }
            if (info.Link != null && info.Link.Length > 0)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("field");
                Writer.WriteStartGroup();
                Writer.WriteKeyword("fldinst", true);
                Writer.WriteStartGroup();
                Writer.WriteKeyword("hich");
                Writer.WriteText(" HYPERLINK \"" + info.Link + "\"");
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
                Writer.WriteStartGroup();
                Writer.WriteKeyword("fldrslt");
                Writer.WriteStartGroup();
            }

            switch (info.Align)
            {
                case RTFAlignment.Left:
                    Writer.WriteKeyword("ql");
                    break;
                case RTFAlignment.Center:
                    Writer.WriteKeyword("qc");
                    break;
                case RTFAlignment.Right:
                    Writer.WriteKeyword("qr");
                    break;
                case RTFAlignment.Justify:
                    Writer.WriteKeyword("qj");
                    break;
            }

            Writer.WriteKeyword("plain");
            var index = 0;
            index = _myFontTable.IndexOf(info.FontName);
            if (index >= 0)
                Writer.WriteKeyword("f" + index);
            if (info.Bold)
                Writer.WriteKeyword("b");
            if (info.Italic)
                Writer.WriteKeyword("i");
            if (info.Underline)
                Writer.WriteKeyword("ul");
            if (info.Strikeout)
                Writer.WriteKeyword("strike");
            Writer.WriteKeyword("fs" + Convert.ToInt32(info.FontSize*2));

            // back color
            index = _myColorTable.IndexOf(info.BackColor);
            if (index >= 0)
            {
                Writer.WriteKeyword("chcbpat" + Convert.ToString(index + 1));
            }

            index = _myColorTable.IndexOf(info.TextColor);
            if (index >= 0)
            {
                Writer.WriteKeyword("cf" + Convert.ToString(index + 1));
            }
            if (info.Subscript)
            {
                Writer.WriteKeyword("sub");
            }
            if (info.Superscript)
                Writer.WriteKeyword("super");
            if (info.NoWwrap)
                Writer.WriteKeyword("nowwrap");
            if (info.LeftBorder
                || info.TopBorder
                || info.RightBorder
                || info.BottomBorder)
            {
                // border color
                if (info.BorderColor.A != 0)
                {
                    Writer.WriteKeyword("chbrdr");
                    Writer.WriteKeyword("brdrs");
                    Writer.WriteKeyword("brdrw10");
                    index = _myColorTable.IndexOf(info.BorderColor);
                    if (index >= 0)
                    {
                        Writer.WriteKeyword("brdrcf" + Convert.ToString(index + 1));
                    }
                }
            }
        }

        public void WriteEndString(DocumentFormatInfo info)
        {
            if (_bolCollectionInfo)
            {
                return;
            }

            if (info.Subscript)
                Writer.WriteKeyword("sub0");
            if (info.Superscript)
                Writer.WriteKeyword("super0");

            if (info.Bold)
                Writer.WriteKeyword("b0");
            if (info.Italic)
                Writer.WriteKeyword("i0");
            if (info.Underline)
                Writer.WriteKeyword("ul0");
            if (info.Strikeout)
                Writer.WriteKeyword("strike0");
            if (info.Link != null && info.Link.Length > 0)
            {
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// write formatted string
        /// </summary>
        /// <param name="strText">text</param>
        /// <param name="info">format</param>
        public void WriteString(string strText, DocumentFormatInfo info)
        {
            if (_bolCollectionInfo)
            {
                _myFontTable.Add(info.FontName);
                _myColorTable.Add(info.TextColor);
                _myColorTable.Add(info.BackColor);
            }
            else
            {
                WriteStartString(info);

                if (info.Multiline)
                {
                    if (strText != null)
                    {
                        strText = strText.Replace("\n", "");
                        using (var reader = new StringReader(strText))
                        {
                            var strLine = reader.ReadLine();
                            var iCount = 0;
                            while (strLine != null)
                            {
                                if (iCount > 0)
                                {
                                    Writer.WriteKeyword("line");
                                }

                                iCount++;
                                Writer.WriteText(strLine);
                                strLine = reader.ReadLine();
                            }
                        }
                    }
                }
                else
                {
                    Writer.WriteText(strText);
                }

                WriteEndString(info);
            }
        }

        /// <summary>
        /// end write string
        /// </summary>
        public void WriteEndString()
        {
        }

        /// <summary>
        /// start write bookmark
        /// </summary>
        /// <param name="strName">bookmark name</param>
        public void WriteStartBookmark(string strName)
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteStartGroup();
                Writer.WriteKeyword("bkmkstart", true);
                Writer.WriteKeyword("f0");
                Writer.WriteText(strName);
                Writer.WriteEndGroup();

                Writer.WriteStartGroup();
                Writer.WriteKeyword("bkmkend", true);
                Writer.WriteKeyword("f0");
                Writer.WriteText(strName);
                Writer.WriteEndGroup();
            }
        }

        /// <summary>
        /// end write bookmark
        /// </summary>
        /// <param name="strName">bookmark name</param>
        public void WriteEndBookmark(string strName)
        {
        }

        /// <summary>
        /// write a line break
        /// </summary>
        public void WriteLineBreak()
        {
            if (_bolCollectionInfo == false)
            {
                Writer.WriteKeyword("line");
            }
        }

        /*
		/// <summary>
		/// write image
		/// </summary>
		/// <param name="img">image</param>
		/// <param name="width">pixel width</param>
		/// <param name="height">pixel height</param>
		/// <param name="ImageData">image binary data</param>
		public void WriteImage( System.Drawing.Image img , int width , int height , byte[] ImageData )
		{
			if( this.bolCollectionInfo )
			{
				return ;
			}
			else
			{
				if( ImageData == null )
					return ;

				System.IO.MemoryStream ms = new System.IO.MemoryStream();
				img.Save( ms , System.Drawing.Imaging.ImageFormat.Jpeg );
				ms.Close();
				byte[] bs = ms.ToArray();
				myWriter.WriteStartGroup();
				
				myWriter.WriteKeyword("pict");
				myWriter.WriteKeyword("jpegblip");
				myWriter.WriteKeyword("picscalex" + Convert.ToInt32( width * 100.0 / img.Size.Width ));
				myWriter.WriteKeyword("picscaley" + Convert.ToInt32( height * 100.0 / img.Size.Height ));
				myWriter.WriteKeyword("picwgoal" + Convert.ToString( img.Size.Width * 15 ));
				myWriter.WriteKeyword("pichgoal" + Convert.ToString( img.Size.Height * 15 ));
				myWriter.WriteBytes( bs );
				myWriter.WriteEndGroup();
			}
		}
         */
    }
}