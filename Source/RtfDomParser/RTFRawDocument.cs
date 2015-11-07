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

namespace RtfDomParser
{
    /// <summary>
    /// RTF raw document,this source code evolution from other software.
    /// </summary>
    public class RTFRawDocument : RTFNodeGroup
    {
        /// <summary>
        /// text encoding for current associate font
        /// </summary>
        private Encoding _myAssociateFontChartset;

        private Encoding _myEncoding;

        /// <summary>
        /// text encoding for current font
        /// </summary>
        private Encoding _myFontChartset;

        /// <summary>
        /// color table
        /// </summary>
        protected RTFColorTable MyColorTable = new RTFColorTable();

        /// <summary>
        /// font table
        /// </summary>
        protected RTFFontTable MyFontTable = new RTFFontTable();

        /// <summary>
        /// document information
        /// </summary>
        protected RTFDocumentInfo MyInfo = new RTFDocumentInfo();


        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFRawDocument()
        {
            MyOwnerDocument = this;
            MyParent = null;
            MyColorTable.CheckValueExistWhenAdd = false;
        }

        /// <summary>
        /// this owner document is myself
        /// </summary>
        public override RTFRawDocument OwnerDocument
        {
            get { return this; }
            set { }
        }

        /// <summary>
        /// no parent node
        /// </summary>
        public override RTFNodeGroup Parent
        {
            get { return null; }
            set { }
        }

        /// <summary>
        /// color table
        /// </summary>
        public RTFColorTable ColorTable
        {
            get { return MyColorTable; }
        }

        /// <summary>
        /// font table
        /// </summary>
        public RTFFontTable FontTable
        {
            get { return MyFontTable; }
        }

        /// <summary>
        /// document information
        /// </summary>
        public RTFDocumentInfo Info
        {
            get { return MyInfo; }
        }

        /// <summary>
        /// text encoding
        /// </summary>
        public Encoding Encoding
        {
            get
            {
                if (_myEncoding == null)
                {
                    var node = MyNodes[RTFConsts.Ansicpg];
                    if (node != null && node.HasParameter)
                    {
                        //todo look up code in encoding dictionary
                        //myEncoding = System.Text.Encoding.GetEncoding( node.Parameter );
                    }
                }
                if (_myEncoding == null)
                    _myEncoding = Encoding.UTF8;
                return _myEncoding;
            }
        }

        /// <summary>
        /// current text encoding
        /// </summary>
        internal Encoding RuntimeEncoding
        {
            get
            {
                if (_myFontChartset != null)
                {
                    return _myFontChartset;
                }
                if (_myAssociateFontChartset != null)
                {
                    return _myAssociateFontChartset;
                }
                return Encoding;
            }
        }

        /// <summary>
        /// read font table
        /// </summary>
        /// <param name="group"></param>
        private void ReadFontTable(RTFNodeGroup group)
        {
            MyFontTable.Clear();
            foreach (var node in group.Nodes)
            {
                if (node is RTFNodeGroup)
                {
                    var index = -1;
                    string name = null;
                    var charset = 0;
                    foreach (var item in node.Nodes)
                    {
                        if (item.Keyword == "f" && item.HasParameter)
                        {
                            index = item.Parameter;
                        }
                        else if (item.Keyword == RTFConsts.Fcharset)
                        {
                            charset = item.Parameter;
                        }
                        else if (item.Type == RTFNodeType.Text)
                        {
                            if (item.Keyword != null && item.Keyword.Length > 0)
                            {
                                name = item.Keyword;
                                break;
                            }
                        }
                    }
                    if (index >= 0 && name != null)
                    {
                        if (name.EndsWith(";"))
                            name = name.Substring(0, name.Length - 1);
                        name = name.Trim();
                        //System.Console.WriteLine( "Index:" + index + "  Name:" + name );
                        var font = new RTFFont(index, name);
                        font.Charset = charset;
                        MyFontTable.Add(font);
                    }
                }
            }
        }

        /// <summary>
        /// read color table
        /// </summary>
        /// <param name="group"></param>
        private void ReadColorTable(RTFNodeGroup group)
        {
            MyColorTable.Clear();
            var r = -1;
            var g = -1;
            var b = -1;
            foreach (var node in group.Nodes)
            {
                if (node.Keyword == "red")
                {
                    r = node.Parameter;
                }
                else if (node.Keyword == "green")
                {
                    g = node.Parameter;
                }
                else if (node.Keyword == "blue")
                {
                    b = node.Parameter;
                }
                if (node.Keyword == ";")
                {
                    if (r >= 0 && g >= 0 && b >= 0)
                    {
                        var c = Color.FromArgb(255, r, g, b);
                        MyColorTable.Add(c);
                        r = -1;
                        g = -1;
                        b = -1;
                    }
                }
            }
            if (r >= 0 && g >= 0 && b >= 0)
            {
                // read the last color
                var c = Color.FromArgb(255, r, g, b);
                MyColorTable.Add(c);
            }
        }

        /// <summary>
        /// read document information
        /// </summary>
        /// <param name="group"></param>
        private void ReadDocumentInfo(RTFNodeGroup group)
        {
            MyInfo.Clear();
            var list = group.GetAllNodes(false);
            foreach (var node in group.Nodes)
            {
                if (node is RTFNodeGroup == false)
                {
                    continue;
                }
                if (node.Keyword == "creatim")
                {
                    MyInfo.Creatim = ReadDateTime(node);
                }
                else if (node.Keyword == "revtim")
                {
                    MyInfo.Revtim = ReadDateTime(node);
                }
                else if (node.Keyword == "printim")
                {
                    MyInfo.Printim = ReadDateTime(node);
                }
                else if (node.Keyword == "buptim")
                {
                    MyInfo.Buptim = ReadDateTime(node);
                }
                else
                {
                    if (node.HasParameter)
                        MyInfo.SetInfo(node.Keyword, node.Parameter.ToString());
                    else
                    {
                        MyInfo.SetInfo(node.Keyword, node.Nodes.Text);
                    }
                }
            }
        }

        private DateTime ReadDateTime(RTFNode g)
        {
            var yr = g.Nodes.GetParameter("yr", 1900);
            var mo = g.Nodes.GetParameter("mo", 1);
            var dy = g.Nodes.GetParameter("dy", 1);
            var hr = g.Nodes.GetParameter("hr", 0);
            var min = g.Nodes.GetParameter("min", 0);
            var sec = g.Nodes.GetParameter("sec", 0);
            return new DateTime(yr, mo, dy, hr, min, sec);
        }

        /// <summary>
        /// load rtf text
        /// </summary>
        /// <param name="strText">text in rtf format</param>
        public void LoadRTFText(string strText)
        {
            _myEncoding = null;
            using (var reader = new RTFReader())
            {
                if (reader.LoadRTFText(strText))
                {
                    Load(reader);
                    reader.Close();
                }
                reader.Close();
            }
        }

        /// <summary>
        /// load rtf file
        /// </summary>
        public void Load(Stream stream)
        {
            _myEncoding = null;
            using (var reader = new RTFReader())
            {
                reader.LoadStream(stream);

                Load(reader);
                reader.Close();
            }
        }

        public void Load(TextReader reader)
        {
            var myReader = new RTFReader();
            myReader.LoadReader(reader);
            Load(myReader);
        }

        /// <summary>
        /// load rtf
        /// </summary>
        /// <param name="reader">RTF text reader</param>
        public void Load(RTFReader reader)
        {
            MyNodes.Clear();
            var groups = new Stack<RTFNodeGroup>();
            RTFNodeGroup newGroup = null;
            RTFNode newNode = null;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    // begin group
                    if (newGroup == null)
                    {
                        newGroup = this;
                    }
                    else
                    {
                        newGroup = new RTFNodeGroup();
                        newGroup.OwnerDocument = this;
                    }
                    if (newGroup != this)
                    {
                        var g = groups.Peek();
                        g.AppendChild(newGroup);
                    }
                    groups.Push(newGroup);
                }
                else if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    // end group
                    newGroup = groups.Pop();
                    newGroup.MergeText();
                    if (newGroup.FirstNode is RTFNode)
                    {
                        switch (newGroup.Keyword)
                        {
                            case RTFConsts.Fonttbl:
                                // read font table
                                ReadFontTable(newGroup);
                                break;
                            case RTFConsts.Colortbl:
                                // read color table
                                ReadColorTable(newGroup);
                                break;
                            case RTFConsts.Info:
                                // read document information
                                ReadDocumentInfo(newGroup);
                                break;
                        }
                    }
                    if (groups.Count > 0)
                    {
                        newGroup = groups.Peek();
                    }
                    else
                    {
                        break;
                    }
                    //NewGroup.MergeText();
                }
                else
                {
                    // read content

                    newNode = new RTFNode(reader.CurrentToken);
                    newNode.OwnerDocument = this;
                    newGroup.AppendChild(newNode);
                    if (newNode.Keyword == RTFConsts.F)
                    {
                        var font = FontTable[newNode.Parameter];
                        if (font != null)
                        {
                            _myFontChartset = font.Encoding;
                        }
                        else
                        {
                            _myFontChartset = null;
                        }
                        //myFontChartset = RTFFont.GetRTFEncoding( NewNode.Parameter );
                    }
                    else if (newNode.Keyword == RTFConsts.Af)
                    {
                        var font = FontTable[newNode.Parameter];
                        if (font != null)
                        {
                            _myAssociateFontChartset = font.Encoding;
                        }
                        else
                        {
                            _myAssociateFontChartset = null;
                        }
                    }
                }
            } // while( reader.ReadToken() != null )
            while (groups.Count > 0)
            {
                newGroup = groups.Pop();
                newGroup.MergeText();
            }
            //this.UpdateInformation();
        }

        /// <summary>
        /// write rtf
        /// </summary>
        /// <param name="writer">RTF writer</param>
        public override void Write(RTFWriter writer)
        {
            writer.Encoding = Encoding;
            base.Write(writer);
        }


        /// <summary>
        /// Save rtf to a stream
        /// </summary>
        /// <param name="stream">stream</param>
        public void Save(Stream stream)
        {
            using (var writer = new RTFWriter(new StreamWriter(stream, Encoding)))
            {
                Write(writer);
                writer.Close();
            }
        }
    }
}