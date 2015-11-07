/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;

namespace RtfDomParser
{
    /// <summary>
    /// RTF Document
    /// </summary>
    /// <remarks>
    /// This type is the root of RTF Dom tree structure
    /// </remarks>
    public class RTFDomDocument : RTFDomElement
    {
        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomDocument()
        {
            LeftMargin = 1800;
            PaperHeight = 15840;
            PaperWidth = 12240;
            Generator = null;
            ListOverrideTable = new RTFListOverrideTable();
            ListTable = new RTFListTable();
            ColorTable = new RTFColorTable();
            FontTable = new RTFFontTable();
            LeadingChars = null;
            FollowingChars = null;
            OwnerDocument = this;
        }

        /// <summary>
        /// following characters
        /// </summary>
        [DefaultValue(null)]
        public string FollowingChars { get; set; }

        /// <summary>
        /// leading characters
        /// </summary>
        [DefaultValue(null)]
        public string LeadingChars { get; set; }

        private Encoding _myDefaultEncoding = Encoding.UTF8;
        /// <summary>
        /// text encoding of current font
        /// </summary>
        private Encoding _myFontChartset;
        /// <summary>
        /// text encoding of associate font 
        /// </summary>
        private Encoding _myAssociateFontChartset;
        /// <summary>
        /// text encoding
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
                return _myDefaultEncoding;
            }
        }

        /// <summary>
        /// default font name
        /// </summary>
        private static string _defaultFontName = "Microsoft Sans Serif";


        /// <summary>
        /// font table
        /// </summary>
        public RTFFontTable FontTable { get; set; }

        /// <summary>
        /// color table
        /// </summary>
        public RTFColorTable ColorTable { get; set; }

        public RTFListTable ListTable { get; set; }

        public RTFListOverrideTable ListOverrideTable { get; set; }

        private RTFDocumentInfo _myInfo = new RTFDocumentInfo();
        /// <summary>
        /// document information
        /// </summary>
        public RTFDocumentInfo Info
        {
            get
            {
                return _myInfo;
            }
            set
            {
                _myInfo = value;
            }
        }

        /// <summary>
        /// document generator
        /// </summary>
        [DefaultValue(null)]
        public string Generator { get; set; }

        /// <summary>
        /// paper width,unit twips
        /// </summary>
        [DefaultValue(12240)]
        public int PaperWidth { get; set; }

        /// <summary>
        /// paper height,unit twips
        /// </summary>
        [DefaultValue(15840)]
        public int PaperHeight { get; set; }

        /// <summary>
        /// left margin,unit twips
        /// </summary>
        [DefaultValue(1800)]
        public int LeftMargin { get; set; }

        private int _intTopMargin = 1440;
        /// <summary>
        /// top margin,unit twips
        /// </summary>
        [DefaultValue(1440)]
        public int TopMargin
        {
            get
            {
                return _intTopMargin;
            }
            set
            {
                _intTopMargin = value;
            }
        }

        private int _intRightMargin = 1800;
        /// <summary>
        /// right margin,unit twips
        /// </summary>
        [DefaultValue(1800)]
        public int RightMargin
        {
            get
            {
                return _intRightMargin;
            }
            set
            {
                _intRightMargin = value;
            }
        }

        private int _intBottomMargin = 1440;
        /// <summary>
        /// bottom margin,unit twips
        /// </summary>
        [DefaultValue(1440)]
        public int BottomMargin
        {
            get
            {
                return _intBottomMargin;
            }
            set
            {
                _intBottomMargin = value;
            }
        }

        private bool _bolLandscape;
        /// <summary>
        /// landscape
        /// </summary>
        [DefaultValue(false)]
        public bool Landscape
        {
            get
            {
                return _bolLandscape;
            }
            set
            {
                _bolLandscape = value;
            }
        }

        private int _headerDistance = 720;
        /// <summary>
        /// Header's distance from the top of the page( Twips)
        /// </summary>
        [DefaultValue(720)]
        public int HeaderDistance
        {
            get
            {
                return _headerDistance;
            }
            set
            {
                _headerDistance = value;
            }
        }

        private int _footerDistance = 720;
        /// <summary>
        /// Footer's distance from the bottom of the page( twips)
        /// </summary>
        [DefaultValue(720)]
        public int FooterDistance
        {
            get
            {
                return _footerDistance;
            }
            set
            {
                _footerDistance = value;
            }
        }
        /// <summary>
        /// client area width,unit twips
        /// </summary>
        public int ClientWidth
        {
            get
            {
                if (_bolLandscape)
                {
                    return PaperHeight - LeftMargin - _intRightMargin;
                }
                else
                {
                    return PaperWidth - LeftMargin - _intRightMargin;
                }
            }
        }

        private bool _bolChangeTimesNewRoman;
        /// <summary>
        /// convert "Times new roman" to default font when parse rtf content
        /// </summary>
        [DefaultValue(true)]
        public bool ChangeTimesNewRoman
        {
            get
            {
                return _bolChangeTimesNewRoman;
            }
            set
            {
                _bolChangeTimesNewRoman = value;
            }
        }
        //private Stack myElements = new Stack();

        /// <summary>
        /// progress event
        /// </summary>
        public event ProgressEventHandler Progress;

        /// <summary>
        /// raise progress event
        /// </summary>
        /// <param name="max">progress max value</param>
        /// <param name="value">progress value</param>
        /// <param name="message">progress message</param>
        /// <returns>user cancel</returns>
        protected bool OnProgress(int max, int value, string message)
        {
            if (Progress != null)
            {
                ProgressEventArgs args = new ProgressEventArgs(max, value, message);
                Progress(this, args);
                return args.Cancel;
            }
            return false;
        }


        ///// <summary>
        ///// load a rtf file and parse
        ///// </summary>
        ///// <param name="fileName">file name</param>
        //public void Load(string fileName)
        //{
        //    using (System.IO.FileStream stream = new System.IO.FileStream(
        //        fileName,
        //        System.IO.FileMode.Open, System.IO.FileAccess.Read))
        //    {
        //        Load(stream);
        //    }
        //}

        /// <summary>
        /// load a rtf document from a stream and parse content
        /// </summary>
        /// <param name="stream">stream</param>
        public void Load(System.IO.Stream stream)
        {
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true ;
            HtmlContent = null;
            Elements.Clear();
            _bolStartContent = false;
            RTFReader reader = new RTFReader(stream);
            DocumentFormatInfo format = new DocumentFormatInfo();
            _paragraphFormat = null;
            Load(reader, format);
            // combination table rows to table
            CombineTable(this);
            FixElements(this);
            FixRTFHtml();
        }


        /// <summary>
        /// load a rtf document from a text reader and parse content
        /// </summary>
        /// <param name="reader">text reader</param>
        public void Load(System.IO.TextReader reader)
        {
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true;
            HtmlContent = null;
            Elements.Clear();
            _bolStartContent = false;
            RTFReader r = new RTFReader(reader);
            DocumentFormatInfo format = new DocumentFormatInfo();
            _paragraphFormat = null;
            Load(r, format);
            // combination table rows to table
            CombineTable(this);
            FixElements(this);
            FixRTFHtml();
        }

        /// <summary>
        /// load a rtf document from a string in rtf format and parse content
        /// </summary>
        /// <param name="rtfText">text</param>
        public void LoadRTFText(string rtfText)
        {
            System.IO.StringReader reader = new System.IO.StringReader(rtfText);
            //_HtmlContentBuilder = new StringBuilder();
            //_RTFHtmlState = true;
            HtmlContent = null;
            Elements.Clear();
            _bolStartContent = false;
            RTFReader rtfReader = new RTFReader(reader);
            DocumentFormatInfo format = new DocumentFormatInfo();
            _paragraphFormat = null;
            Load(rtfReader, format);
            CombineTable(this);
            FixElements(this);
            FixRTFHtml();
        }

        private void FixRTFHtml()
        {
            //if (_HtmlContentBuilder != null)
            //{
            //    _HtmlContent = _HtmlContentBuilder.ToString();
            //    _HtmlContentBuilder = null;
            //}
            //_RTFHtmlState = false;
        }

        public void FixForParagraphs(RTFDomElement parentElement)
        {
            RTFDomParagraph lastParagraph = null;
            List<RTFDomElement> list = new List<RTFDomElement>();
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomHeader
                    || element is RTFDomFooter)
                {
                    FixForParagraphs(element);
                    lastParagraph = null;
                    list.Add(element);
                    continue;
                }
                if (element is RTFDomParagraph
                    || element is RTFDomTableRow
                    || element is RTFDomTable
                    || element is RTFDomTableCell)
                {
                    lastParagraph = null;
                    list.Add(element);
                    continue;
                }
                if (lastParagraph == null)
                {
                    lastParagraph = new RTFDomParagraph();
                    list.Add(lastParagraph);
                    if (element is RTFDomText)
                    {
                        lastParagraph.Format = ((RTFDomText)element).Format.Clone();
                    }
                }
                lastParagraph.Elements.Add(element);
            }//foreach
            parentElement.Elements.Clear();
            foreach (RTFDomElement element in list)
            {
                //if (element is RTFDomHeader
                //    || element is RTFDomFooter)
                //{
                //    FixForParagraphs(element);
                //}
                parentElement.Elements.Add(element);
            }
        }

        private void FixElements(RTFDomElement parentElement)
        {
            // combin text element , decrease number of RTFDomText instance
            List<RTFDomElement> result = new List<RTFDomElement>();
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomParagraph)
                {
                    RTFDomParagraph p = (RTFDomParagraph)element;
                    if (p.Format.PageBreak)
                    {
                        p.Format.PageBreak = false;
                        result.Add(new RTFDomPageBreak());
                    }
                }
                if (element is RTFDomText)
                {
                    if (result.Count > 0 && result[result.Count - 1] is RTFDomText)
                    {
                        RTFDomText lastText = (RTFDomText)result[result.Count - 1];
                        RTFDomText txt = (RTFDomText)element;
                        if (lastText.Text.Length == 0 || txt.Text.Length == 0)
                        {
                            if (lastText.Text.Length == 0)
                            {
                                // close text format
                                lastText.Format = txt.Format.Clone();
                            }
                            lastText.Text = lastText.Text + txt.Text;
                        }
                        else
                        {
                            if (lastText.Format.EqualsSettings(txt.Format))
                            {
                                lastText.Text = lastText.Text + txt.Text;
                            }
                            else
                            {
                                result.Add(txt);
                            }
                        }
                    }
                    else
                    {
                        result.Add(element);
                    }
                }
                else
                {
                    result.Add(element);
                }
            }//foreach
            parentElement.Elements.Clear();
            parentElement.Locked = false;
            foreach (RTFDomElement element in result)
            {
                parentElement.AppendChild(element);
            }

            foreach (RTFDomElement element in parentElement.Elements.ToArray())
            {
                if (element is RTFDomTable)
                {
                    UpdateTableCells((RTFDomTable)element, true);
                }
            }


            //RTFDomParagraph tempP = null;
            //RTFDomParagraph lastP = null;
            //foreach (RTFDomElement element in parentElement.Elements)
            //{
            //    if (element is RTFDomParagraph)
            //    {
            //        RTFDomParagraph p = (RTFDomParagraph)element;
            //        if (p.TemplateGenerated)
            //        {
            //            tempP = p;
            //        }
            //        else
            //        {
            //            if (tempP != null)
            //            {
            //                tempP.TemplateGenerated = false;
            //                tempP.Format = p.Format;
            //                tempP = p;
            //            }
            //        }
            //        lastP = p;
            //    }
            //}//foreach
            //if (tempP != null && lastP != null)
            //{
            //    if (lastP.Elements.Count == 0)
            //    {
            //        parentElement.Elements.Remove(lastP);
            //    }
            //}

            foreach (RTFDomElement element in parentElement.Elements)
            {
                FixElements(element);
            }
        }

        private RTFDomElement[] GetLastElements(bool checkLockState)
        {
            List<RTFDomElement> result = new List<RTFDomElement>();
            RTFDomElement element = this;
            while (element != null)
            {
                if (checkLockState)
                {
                    if (element.Locked)
                    {
                        break;
                    }
                }
                result.Add(element);
                element = element.Elements.LastOrDefault();
            }
            if (checkLockState)
            {
                for (int iCount = result.Count - 1; iCount >= 0; iCount--)
                {
                    if (result[iCount].Locked)
                    {
                        result.RemoveAt(iCount);
                    }
                }
            }
            return result.ToArray();
        }

        public T GetLastElement<T>()
        {
            return GetLastElements(true).OfType<T>().LastOrDefault();
        }

        public T GetLastElement<T>(bool lockStatus) where T : RTFDomElement
        {
            return GetLastElements(true).OfType<T>().LastOrDefault(i => i.Locked == lockStatus);
        }

        public RTFDomElement GetLastElement()
        {
            RTFDomElement[] elements = GetLastElements(true);
            return elements[elements.Length - 1];
        }

        private void CompleteParagraph()
        {
            RTFDomElement lastElement = GetLastElement();
            while (lastElement != null)
            {
                if (lastElement is RTFDomParagraph)
                {
                    RTFDomParagraph p = (RTFDomParagraph)lastElement;
                    p.Locked = true;
                    if (_paragraphFormat != null)
                    {
                        p.Format = _paragraphFormat;
                        _paragraphFormat = _paragraphFormat.Clone();
                    }
                    else
                    {
                        _paragraphFormat = new DocumentFormatInfo();
                    }
                    break;
                }
                lastElement = lastElement.Parent;
            }
        }

        private void AddContentElement(RTFDomElement newElement)
        {
            RTFDomElement[] elements = GetLastElements(true);
            RTFDomElement lastElement = null;
            if (elements.Length > 0)
            {
                lastElement = elements[elements.Length - 1];
            }
            if (lastElement is RTFDomDocument
                || lastElement is RTFDomHeader
                || lastElement is RTFDomFooter)
            {
                if (newElement is RTFDomText
                    || newElement is RTFDomImage
                    || newElement is RTFDomObject
                    || newElement is RTFDomShape
                    || newElement is RTFDomShapeGroup)
                {
                    RTFDomParagraph p = new RTFDomParagraph();
                    if (lastElement.Elements.Count > 0)
                    {
                        p.TemplateGenerated = true;
                    }
                    if (_paragraphFormat != null)
                    {
                        p.Format = _paragraphFormat;
                    }
                    lastElement.AppendChild(p);
                    p.Elements.Add(newElement);
                    return;
                }
            }
            RTFDomElement element = elements[elements.Length - 1];
            //if ( newElement is RTFDomTableRow)
            //{
            //    System.Diagnostics.Debugger.Break();
            //}
            if (newElement != null && newElement.NativeLevel > 0)
            {
                for (int iCount = elements.Length - 1; iCount >= 0; iCount--)
                {
                    if (elements[iCount].NativeLevel == newElement.NativeLevel)
                    {
                        for (int iCount2 = iCount; iCount2 < elements.Length; iCount2++)
                        {
                            RTFDomElement e2 = elements[iCount2];
                            //if (newElement.GetType().Equals(e2.GetType()))
                            //{
                            //}
                            if (newElement is RTFDomText
                                || newElement is RTFDomImage
                                || newElement is RTFDomObject
                                || newElement is RTFDomShape
                                || newElement is RTFDomShapeGroup
                                || newElement is RTFDomField
                                || newElement is RTFDomBookmark
                                || newElement is RTFDomLineBreak)
                            {
                                if (newElement.NativeLevel == e2.NativeLevel)
                                {
                                    if (e2 is RTFDomTableRow
                                        || e2 is RTFDomTableCell
                                        || e2 is RTFDomField
                                        || e2 is RTFDomParagraph)
                                    {
                                        continue;
                                    }
                                }
                            }

                            elements[iCount2].Locked = true;
                        }
                        break;
                    }
                }
            }

            for (int iCount = elements.Length - 1; iCount >= 0; iCount--)
            {
                if (elements[iCount].Locked == false)
                {
                    element = elements[iCount];
                    if (element is RTFDomImage)
                    {
                        element.Locked = true;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            if (element is RTFDomTableRow)
            {
                // If the last element is table row 
                // can not contains any element , 
                // so need create a cell element.
                RTFDomTableCell cell = new RTFDomTableCell();
                cell.NativeLevel = element.NativeLevel;
                element.AppendChild(cell);
                if (newElement is RTFDomTableRow)
                {
                    cell.Elements.Add(newElement);
                }
                else
                {
                    RTFDomParagraph cellP = new RTFDomParagraph();
                    cellP.Format = _paragraphFormat.Clone();
                    cellP.NativeLevel = cell.NativeLevel;
                    cell.AppendChild(cellP);
                    if (newElement != null)
                    {
                        cellP.AppendChild(newElement);
                    }
                }
            }
            else
            {
                if (newElement != null)
                {
                    if (element is RTFDomParagraph &&
                        (newElement is RTFDomParagraph
                        || newElement is RTFDomTableRow))
                    {
                        // If both is paragraph , append new paragraph to the parent of old paragraph
                        element.Locked = true;
                        element.Parent.AppendChild(newElement);
                    }
                    else
                    {
                        element.AppendChild(newElement);
                    }
                }
            }
        }



        private int _listTextFlag;
        private bool _bolStartContent;

        /// <summary>
        /// convert a hex string to a byte array
        /// </summary>
        /// <param name="hex">hex string</param>
        /// <returns>byte array</returns>
        private byte[] HexToBytes(string hex)
        {
            string chars = "0123456789abcdef";

            int index;
            int value = 0;
            int charCount = 0;
            ByteBuffer buffer = new ByteBuffer();
            for (int iCount = 0; iCount < hex.Length; iCount++)
            {
                char c = hex[iCount];
                c = char.ToLower(c);
                index = chars.IndexOf(c);
                if (index >= 0)
                {
                    charCount++;
                    value = value * 16 + index;
                    if (charCount > 0 && (charCount % 2) == 0)
                    {
                        buffer.Add((byte)value);
                        value = 0;
                    }
                }
            }
            return buffer.ToArray();
        }

        private void CombineTable(RTFDomElement parentElement)
        {
            List<RTFDomElement> result = new List<RTFDomElement>();
            List<RTFDomTableRow> rows = new List<RTFDomTableRow>();
            int lastRowWidth = -1;
            RTFDomTableRow lastRow = null;
            foreach (RTFDomElement element in parentElement.Elements)
            {
                if (element is RTFDomTableRow)
                {
                    RTFDomTableRow row = (RTFDomTableRow)element;
                    row.Locked = false;
                    var cellSettings = row.CellSettings;
                    if (cellSettings.Count == 0)
                    {
                        if (lastRow != null && lastRow.CellSettings.Count == row.Elements.Count)
                        {
                            cellSettings = lastRow.CellSettings;
                        }
                    }
                    if (cellSettings.Count == row.Elements.Count)
                    {
                        for (int iCount = 0; iCount < row.Elements.Count; iCount++)
                        {
                            row.Elements[iCount].Attributes = (RTFAttributeList)cellSettings[iCount];
                        }
                    }
                    bool isLastRow = row.HasAttribute(RTFConsts.Lastrow);
                    if (isLastRow == false)
                    {
                        int index = parentElement.Elements.IndexOf(element);
                        if (index == parentElement.Elements.Count - 1)
                        {
                            // this element is the last element
                            // then this row is the last row
                            isLastRow = true;
                        }
                        else
                        {
                            RTFDomElement e2 = parentElement.Elements[index + 1];
                            if (!(e2 is RTFDomTableRow))
                            {
                                // next element is not row 
                                isLastRow = true;
                            }
                        }
                    }
                    // split to table
                    if (isLastRow)
                    {
                        // if current row mark the last row , then 
                        // generate a new table
                        rows.Add(row);
                        result.Add(CreateTable(rows));
                        lastRowWidth = -1;
                    }
                    else
                    {
                        int width = 0;
                        if (row.HasAttribute(RTFConsts.TrwWidth))
                        {
                            width = row.Attributes[RTFConsts.TrwWidth];
                            if (row.HasAttribute(RTFConsts.TrwWidthA))
                            {
                                width = width - row.Attributes[RTFConsts.TrwWidthA];
                            }
                        }
                        else
                        {
                            foreach (RTFDomTableCell cell in row.Elements)
                            {
                                if (cell.HasAttribute(RTFConsts.Cellx))
                                {
                                    width = Math.Max(width, cell.Attributes[RTFConsts.Cellx]);
                                }
                            }
                        }
                        if (lastRowWidth > 0 && lastRowWidth != width)
                        {
                            // If row's width is change , then can consider multi-table combin
                            // then split and generate new table
                            if (rows.Count > 0)
                            {
                                result.Add(CreateTable(rows));
                            }
                        }
                        lastRowWidth = width;
                        rows.Add(row);
                    }
                    lastRow = row;
                }
                else if (element is RTFDomTableCell)
                {
                    lastRow = null;
                    CombineTable(element);
                    if (rows.Count > 0)
                    {
                        result.Add(CreateTable(rows));
                    }
                    result.Add(element);
                    lastRowWidth = -1;
                }
                else
                {
                    lastRow = null;
                    CombineTable(element);
                    if (rows.Count > 0)
                    {
                        result.Add(CreateTable(rows));
                    }
                    result.Add(element);
                    lastRowWidth = -1;
                }
            }//foreach
            if (rows.Count > 0)
            {
                result.Add(CreateTable(rows));
            }
            parentElement.Locked = false;
            parentElement.Elements.Clear();
            foreach (RTFDomElement element in result)
            {
                parentElement.AppendChild(element);
            }

        }
        /// <summary>
        /// create table
        /// </summary>
        /// <param name="rows">table rows</param>
        /// <returns>new table</returns>
        private RTFDomTable CreateTable(List<RTFDomTableRow> rows)
        {
            if (rows.Count > 0)
            {
                RTFDomTable table = new RTFDomTable();
                int index = 0;
                foreach (RTFDomTableRow row in rows)
                {
                    row.RowIndex = index;
                    index++;
                    table.AppendChild(row);
                }
                rows.Clear();
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        CombineTable(cell);
                    }
                }
                return table;
            }
            else
            {
                throw new ArgumentException("rows");
            }
        }

        private int _intDefaultRowHeight = 400;
        /// <summary>
        /// default row's height, in twips.
        /// </summary>
        public int DefaultRowHeight
        {
            get
            {
                return _intDefaultRowHeight;
            }
            set
            {
                _intDefaultRowHeight = value;
            }
        }

        private void UpdateTableCells(RTFDomTable table, bool fixTableCellSize)
        {
            // number of table column
            int columns = 0;
            // flag of cell merge
            bool merge = false;
            // right position of all cells
            List<int> rights = new List<int>();

            // right position of table
            int tableLeft = 0;
            for (int iCount = table.Elements.Count - 1; iCount >= 0; iCount--)
            {
                RTFDomTableRow row = (RTFDomTableRow)table.Elements[iCount];
                if (row.Elements.Count == 0)
                {
                    table.Elements.RemoveAt(iCount);
                }
            }
            if (table.Elements.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("");
            }
            foreach (RTFDomTableRow row in table.Elements)
            {
                int lastCellX = 0;

                columns = Math.Max(columns, row.Elements.Count);
                if (row.HasAttribute(RTFConsts.Irow))
                {
                    row.RowIndex = row.Attributes[RTFConsts.Irow];
                }
                row.IsLastRow = row.HasAttribute(RTFConsts.Lastrow);
                row.Header = row.HasAttribute(RTFConsts.Trhdr);
                // read row height
                if (row.HasAttribute(RTFConsts.Trrh))
                {
                    row.Height = row.Attributes[RTFConsts.Trrh];
                    if (row.Height == 0)
                    {
                        row.Height = DefaultRowHeight;
                    }
                    else if (row.Height < 0)
                    {
                        row.Height = -row.Height;
                    }
                }
                else
                {
                    row.Height = DefaultRowHeight;
                }
                // read default padding of cell
                if (row.HasAttribute(RTFConsts.Trpaddl))
                {
                    row.PaddingLeft = row.Attributes[RTFConsts.Trpaddl];
                }
                else
                {
                    row.PaddingLeft = int.MinValue;
                }
                if (row.HasAttribute(RTFConsts.Trpaddt))
                {
                    row.PaddingTop = row.Attributes[RTFConsts.Trpaddt];
                }
                else
                {
                    row.PaddingTop = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts.Trpaddr))
                {
                    row.PaddingRight = row.Attributes[RTFConsts.Trpaddr];
                }
                else
                {
                    row.PaddingRight = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts.Trpaddb))
                {
                    row.PaddingBottom = row.Attributes[RTFConsts.Trpaddb];
                }
                else
                {
                    row.PaddingBottom = int.MinValue;
                }

                if (row.HasAttribute(RTFConsts.Trleft))
                {
                    tableLeft = row.Attributes[RTFConsts.Trleft];
                }
                if (row.HasAttribute(RTFConsts.Trcbpat))
                {
                    row.Format.BackColor = ColorTable.GetColor(
                        row.Attributes[RTFConsts.Trcbpat],
                        Color.Transparent);
                }
                int widthCount = 0;
                foreach (RTFDomTableCell cell in row.Elements)
                {
                    // set cell's dispaly format

                    if (cell.HasAttribute(RTFConsts.Clvmgf))
                    {
                        // cell vertically merge
                        merge = true;
                    }
                    if (cell.HasAttribute(RTFConsts.Clvmrg))
                    {
                        // cell vertically merge by another cell
                        merge = true;
                    }
                    if (cell.HasAttribute(RTFConsts.Clpadl))
                    {
                        cell.PaddingLeft = cell.Attributes[RTFConsts.Clpadl];
                    }
                    else
                    {
                        cell.PaddingLeft = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts.Clpadr))
                    {
                        cell.PaddingRight = cell.Attributes[RTFConsts.Clpadr];
                    }
                    else
                    {
                        cell.PaddingRight = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts.Clpadt))
                    {
                        cell.PaddingTop = cell.Attributes[RTFConsts.Clpadt];
                    }
                    else
                    {
                        cell.PaddingTop = int.MinValue;
                    }
                    if (cell.HasAttribute(RTFConsts.Clpadb))
                    {
                        cell.PaddingBottom = cell.Attributes[RTFConsts.Clpadb];
                    }
                    else
                    {
                        cell.PaddingBottom = int.MinValue;
                    }

                    // whether dispaly border line
                    cell.Format.LeftBorder = cell.HasAttribute(RTFConsts.Clbrdrl);
                    cell.Format.TopBorder = cell.HasAttribute(RTFConsts.Clbrdrt);
                    cell.Format.RightBorder = cell.HasAttribute(RTFConsts.Clbrdrr);
                    cell.Format.BottomBorder = cell.HasAttribute(RTFConsts.Clbrdrb);
                    if (cell.HasAttribute(RTFConsts.Brdrcf))
                    {
                        cell.Format.BorderColor = ColorTable.GetColor(
                            cell.GetAttributeValue(RTFConsts.Brdrcf, 1),
                            Color.Black);
                    }
                    for (int iCount = cell.Attributes.Count - 1; iCount >= 0; iCount--)
                    {
                        string name3 = cell.Attributes[iCount].Name;
                        if (name3 == RTFConsts.Brdrtbl
                            || name3 == RTFConsts.Brdrnone
                            || name3 == RTFConsts.Brdrnil)
                        {
                            for (int iCount2 = iCount - 1; iCount2 >= 0; iCount2--)
                            {
                                string name2 = cell.Attributes[iCount2].Name;
                                if (name2 == RTFConsts.Clbrdrl)
                                {
                                    cell.Format.LeftBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts.Clbrdrt)
                                {
                                    cell.Format.TopBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts.Clbrdrr)
                                {
                                    cell.Format.RightBorder = false;
                                    break;
                                }
                                else if (name2 == RTFConsts.Clbrdrb)
                                {
                                    cell.Format.BottomBorder = false;
                                    break;
                                }
                            }
                        }
                    }

                    if (cell.HasAttribute(RTFConsts.Clvertalt))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Top;
                    }
                    else if (cell.HasAttribute(RTFConsts.Clvertalc))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Middle;
                    }
                    else if (cell.HasAttribute(RTFConsts.Clvertalb))
                    {
                        cell.VerticalAlignment = RTFVerticalAlignment.Bottom;
                    }
                    // background color
                    if (cell.HasAttribute(RTFConsts.Clcbpat))
                    {
                        cell.Format.BackColor = ColorTable.GetColor(cell.Attributes[RTFConsts.Clcbpat], Color.Transparent);
                    }
                    else
                    {
                        cell.Format.BackColor = Color.Transparent;
                    }
                    if (cell.HasAttribute(RTFConsts.Clcfpat))
                    {
                        cell.Format.BorderColor = ColorTable.GetColor(cell.Attributes[RTFConsts.Clcfpat], Color.Black);
                    }

                    // cell's width
                    int cellWidth = 2763;// cell's default with is 2763 Twips(570 Document)
                    if (cell.HasAttribute(RTFConsts.Cellx))
                    {
                        cellWidth = cell.Attributes[RTFConsts.Cellx] - lastCellX;
                        if (cellWidth < 100)
                        {
                            cellWidth = 100;
                        }
                    }
                    int right = lastCellX + cellWidth;
                    // fix cell's right position , if this position is very near with another cell's 
                    // right position( less then 45 twips or 3 pixel), then consider these two position
                    // is the same , this can decrease number of table columns
                    for (int iCount = 0; iCount < rights.Count; iCount++)
                    {
                        if (Math.Abs(right - rights[iCount]) < 45)
                        {
                            right = rights[iCount];
                            cellWidth = right - lastCellX;
                            break;
                        }
                    }

                    cell.Left = lastCellX;
                    cell.Width = cellWidth;
                    if (cell.HasAttribute(RTFConsts.Cellx) == false)
                    {

                    }
                    widthCount += cellWidth;
                    //int right = cell.Left + cell.Width;
                    if (rights.Contains(right) == false)
                    {
                        // becase of convert twips to unit of document may cause truncation error.
                        // This may cause rights.Contains mistake . so scale cell's with with 
                        // native twips unit , after all computing , convert to unit of document.
                        rights.Add(right);
                    }
                    lastCellX = lastCellX + cellWidth;
                }
                row.Width = widthCount;
            }
            if (rights.Count == 0)
            {
                // can not detect cell's width , so consider set cell's width
                // automatic, then set cell's default width.
                int cols = 1;
                foreach (RTFDomTableRow row in table.Elements)
                {
                    cols = Math.Max(cols, row.Elements.Count);
                }
                int w = (ClientWidth / cols);
                for (int iCount = 0; iCount < cols; iCount++)
                {
                    rights.Add(iCount * w + w);
                }
            }
            // computing cell's rowspan and colspan , number of rights array is the number of table columns.

            rights.Add(0);
            rights.Sort();
            // add table column instance
            for (int iCount = 1; iCount < rights.Count; iCount++)
            {
                RTFDomTableColumn col = new RTFDomTableColumn();
                col.Width = rights[iCount] - rights[iCount - 1];
                table.Columns.Add(col);
            }

            for (int rowIndex = 1; rowIndex < table.Elements.Count; rowIndex++)
            {
                RTFDomTableRow row = (RTFDomTableRow)table.Elements[rowIndex];
                for (int colIndex = 0; colIndex < row.Elements.Count; colIndex++)
                {
                    RTFDomTableCell cell = (RTFDomTableCell)row.Elements[colIndex];
                    if (cell.Width == 0)
                    {
                        // If current cell not special width , then use the width of cell which 
                        // in the same colum and in the last row
                        RTFDomTableRow preRow = (RTFDomTableRow)table.Elements[rowIndex - 1];
                        if (preRow.Elements.Count > colIndex)
                        {
                            RTFDomTableCell preCell = (RTFDomTableCell)preRow.Elements[colIndex];
                            cell.Left = preCell.Left;
                            cell.Width = preCell.Width;
                            CopyStyleAttribute(cell, preCell.Attributes);
                        }
                    }
                }
            }
            if (merge == false)
            {
                // If not detect cell merge , maby exist cell merge in the same row
                foreach (RTFDomTableRow row in table.Elements)
                {
                    if (row.Elements.Count < table.Columns.Count)
                    {
                        // if number of row's cells not equals the number of table's columns
                        // then exist cell merge.
                        merge = true;
                        break;
                    }
                }
            }
            if (merge)
            {
                // detect cell merge , begin merge operation

                // Becase of in rtf format,cell which merged by another cell in the same row , 
                // does no written in rtf text , so delay create those cell instance .
                foreach (RTFDomTableRow row in table.Elements)
                {
                    if (row.Elements.Count != table.Columns.Count)
                    {
                        // If number of row's cells not equals number of table's columns ,
                        // then consider there are hanppend  horizontal merge.
                        RTFDomElement[] cells = row.Elements.ToArray();
                        foreach (RTFDomTableCell cell in cells)
                        {
                            int index = rights.IndexOf(cell.Left);
                            int index2 = rights.IndexOf(cell.Left + cell.Width);
                            int intColSpan = index2 - index;
                            // detect vertical merge
                            bool verticalMerge = cell.HasAttribute(RTFConsts.Clvmrg);
                            if (verticalMerge == false)
                            {
                                // If this cell does not merged by another cell abover , 
                                // then set colspan
                                cell.ColSpan = intColSpan;
                            }
                            if (row.Elements.LastOrDefault() == cell)
                            {
                                cell.ColSpan = table.Columns.Count - row.Elements.Count + 1;
                                intColSpan = cell.ColSpan;
                            }
                            for (int iCount = 0; iCount < intColSpan - 1; iCount++)
                            {
                                RTFDomTableCell newCell = new RTFDomTableCell();
                                newCell.Attributes = cell.Attributes.Clone();
                                row.Elements.Insert(row.Elements.IndexOf(cell) + 1, newCell);
                                if (verticalMerge)
                                {
                                    // This cell has been merged.
                                    newCell.Attributes[RTFConsts.Clvmrg] = 1;
                                    newCell.OverrideCell = cell;
                                }
                            }//for
                        }
                        if (row.Elements.Count != table.Columns.Count)
                        {
                            // If the last cell has been merged. then supply new cells.
                            RTFDomTableCell lastCell = (RTFDomTableCell)row.Elements.LastOrDefault();
                            //if (lastCell.OverrideCell == null && lastCell.ColSpan > 1)
                            //{
                            //    lastCell.ColSpan = table.Columns.Count - row.Elements.IndexOf(lastCell);
                            //}
                            for (int iCount = row.Elements.Count; iCount < rights.Count; iCount++)
                            {
                                RTFDomTableCell newCell = new RTFDomTableCell();
                                CopyStyleAttribute(newCell, lastCell.Attributes);
                                row.Elements.Add(newCell);
                            }
                        }//if
                    }//if
                }//foreach

                // set cell's vertial merge.
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        if (cell.HasAttribute(RTFConsts.Clvmgf) == false)
                        {
                            //if this cell does not mark vertial merge , then next cell
                            continue;
                        }
                        // if this cell mark vertial merge.
                        int colIndex = row.Elements.IndexOf(cell);
                        for (int rowIndex = table.Elements.IndexOf(row) + 1;
                            rowIndex < table.Elements.Count;
                            rowIndex++)
                        {
                            RTFDomTableRow row2 = (RTFDomTableRow)table.Elements[rowIndex];
                            RTFDomTableCell cell2 = (RTFDomTableCell)row2.Elements[colIndex];
                            if (cell2.HasAttribute(RTFConsts.Clvmrg))
                            {
                                if (cell2.OverrideCell != null)
                                {
                                    // If this cell has been merge by another cell( must in the same row )
                                    // then break the circle
                                    break;
                                }
                                // increase vertial merge.
                                cell.RowSpan++;
                                cell2.OverrideCell = cell;
                            }//if
                            else
                            {
                                // if this cell not mark merged by another cell , then break the circel
                                break;
                            }
                        }//for
                    }
                }

                // set cell's OverridedCell information
                foreach (RTFDomTableRow row in table.Elements)
                {
                    foreach (RTFDomTableCell cell in row.Elements)
                    {
                        if (cell.RowSpan > 1 || cell.ColSpan > 1)
                        {
                            for (int rowIndex = 1; rowIndex <= cell.RowSpan; rowIndex++)
                            {
                                for (int colIndex = 1; colIndex <= cell.ColSpan; colIndex++)
                                {
                                    int r = table.Elements.IndexOf(row) + rowIndex - 1;
                                    int c = row.Elements.IndexOf(cell) + colIndex - 1;
                                    RTFDomTableCell cell2 = (RTFDomTableCell)table.Elements[r].Elements[c];
                                    if (cell != cell2)
                                    {
                                        cell2.OverrideCell = cell;
                                    }
                                }//for
                            }//for
                        }//if
                    }//foreach
                }//foreach

            }//if

            if (fixTableCellSize)
            {

                // Fix table's left position use the first table column
                if (table.Columns.Count > 0)
                {
                    ((RTFDomTableColumn)table.Columns[0]).Width -= tableLeft;
                }
            }

        }

        private void CopyStyleAttribute(RTFDomTableCell cell, RTFAttributeList table)
        {
            RTFAttributeList attrs = table.Clone();
            attrs.Remove(RTFConsts.Clvmgf);
            attrs.Remove(RTFConsts.Clvmrg);
            cell.Attributes = attrs;
        }

        public override string ToString()
        {
            return "RTFDocument:" + _myInfo.Title;
        }

        private bool ApplyText(
            RTFTextContainer myText,
            RTFReader reader,
            DocumentFormatInfo format)
        {
            if (myText.HasContent)
            {
                string strText = myText.Text;
                myText.Clear();
                //if (_RTFHtmlState == false)
                //{
                //    _HtmlContentBuilder.Append(strText);
                //    return false ;
                //}
                // if current element is image element , then finish handle image element
                RTFDomImage img = GetLastElement<RTFDomImage>();
                if (img != null && img.Locked == false)
                {
                    img.Data = HexToBytes(strText);
                    img.Format = format.Clone();
                    img.Width = (img.DesiredWidth * img.ScaleX / 100);
                    img.Height = (img.DesiredHeight * img.ScaleY / 100);
                    img.Locked = true;
                    if (reader.TokenType != RTFTokenType.GroupEnd)
                    {
                        ReadToEndGround(reader);
                    }
                    return true;
                }
                else if (format.ReadText && _bolStartContent)
                {
                    RTFDomText txt = new RTFDomText();
                    txt.NativeLevel = myText.Level;
                    txt.Format = format.Clone();
                    if (txt.Format.Align == RTFAlignment.Justify)
                        txt.Format.Align = RTFAlignment.Left;
                    txt.Text = strText;
                    AddContentElement(txt);
                }
            }
            return false;
        }


        private int _intTokenCount;
        private DocumentFormatInfo _paragraphFormat;
        private void Load(RTFReader reader, DocumentFormatInfo parentFormat)
        {
            bool forbitPard = false;
            DocumentFormatInfo format;
            if (_paragraphFormat == null)
            {
                _paragraphFormat = new DocumentFormatInfo();
            }
            if (parentFormat == null)
            {
                format = new DocumentFormatInfo();
            }
            else
            {
                format = parentFormat.Clone();
                format.NativeLevel = parentFormat.NativeLevel + 1;
            }
            RTFTextContainer myText = new RTFTextContainer(this);
            int levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenCount - _intTokenCount > 100)
                {
                    _intTokenCount = reader.TokenCount;
                    OnProgress(reader.ContentLength, reader.ContentPosition, null);
                }
                if (_bolStartContent)
                {
                    if (myText.Accept(reader.CurrentToken, reader))
                    {
                        myText.Level = reader.Level;
                        continue;
                    }
                    else if (myText.HasContent)
                    {
                        if (ApplyText(myText, reader, format))
                        {
                            break;
                        }
                    }
                }

                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    RTFDomElement[] elements = GetLastElements(true);
                    for (int iCount = 0; iCount < elements.Length; iCount++)
                    {
                        RTFDomElement element = elements[iCount];
                        if (element.NativeLevel >= 0 && element.NativeLevel > reader.Level)
                        {
                            for (int iCount2 = iCount; iCount2 < elements.Length; iCount2++)
                            {
                                elements[iCount2].Locked = true;
                            }
                            break;
                        }
                    }//for

                    break;
                }

                if (reader.Level < levelBack)
                {
                    break;
                }

                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    //level++;
                    Load(reader, format);
                    if (reader.Level < levelBack)
                    {
                        break;
                    }
                    //continue;
                }

                if (reader.TokenType == RTFTokenType.Control
                    || reader.TokenType == RTFTokenType.Keyword
                    || reader.TokenType == RTFTokenType.ExtKeyword)
                {


                    switch (reader.Keyword)
                    {
                        case "fromhtml":
                            ReadHtmlContent(reader);
                            return;
                        case RTFConsts.Listtable:
                            ReadListTable(reader);
                            return;
                        case RTFConsts.Listoverride:
                            // unknow keyword
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts.Ansi:
                            break;
                        case RTFConsts.Ansicpg:
                            // read default encoding
                            //todo: look up in encoding dictionary
                            //myDefaultEncoding = Encoding.GetEncoding(reader.Parameter);
                            break;
                        case RTFConsts.Fonttbl:
                            // read font table
                            ReadFontTable(reader);
                            break;
                        case "listoverridetable":
                            ReadListOverrideTable(reader);
                            break;
                        case "filetbl":
                            // unsupport file list
                            ReadToEndGround(reader);
                            break;// finish current level
                        //break;
                        case RTFConsts.Colortbl:
                            // read color table
                            ReadColorTable(reader);
                            return;// finish current level
                        //break;
                        case "stylesheet":
                            // unsupport style sheet list
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts.Generator:
                            // read document generator
                            Generator = ReadInnerText(reader, true);
                            break;
                        case RTFConsts.Info:
                            // read document information
                            ReadDocumentInfo(reader);
                            return;
                        case RTFConsts.Headery:
                            {
                                if (reader.HasParam)
                                {
                                    HeaderDistance = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts.Footery:
                            {
                                if (reader.HasParam)
                                {
                                    FooterDistance = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts.Header:
                            {
                                // analyse header
                                RTFDomHeader header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.AllPages;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Headerl:
                            {
                                // analyse header
                                RTFDomHeader header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.LeftPages;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Headerr:
                            {
                                // analyse header
                                RTFDomHeader header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.RightPages;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Headerf:
                            {
                                // analyse header
                                RTFDomHeader header = new RTFDomHeader();
                                header.Style = HeaderFooterStyle.FirstPage;
                                AppendChild(header);
                                Load(reader, parentFormat);
                                header.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Footer:
                            {
                                // analyse footer
                                RTFDomFooter footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.FirstPage;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Footerl:
                            {
                                // analyse footer
                                RTFDomFooter footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.LeftPages;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Footerr:
                            {
                                // analyse footer
                                RTFDomFooter footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.RightPages;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Footerf:
                            {
                                // analyse footer
                                RTFDomFooter footer = new RTFDomFooter();
                                footer.Style = HeaderFooterStyle.FirstPage;
                                AppendChild(footer);
                                Load(reader, parentFormat);
                                footer.Locked = true;
                                _paragraphFormat = new DocumentFormatInfo();
                                break;
                            }
                        case RTFConsts.Xmlns:
                            {
                                // unsupport xml namespace
                                ReadToEndGround(reader);
                                break;
                            }
                        case RTFConsts.Nonesttables:
                            {
                                // I support nest table , then ignore this keyword
                                ReadToEndGround(reader);
                                break;
                            }
                        case RTFConsts.Xmlopen:
                            {
                                // unsupport xmlopen keyword
                                break;
                            }
                        case RTFConsts.Revtbl:
                            {
                                //ReadToEndGround(reader);
                                break;
                            }


                        //**************** read document information ***********************
                        case RTFConsts.Paperw:
                            {
                                // read paper width
                                PaperWidth = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Paperh:
                            {
                                // read paper height
                                PaperHeight = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Margl:
                            {
                                // read left margin
                                LeftMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Margr:
                            {
                                // read right margin
                                _intRightMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Margb:
                            {
                                // read bottom margin
                                _intBottomMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Margt:
                            {
                                // read top margin 
                                _intTopMargin = reader.Parameter;
                                break;
                            }
                        case RTFConsts.Landscape:
                            {
                                // set landscape
                                _bolLandscape = true;
                                break;
                            }
                        case RTFConsts.Fchars:
                            FollowingChars = ReadInnerText(reader, true);
                            break;
                        case RTFConsts.Lchars:
                            LeadingChars = ReadInnerText(reader, true);
                            break;
                        case "pnseclvl":
                            // ignore this keyword
                            ReadToEndGround(reader);
                            break;
                        ////**************** read html content ***************************
                        //case "htmlrtf":
                        //    {
                        //        if (reader.HasParam && reader.Parameter == 0)
                        //        {
                        //            _RTFHtmlState = false;
                        //        }
                        //        else
                        //        {
                        //            _RTFHtmlState = true;
                        //            //while ( reader.PeekTokenType() == RTFTokenType.Text )
                        //            //{
                        //            //    reader.ReadToken();
                        //            //}
                        //            //string text = ReadInnerText(reader, null, false, true, false);
                        //        }
                        //    }
                        //    break;
                        //case "htmltag":
                        //    {
                        //        if (reader.InnerReader.Peek() == (int)' ')
                        //        {
                        //            reader.InnerReader.Read();
                        //        }
                        //        string text = ReadInnerText(reader, null, true, false, true);
                        //        if (string.IsNullOrEmpty(text) == false)
                        //        {
                        //            _HtmlContentBuilder.Append(text);
                        //        }
                        //        //while (true)
                        //        //{
                        //        //    int c = reader.InnerReader.Peek();
                        //        //    if (c < 0)
                        //        //    {
                        //        //        break;
                        //        //    }
                        //        //    if (c == '}')
                        //        //    {
                        //        //        break;
                        //        //    }
                        //        //    _HtmlContentBuilder.Append((char)c);
                        //        //    reader.InnerReader.Read();
                        //        //}
                        //    }
                        //    break;
                        //**************** read paragraph format ***********************
                        case RTFConsts.Pard:
                            {
                                _bolStartContent = true;
                                if (forbitPard)
                                    continue;
                                // clear paragraph format
                                _paragraphFormat.ResetParagraph();
                                //format.ResetParagraph();
                                break;
                            }
                        case RTFConsts.Par:
                            {
                                _bolStartContent = true;
                                // new paragraph
                                if (GetLastElement<RTFDomParagraph>() == null)
                                {
                                    RTFDomParagraph p = new RTFDomParagraph();
                                    p.Format = _paragraphFormat;
                                    _paragraphFormat = _paragraphFormat.Clone();
                                    AddContentElement(p);
                                    p.Locked = true;
                                }
                                else
                                {
                                    CompleteParagraph();
                                    RTFDomParagraph p = new RTFDomParagraph();
                                    p.Format = _paragraphFormat;
                                    AddContentElement(p);
                                }
                                _bolStartContent = true;
                                break;
                            }
                        case RTFConsts.Page:
                            {
                                _bolStartContent = true;
                                CompleteParagraph();
                                AddContentElement(new RTFDomPageBreak());
                                break;
                            }
                        case RTFConsts.Pagebb:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.PageBreak = true;
                                break;
                            }
                        case RTFConsts.Ql:
                            {
                                // left alignment
                                _bolStartContent = true;
                                _paragraphFormat.Align = RTFAlignment.Left;
                                break;
                            }
                        case RTFConsts.Qc:
                            {
                                // center alignment
                                _bolStartContent = true;
                                _paragraphFormat.Align = RTFAlignment.Center;
                                break;
                            }
                        case RTFConsts.Qr:
                            {
                                // right alignment
                                _bolStartContent = true;
                                _paragraphFormat.Align = RTFAlignment.Right;
                                break;
                            }
                        case RTFConsts.Qj:
                            {
                                // jusitify alignment
                                _bolStartContent = true;
                                _paragraphFormat.Align = RTFAlignment.Justify;
                                break;
                            }
                        case RTFConsts.Sl:
                            {
                                // line spacing
                                _bolStartContent = true;
                                if (reader.Parameter >= 0)
                                {
                                    _paragraphFormat.LineSpacing = reader.Parameter;
                                }
                            }
                            break;
                        case RTFConsts.Slmult:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.MultipleLineSpacing = (reader.Parameter == 1);
                            }
                            break;
                        case RTFConsts.Sb:
                            {
                                // spacing before paragraph
                                _bolStartContent = true;
                                _paragraphFormat.SpacingBefore = reader.Parameter;
                            }
                            break;
                        case RTFConsts.Sa:
                            {
                                // spacing after paragraph
                                _bolStartContent = true;
                                _paragraphFormat.SpacingAfter = reader.Parameter;
                            }
                            break;
                        case RTFConsts.Fi:
                            {
                                // indent first line
                                _bolStartContent = true;
                                _paragraphFormat.ParagraphFirstLineIndent = reader.Parameter;
                                //if (reader.Parameter >= 400)
                                //{
                                //    _ParagraphFormat.ParagraphFirstLineIndent = reader.Parameter; //doc.StandTabWidth;
                                //}
                                //else
                                //{
                                //    _ParagraphFormat.ParagraphFirstLineIndent = 0;
                                //}
                                ////UpdateParagraph( CurrentParagraphEOF , format );
                                break;
                            }
                        case RTFConsts.Brdrw:
                            {
                                _bolStartContent = true;
                                if (reader.HasParam)
                                {
                                    _paragraphFormat.BorderWidth = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Pn:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.ListId = -1;
                                break;
                            }

                        case RTFConsts.Pntext:
                            break;
                        case RTFConsts.Pntxtb:
                            break;
                        case RTFConsts.Pntxta:
                            break;

                        case RTFConsts.Pnlvlbody:
                            {
                                // numbered list style
                                _bolStartContent = true;
                                //_ParagraphFormat.NumberedList = true;
                                //_ParagraphFormat.BulletedList = false;
                                //if (_ParagraphFormat.Parent != null)
                                //{
                                //    _ParagraphFormat.Parent.NumberedList = format.NumberedList;
                                //    _ParagraphFormat.Parent.BulletedList = format.BulletedList;
                                //}
                                break;
                            }
                        case RTFConsts.Pnlvlblt:
                            {
                                // bulleted list style
                                _bolStartContent = true;
                                //_ParagraphFormat.NumberedList = false;
                                //_ParagraphFormat.BulletedList = true;
                                //if (_ParagraphFormat.Parent != null)
                                //{
                                //    _ParagraphFormat.Parent.NumberedList = format.NumberedList;
                                //    _ParagraphFormat.Parent.BulletedList = format.BulletedList;
                                //}
                                break;
                            }
                        case RTFConsts.Listtext:
                            {
                                _bolStartContent = true;
                                string txt = ReadInnerText(reader, true);
                                if (txt != null)
                                {
                                    txt = txt.Trim();
                                    if (txt.StartsWith("l"))
                                    {
                                        _listTextFlag = 1;
                                    }
                                    else
                                    {
                                        _listTextFlag = 2;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Ls:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.ListId = reader.Parameter;

                                //if (ListTextFlag == 1)
                                //{
                                //    _ParagraphFormat.NumberedList = false;
                                //    _ParagraphFormat.BulletedList = true;
                                //}
                                //else if (ListTextFlag == 2)
                                //{
                                //    _ParagraphFormat.NumberedList = true;
                                //    _ParagraphFormat.BulletedList = false;
                                //}
                                _listTextFlag = 0;
                                break;
                            }
                        case RTFConsts.Li:
                            {
                                _bolStartContent = true;
                                if (reader.HasParam)
                                {
                                    _paragraphFormat.LeftIndent = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Line:
                            {
                                _bolStartContent = true;
                                // break line
                                //if (_RTFHtmlState == false)
                                //{
                                //    _HtmlContentBuilder.Append(Environment.NewLine);
                                //}
                                //else
                                {
                                    if (format.ReadText)
                                    {
                                        RTFDomLineBreak line = new RTFDomLineBreak();
                                        line.NativeLevel = reader.Level;
                                        AddContentElement(line);
                                    }
                                }
                                break;
                            }
                        // ****************** read text format ******************************
                        case RTFConsts.Insrsid:
                            break;
                        case RTFConsts.Plain:
                            {
                                // clear text format
                                _bolStartContent = true;
                                format.ResetText();
                                break;
                            }
                        case RTFConsts.F:
                            {
                                // font name
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    string fontName = FontTable.GetFontName(reader.Parameter);
                                    if (fontName != null)
                                        fontName = fontName.Trim();
                                    if (fontName == null || fontName.Length == 0)
                                        fontName = _defaultFontName;

                                    if (ChangeTimesNewRoman)
                                    {
                                        if (fontName == "Times New Roman")
                                        {
                                            fontName = _defaultFontName;
                                        }
                                    }
                                    format.FontName = fontName;
                                }
                                _myFontChartset = FontTable[reader.Parameter].Encoding;
                                break;
                            }
                        case RTFConsts.Af:
                            {
                                _myAssociateFontChartset = FontTable[reader.Parameter].Encoding;
                                break;
                            }
                        case RTFConsts.Fs:
                            {
                                // font size
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.FontSize = reader.Parameter / 2.0f;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Cf:
                            {
                                // font color
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.TextColor = ColorTable.GetColor(reader.Parameter, Color.Black);
                                    }
                                }
                                break;
                            }

                        case RTFConsts.Cb:
                        case RTFConsts.Chcbpat:
                            {
                                // background color
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.BackColor = ColorTable.GetColor(reader.Parameter, Color.Empty);
                                    }
                                }
                                break;
                            }
                        case RTFConsts.B:
                            {
                                // bold
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Bold = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.V:
                            {
                                // hidden text
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam && reader.Parameter == 0)
                                    {
                                        format.Hidden = false;
                                    }
                                    else
                                    {
                                        format.Hidden = true;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Highlight:
                            {
                                // highlight content
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    if (reader.HasParam)
                                    {
                                        format.BackColor = ColorTable.GetColor(
                                            reader.Parameter,
                                            Color.Empty);
                                    }
                                }
                                break;
                            }
                        case RTFConsts.I:
                            {
                                // italic
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Italic = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.Ul:
                            {
                                // under line
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Underline = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.Strike:
                            {
                                // strikeout
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Strikeout = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.Sub:
                            {
                                // subscript
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Subscript = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.Super:
                            {
                                // superscript
                                _bolStartContent = true;
                                if (format.ReadText)
                                {
                                    format.Superscript = (reader.HasParam == false || reader.Parameter != 0);
                                }
                                break;
                            }
                        case RTFConsts.Nosupersub:
                            {
                                // nosupersub
                                _bolStartContent = true;
                                format.Subscript = false;
                                format.Superscript = false;
                                break;
                            }
                        case RTFConsts.Brdrb:
                            {
                                _bolStartContent = true;
                                //format.ParagraphBorder.Bottom = true;
                                _paragraphFormat.BottomBorder = true;
                                break;
                            }
                        case RTFConsts.Brdrl:
                            {
                                _bolStartContent = true;
                                //format.ParagraphBorder.Left = true ;
                                _paragraphFormat.LeftBorder = true;
                                break;
                            }
                        case RTFConsts.Brdrr:
                            {
                                _bolStartContent = true;
                                //format.ParagraphBorder.Right = true ;
                                _paragraphFormat.RightBorder = true;
                                break;
                            }
                        case RTFConsts.Brdrt:
                            {
                                _bolStartContent = true;
                                //format.ParagraphBorder.Top = true;
                                _paragraphFormat.BottomBorder = true;
                                break;
                            }
                        case RTFConsts.Brdrcf:
                            {
                                _bolStartContent = true;
                                RTFDomElement element = GetLastElement<RTFDomTableRow>(false);
                                if (element is RTFDomTableRow)
                                {
                                    // reading a table row
                                    RTFDomTableRow row = (RTFDomTableRow)element;
                                    RTFAttributeList style;
                                    if (row.CellSettings.Count > 0)
                                    {
                                        style = (RTFAttributeList)row.CellSettings[row.CellSettings.Count - 1];
                                        style.Add(reader.Keyword, reader.Parameter);
                                    }
                                    //else
                                    //{
                                    //    style = new RTFAttributeList();
                                    //    row.CellSettings.Add(style);
                                    //}

                                }
                                else
                                {
                                    _paragraphFormat.BorderColor = ColorTable.GetColor(reader.Parameter, Color.Black);
                                    format.BorderColor = format.BorderColor;
                                }
                                break;
                            }
                        case RTFConsts.Brdrs:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderThickness = false;
                                format.BorderThickness = false;
                                break;
                            }
                        case RTFConsts.Brdrth:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderThickness = true;
                                format.BorderThickness = true;
                                break;
                            }
                        case RTFConsts.Brdrdot:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderStyle = DashStyle.Dot;
                                format.BorderStyle = DashStyle.Dot;
                                break;
                            }
                        case RTFConsts.Brdrdash:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderStyle = DashStyle.Dash;
                                format.BorderStyle = DashStyle.Dash;
                                break;
                            }
                        case RTFConsts.Brdrdashd:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderStyle = DashStyle.DashDot;
                                format.BorderStyle = DashStyle.DashDot;
                                break;
                            }
                        case RTFConsts.Brdrdashdd:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.BorderStyle = DashStyle.DashDotDot;
                                format.BorderStyle = DashStyle.DashDotDot;
                                break;
                            }
                        case RTFConsts.Brdrnil:
                            {
                                _bolStartContent = true;
                                _paragraphFormat.LeftBorder = false;
                                _paragraphFormat.TopBorder = false;
                                _paragraphFormat.RightBorder = false;
                                _paragraphFormat.BottomBorder = false;

                                format.LeftBorder = false;
                                format.TopBorder = false;
                                format.RightBorder = false;
                                format.BottomBorder = false;
                                break;
                            }
                        case RTFConsts.Brsp:
                            {
                                _bolStartContent = true;
                                if (reader.HasParam)
                                {
                                    _paragraphFormat.BorderSpacing = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Chbrdr:
                            {
                                _bolStartContent = true;
                                format.LeftBorder = true;
                                format.TopBorder = true;
                                format.RightBorder = true;
                                format.BottomBorder = true;
                                break;
                            }
                        case RTFConsts.Bkmkstart:
                            {
                                // book mark
                                _bolStartContent = true;
                                if (format.ReadText && _bolStartContent)
                                {
                                    RTFDomBookmark bk = new RTFDomBookmark();
                                    bk.Name = ReadInnerText(reader, true);
                                    bk.Locked = true;
                                    AddContentElement(bk);
                                }
                                break;
                            }
                        case RTFConsts.Bkmkend:
                            {
                                forbitPard = true;
                                format.ReadText = false;
                                break;
                            }
                        case RTFConsts.Field:
                            {
                                // field
                                _bolStartContent = true;
                                ReadDomField(reader, format);
                                return; // finish current level
                                //break;
                            }

                        //case RTFConsts._objdata:
                        //case RTFConsts._objclass:
                        //    {
                        //        ReadToEndGround(reader);
                        //        break;
                        //    }

                        #region read object *********************************

                        case RTFConsts.Object:
                            {
                                // object
                                _bolStartContent = true;
                                ReadDomObject(reader, format);
                                return;// finish current level
                            }

                        #endregion

                        #region read image **********************************

                        case RTFConsts.Shppict:
                            // continue the following token
                            break;
                        case RTFConsts.Nonshppict:
                            // unsupport keyword
                            ReadToEndGround(reader);
                            break;
                        case RTFConsts.Pict:
                            {
                                // read image data
                                //ReadDomImage(reader, format);
                                _bolStartContent = true;
                                RTFDomImage img = new RTFDomImage();
                                img.NativeLevel = reader.Level;
                                AddContentElement(img);
                            }
                            break;
                        case RTFConsts.Picscalex:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.ScaleX = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Picscaley:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.ScaleY = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Picwgoal:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.DesiredWidth = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Pichgoal:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.DesiredHeight = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Blipuid:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.Id = ReadInnerText(reader, true);
                                }
                                break;
                            }
                        case RTFConsts.Emfblip:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Emfblip;
                                }
                                break;
                            }
                        case RTFConsts.Pngblip:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Pngblip;
                                }
                                break;
                            }
                        case RTFConsts.Jpegblip:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Jpegblip;
                                }
                                break;
                            }
                        case RTFConsts.Macpict:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Macpict;
                                }
                                break;
                            }
                        case RTFConsts.Pmmetafile:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Pmmetafile;
                                }
                                break;
                            }
                        case RTFConsts.Wmetafile:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Wmetafile;
                                }
                                break;
                            }
                        case RTFConsts.Dibitmap:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Dibitmap;
                                }
                                break;
                            }
                        case RTFConsts.Wbitmap:
                            {
                                RTFDomImage img = GetLastElement<RTFDomImage>();
                                if (img != null)
                                {
                                    img.PicType = RTFPicType.Wbitmap;
                                }
                                break;
                            }
                        #endregion

                        #region read shape ************************************************
                        case RTFConsts.Sp:
                            {
                                // begin read shape property
                                int level = 0;
                                string vName = null;
                                string vValue = null;
                                while (reader.ReadToken() != null)
                                {
                                    if (reader.TokenType == RTFTokenType.GroupStart)
                                    {
                                        level++;
                                    }
                                    else if (reader.TokenType == RTFTokenType.GroupEnd)
                                    {
                                        level--;
                                        if (level < 0)
                                        {
                                            break;
                                        }
                                    }
                                    else if (reader.Keyword == RTFConsts.Sn)
                                    {
                                        vName = ReadInnerText(reader, true);
                                    }
                                    else if (reader.Keyword == RTFConsts.Sv)
                                    {
                                        vValue = ReadInnerText(reader, true);
                                    }
                                }//while
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.ExtAttrbutes[vName] = vValue;
                                }
                                else
                                {
                                    RTFDomShapeGroup g = GetLastElement<RTFDomShapeGroup>();
                                    if (g != null)
                                    {
                                        g.ExtAttrbutes[vName] = vValue;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Shptxt:
                            {
                                // handle following token
                                break;
                            }
                        case RTFConsts.Shprslt:
                            {
                                // ignore this level
                                ReadToEndGround(reader);
                                break;
                            }
                        case RTFConsts.Shp:
                            {
                                _bolStartContent = true;
                                RTFDomShape shape = new RTFDomShape();
                                shape.NativeLevel = reader.Level;
                                AddContentElement(shape);
                                break;
                            }
                        case RTFConsts.Shpleft:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.Left = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Shptop:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.Top = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Shpright:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.Width = reader.Parameter - shape.Left;
                                }
                                break;
                            }
                        case RTFConsts.Shpbottom:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.Height = reader.Parameter - shape.Top;
                                }
                                break;
                            }
                        case RTFConsts.Shplid:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.ShapeId = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Shpz:
                            {
                                RTFDomShape shape = GetLastElement<RTFDomShape>();
                                if (shape != null)
                                {
                                    shape.ZIndex = reader.Parameter;
                                }
                                break;
                            }
                        case RTFConsts.Shpgrp:
                            {
                                RTFDomShapeGroup group = new RTFDomShapeGroup();
                                group.NativeLevel = reader.Level;
                                AddContentElement(group);
                                break;
                            }
                        case RTFConsts.Shpinst:
                            {
                                break;
                            }
                        #endregion

                        #region read table ************************************************
                        case RTFConsts.Intbl:
                        case RTFConsts.Trowd:
                        case RTFConsts.Itap:
                            {
                                // these keyword said than current paragraph is table row
                                _bolStartContent = true;
                                RTFDomElement[] es = GetLastElements(true);
                                RTFDomElement lastUnlockElement = null;
                                RTFDomElement lastTableElement = null;
                                for (int iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    RTFDomElement e = es[iCount];
                                    if (e.Locked == false)
                                    {
                                        if (lastUnlockElement == null && !(e is RTFDomParagraph))
                                        {
                                            lastUnlockElement = e;
                                        }
                                        if (e is RTFDomTableRow || e is RTFDomTableCell)
                                        {
                                            lastTableElement = e;
                                            break;
                                        }
                                    }
                                }
                                if (reader.Keyword == RTFConsts.Intbl)
                                {
                                    if (lastTableElement == null)
                                    {
                                        // if can not find unlocked row 
                                        // then new row
                                        RTFDomTableRow row = new RTFDomTableRow();
                                        row.NativeLevel = reader.Level;
                                        lastUnlockElement.AppendChild(row);
                                    }
                                }
                                else if (reader.Keyword == RTFConsts.Trowd)
                                {
                                    // clear row format
                                    RTFDomTableRow row;
                                    if (lastTableElement == null)
                                    {
                                        row = new RTFDomTableRow();
                                        row.NativeLevel = reader.Level;
                                        lastUnlockElement.AppendChild(row);
                                    }
                                    else
                                    {
                                        row = lastTableElement as RTFDomTableRow;
                                        if (row == null)
                                        {
                                            row = (RTFDomTableRow)lastTableElement.Parent;
                                        }
                                    }
                                    row.Attributes.Clear();
                                    row.CellSettings.Clear();
                                    _paragraphFormat.ResetParagraph();
                                }
                                else if (reader.Keyword == RTFConsts.Itap)
                                {
                                    // set nested level
                                    RTFDomTableRow row;

                                    if (reader.Parameter == 0)
                                    {
                                        // is the 0 level , belong to document , not to a table
                                        //foreach (RTFDomElement element in es)
                                        //{
                                        //    if (element is RTFDomTableRow || element is RTFDomTableCell)
                                        //    {
                                        //        element.Locked = true;
                                        //    }
                                        //}
                                    }
                                    else
                                    {
                                        // in a row
                                        if (lastTableElement == null)
                                        {
                                            row = new RTFDomTableRow();
                                            row.NativeLevel = reader.Level;
                                            lastUnlockElement.AppendChild(row);
                                        }
                                        else
                                        {
                                            row = lastTableElement as RTFDomTableRow;
                                            if (row == null)
                                            {
                                                row = (RTFDomTableRow)lastTableElement.Parent;
                                            }
                                            //row.Attributes.Clear();
                                            //row.CellSettings = new ArrayList();
                                        }
                                        if (reader.Parameter == row.Level)
                                        {
                                        }
                                        else if (reader.Parameter > row.Level)
                                        {
                                            // nested row
                                            RTFDomTableRow newRow = new RTFDomTableRow();
                                            newRow.Level = reader.Parameter;
                                            RTFDomTableCell parentCell = GetLastElement<RTFDomTableCell>(false);
                                            if (parentCell == null)
                                                AddContentElement(newRow);
                                            else
                                                parentCell.AppendChild(newRow);
                                        }
                                        else if (reader.Parameter < row.Level)
                                        {
                                            // exit nested row
                                        }
                                    }
                                }//else if
                                break;
                            }
                        case RTFConsts.Nesttableprops:
                            {
                                // ignore
                                break;
                            }
                        case RTFConsts.Row:
                            {
                                // finish read row
                                _bolStartContent = true;
                                RTFDomElement[] es = GetLastElements(true);
                                for (int iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableRow)
                                    {
                                        break;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Nestrow:
                            {
                                // finish nested row
                                _bolStartContent = true;
                                RTFDomElement[] es = GetLastElements(true);
                                for (int iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableRow)
                                    {
                                        break;
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Trrh:
                        case RTFConsts.Trautofit:
                        case RTFConsts.Irowband:
                        case RTFConsts.Trhdr:
                        case RTFConsts.Trkeep:
                        case RTFConsts.Trkeepfollow:
                        case RTFConsts.Trleft:
                        case RTFConsts.Trqc:
                        case RTFConsts.Trql:
                        case RTFConsts.Trqr:
                        case RTFConsts.Trcbpat:
                        case RTFConsts.Trcfpat:
                        case RTFConsts.Trpat:
                        case RTFConsts.Trshdng:
                        case RTFConsts.TrwWidth:
                        case RTFConsts.TrwWidthA:
                        case RTFConsts.Irow:
                        case RTFConsts.Trpaddb:
                        case RTFConsts.Trpaddl:
                        case RTFConsts.Trpaddr:
                        case RTFConsts.Trpaddt:
                        case RTFConsts.Trpaddfb:
                        case RTFConsts.Trpaddfl:
                        case RTFConsts.Trpaddfr:
                        case RTFConsts.Trpaddft:
                        case RTFConsts.Lastrow:
                            {
                                // meet row control word , not parse at first , just save it 
                                _bolStartContent = true;
                                RTFDomTableRow row = GetLastElement<RTFDomTableRow>(false);
                                if (row != null)
                                {
                                    row.Attributes.Add(reader.Keyword, reader.Parameter);
                                }
                                break;
                            }
                        case RTFConsts.Clvmgf:
                        case RTFConsts.Clvmrg:
                        case RTFConsts.Cellx:
                        case RTFConsts.Clvertalt:
                        case RTFConsts.Clvertalc:
                        case RTFConsts.Clvertalb:
                        case RTFConsts.ClNoWrap:
                        case RTFConsts.Clcbpat:
                        case RTFConsts.Clcfpat:
                        case RTFConsts.Clpadl:
                        case RTFConsts.Clpadt:
                        case RTFConsts.Clpadr:
                        case RTFConsts.Clpadb:
                        case RTFConsts.Clbrdrl:
                        case RTFConsts.Clbrdrt:
                        case RTFConsts.Clbrdrr:
                        case RTFConsts.Clbrdrb:
                        case RTFConsts.Brdrtbl:
                        case RTFConsts.Brdrnone:
                            {
                                // meet cell control word , no parse at first , just save it
                                _bolStartContent = true;
                                RTFDomTableRow row = GetLastElement<RTFDomTableRow>(false);
                                //if (row != null && row.Locked == false )
                                {
                                    RTFAttributeList style = null;
                                    if (row.CellSettings.Count > 0)
                                    {
                                        style = (RTFAttributeList)row.CellSettings[row.CellSettings.Count - 1];
                                        if (style.Contains(RTFConsts.Cellx))
                                        {
                                            // if find repeat control word , then can consider this control word
                                            // belong to the next cell . userly cellx is the last control word of 
                                            // a cell , when meet cellx , the current cell defind is finished.
                                            style = new RTFAttributeList();
                                            row.CellSettings.Add(style);
                                        }
                                    }
                                    if (style == null)
                                    {
                                        style = new RTFAttributeList();
                                        row.CellSettings.Add(style);
                                    }
                                    style.Add(reader.Keyword, reader.Parameter);

                                }
                                break;
                            }
                        case RTFConsts.Cell:
                            {
                                // finish cell content
                                _bolStartContent = true;
                                AddContentElement(null);
                                CompleteParagraph();
                                _paragraphFormat.Reset();
                                format.Reset();
                                RTFDomElement[] es = GetLastElements(true);
                                for (int iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    if (es[iCount].Locked == false)
                                    {
                                        es[iCount].Locked = true;
                                        if (es[iCount] is RTFDomTableCell)
                                        {
                                            //((RTFDomTableCell)es[iCount]).Format = format.Clone();
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        case RTFConsts.Nestcell:
                            {
                                // finish nested cell content
                                _bolStartContent = true;
                                AddContentElement(null);
                                CompleteParagraph();
                                RTFDomElement[] es = GetLastElements(false);
                                for (int iCount = es.Length - 1; iCount >= 0; iCount--)
                                {
                                    es[iCount].Locked = true;
                                    if (es[iCount] is RTFDomTableCell)
                                    {
                                        ((RTFDomTableCell)es[iCount]).Format = format;
                                        break;
                                    }
                                }
                                break;
                            }
                        #endregion
                        default:
                            // unsupport keyword
                            if (reader.TokenType == RTFTokenType.ExtKeyword
                                && reader.FirstTokenInGroup)
                            {
                                // if meet unsupport extern keyword , and this token is the first token in 
                                // current group , then ingore whole group.
                                ReadToEndGround(reader);
                                break;
                            }
                            break;
                    }//switch
                }

            }//while
            if (myText.HasContent)
            {
                ApplyText(myText, reader, format);
            }
        }


        /// <summary>
        /// read data , until at the front of the end token belong the current level.
        /// </summary>
        /// <param name="reader"></param>
        private void ReadToEndGround(RTFReader reader)
        {
            reader.ReadToEndGround();
        }


        private void ReadListOverrideTable(RTFReader reader)
        {
            ListOverrideTable = new RTFListOverrideTable();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    RTFListOverride record = null;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            break;
                        }
                        if (reader.CurrentToken.Key == "listoverride")
                        {
                            record = new RTFListOverride();
                            ListOverrideTable.Add(record);
                            continue;
                        }
                        if (record == null)
                        {
                            continue;
                        }
                        switch (reader.CurrentToken.Key)
                        {
                            case "listid":
                                record.ListId = reader.CurrentToken.Param;
                                break;
                            case "listoverridecount":
                                record.ListOverriedCount = reader.CurrentToken.Param;
                                break;
                            case "ls":
                                record.Id = reader.CurrentToken.Param;
                                break;
                        }
                    }
                }
            }//while
        }

        #region HTML RTF

        /// <summary>
        /// HTML content in RTF
        /// </summary>
        public string HtmlContent { get; set; }

        private void ReadHtmlContent(RTFReader reader)
        {
            StringBuilder htmlStr = new StringBuilder();
            bool htmlState = true;
            while (reader.ReadToken() != null)
            {
                if (reader.Keyword == "htmlrtf")
                {
                    if (reader.HasParam && reader.Parameter == 0)
                    {
                        htmlState = false;
                    }
                    else
                    {
                        htmlState = true;
                    }
                }
                else if (reader.Keyword == "htmltag")
                {
                    if (reader.InnerReader.Peek() == ' ')
                    {
                        reader.InnerReader.Read();
                    }
                    string text = ReadInnerText(reader, null, true, false, true);
                    if (string.IsNullOrEmpty(text) == false)
                    {
                        htmlStr.Append(text);
                    }
                }
                else if (reader.TokenType == RTFTokenType.Keyword
                    || reader.TokenType == RTFTokenType.ExtKeyword)
                {
                    if (htmlState == false)
                    {
                        switch (reader.Keyword)
                        {
                            case "par":
                                htmlStr.Append(Environment.NewLine);
                                break;
                            case "line":
                                htmlStr.Append(Environment.NewLine);
                                break;
                            case "tab":
                                htmlStr.Append("\t");
                                break;
                            case "lquote":
                                htmlStr.Append("&lsquo;");
                                break;
                            case "rquote":
                                htmlStr.Append("&rsquo;");
                                break;
                            case "ldblquote":
                                htmlStr.Append("&ldquo;");
                                break;
                            case "rdblquote":
                                htmlStr.Append("&rdquo;");
                                break;
                            case "bullet":
                                htmlStr.Append("&bull;");
                                break;
                            case "endash":
                                htmlStr.Append("&ndash;");
                                break;
                            case "emdash":
                                htmlStr.Append("&mdash;");
                                break;
                            case "~":
                                htmlStr.Append("&nbsp;");
                                break;
                            case "_":
                                htmlStr.Append("&shy;");
                                break;
                        }
                    }
                }
                else if (reader.TokenType == RTFTokenType.Text)
                {
                    if (htmlState == false)
                    {
                        htmlStr.Append(reader.Keyword);
                    }
                }
            }//while
            HtmlContent = htmlStr.ToString();
        }

        #endregion

        private void ReadListTable(RTFReader reader)
        {
            ListTable = new RTFListTable();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    bool firstRead = true;
                    RTFList currentList = null;
                    int level = reader.Level;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            if (reader.Level < level)
                            {
                                break;
                            }
                        }
                        else if (reader.TokenType == RTFTokenType.GroupStart)
                        {
                            // if meet nested level , then ignore
                            //reader.ReadToken();
                            //ReadToEndGround(reader);
                            //reader.ReadToken();
                        }
                        if (firstRead)
                        {
                            if (reader.CurrentToken.Key != "list")
                            {
                                ReadToEndGround(reader);
                                reader.ReadToken();
                                break;
                            }
                            currentList = new RTFList();
                            ListTable.Add(currentList);
                            firstRead = false;
                        }
                        switch (reader.CurrentToken.Key)
                        {
                            case "listtemplateid":
                                currentList.ListTemplateId = reader.CurrentToken.Param;
                                break;
                            case "listid":
                                currentList.ListId = reader.CurrentToken.Param;
                                break;
                            case "listhybrid":
                                currentList.ListHybrid = true;
                                break;
                            case "levelfollow":
                                currentList.LevelFollow = reader.CurrentToken.Param;
                                break;
                            case "levelstartat":
                                currentList.LevelStartAt = reader.CurrentToken.Param;
                                break;
                            case "levelnfc":
                                if (currentList.LevelNfc == LevelNumberType.None)
                                {
                                    currentList.LevelNfc = (LevelNumberType)reader.CurrentToken.Param;
                                }
                                break;
                            case "levelnfcn":
                                if (currentList.LevelNfc == LevelNumberType.None)
                                {
                                    currentList.LevelNfc = (LevelNumberType)reader.CurrentToken.Param;
                                }
                                break;
                            case "leveljc":
                                currentList.LevelJc = reader.CurrentToken.Param;
                                break;
                            case "leveltext":
                                //if (currentList.LevelNfc == LevelNumberType.Bullet)
                                {
                                    if (string.IsNullOrEmpty(currentList.LevelText))
                                    {
                                        string text = ReadInnerText(reader, true);
                                        if (text != null && text.Length > 2)
                                        {
                                            int len = text[0];
                                            len = Math.Min(len, text.Length - 1);
                                            text = text.Substring(1, len);
                                        }
                                        currentList.LevelText = text;
                                    }
                                }
                                break;
                            case "f":
                                currentList.FontName = FontTable.GetFontName(reader.CurrentToken.Param);
                                break;
                        }
                    }//while
                }
            }//while
        }


        /// <summary>
        /// read font table
        /// </summary>
        private void ReadFontTable(RTFReader reader)
        {
            FontTable.Clear();
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                else if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    int index = -1;
                    string name = null;
                    int charset = 1;
                    bool nilFlag = false;
                    while (reader.ReadToken() != null)
                    {
                        if (reader.TokenType == RTFTokenType.GroupEnd)
                        {
                            break;
                        }
                        else if (reader.TokenType == RTFTokenType.GroupStart)
                        {
                            // if meet nested level , then ignore
                            reader.ReadToken();
                            ReadToEndGround(reader);
                            reader.ReadToken();
                        }
                        else if (reader.Keyword == "f" && reader.HasParam)
                        {
                            index = reader.Parameter;
                        }
                        else if (reader.Keyword == "fnil")
                        {
                            name = "Microsoft Sans Serif";
                            nilFlag = true;
                        }
                        else if (reader.Keyword == RTFConsts.Fcharset)
                        {
                            charset = reader.Parameter;
                        }
                        else if (reader.CurrentToken.IsTextToken)
                        {
                            //if (defaultFont == false)
                            {
                                name = ReadInnerText(
                                    reader,
                                    reader.CurrentToken,
                                    false,
                                    false,
                                    false);
                                if (name != null)
                                {
                                    name = name.Trim();

                                    if (name.EndsWith(";"))
                                    {
                                        name = name.Substring(0, name.Length - 1);
                                    }
                                }
                            }
                        }
                    }
                    if (index >= 0 && name != null)
                    {
                        if (name.EndsWith(";"))
                        {
                            name = name.Substring(0, name.Length - 1);
                        }
                        name = name.Trim();
                        if (string.IsNullOrEmpty(name))
                        {
                            name = "Microsoft Sans Serif";
                        }
                        //System.Console.WriteLine( "Index:" + index + "  Name:" + name );
                        RTFFont font = new RTFFont(index, name);
                        font.Charset = charset;
                        font.NilFlag = nilFlag;
                        FontTable.Add(font);
                    }
                }//else
            }//while
        }

        /// <summary>
        /// read color table
        /// </summary>
        private void ReadColorTable(RTFReader reader)
        {
            ColorTable.Clear();
            ColorTable.CheckValueExistWhenAdd = false;
            int r = -1;
            int g = -1;
            int b = -1;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                switch (reader.Keyword)
                {
                    case "red":
                        r = reader.Parameter;
                        break;
                    case "green":
                        g = reader.Parameter;
                        break;
                    case "blue":
                        b = reader.Parameter;
                        break;
                    case ";":
                        if (r >= 0 && g >= 0 && b >= 0)
                        {
                            Color c = Color.FromArgb(255, r, g, b);
                            ColorTable.Add(c);
                            r = -1;
                            g = -1;
                            b = -1;
                        }
                        break;
                }
            }
            if (r >= 0 && g >= 0 && b >= 0)
            {
                // read the last color
                Color c = Color.FromArgb(255, r, g, b);
                ColorTable.Add(c);
            }
        }

        /// <summary>
        /// read document information
        /// </summary>
        private void ReadDocumentInfo(RTFReader reader)
        {
            _myInfo.Clear();
            int level = 0;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    level++;
                }
                else if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                else
                {
                    switch (reader.Keyword)
                    {
                        case "creatim":
                            _myInfo.Creatim = ReadDateTime(reader);
                            level--;
                            break;
                        case "revtim":
                            _myInfo.Revtim = ReadDateTime(reader);
                            level--;
                            break;
                        case "printim":
                            _myInfo.Printim = ReadDateTime(reader);
                            level--;
                            break;
                        case "buptim":
                            _myInfo.Buptim = ReadDateTime(reader);
                            level--;
                            break;
                        default:
                            if (reader.Keyword != null)
                            {
                                if (reader.HasParam)
                                {
                                    _myInfo.SetInfo(reader.Keyword, reader.Parameter.ToString());
                                }
                                else
                                {
                                    _myInfo.SetInfo(reader.Keyword, ReadInnerText(reader, true));
                                }
                            }
                            break;
                    }
                }
            }//while
        }

        /// <summary>
        /// read datetime
        /// </summary>
        /// <param name="reader">reader</param>
        /// <returns>datetime value</returns>
        private DateTime ReadDateTime(RTFReader reader)
        {
            int yr = 1900;
            int mo = 1;
            int dy = 1;
            int hr = 0;
            int min = 0;
            int sec = 0;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    break;
                }
                switch (reader.Keyword)
                {
                    case "yr":
                        yr = reader.Parameter;
                        break;
                    case "mo":
                        mo = reader.Parameter;
                        break;
                    case "dy":
                        dy = reader.Parameter;
                        break;
                    case "hr":
                        hr = reader.Parameter;
                        break;
                    case "min":
                        min = reader.Parameter;
                        break;
                    case "sec":
                        sec = reader.Parameter;
                        break;
                }//switch
            }//while

            if (mo == 0 || dy == 0)
            {
                return DateTime.MinValue;
            }
            return new DateTime(yr, mo, dy, hr, min, sec);
        }

        //private RTFDomImage ReadDomImage(RTFReader reader, DocumentFormatInfo format)
        //{
        //    bolStartContent = true;
        //    RTFDomImage img = new RTFDomImage();
        //    img.NativeLevel = reader.Level;
        //    AddContentElement(img);
        //    RTFTextContainer txt = new RTFTextContainer( this );
        //    while (reader.ReadToken() != null)
        //    {
        //        if (reader.Level < img.NativeLevel)
        //        {
        //            break;
        //        }
        //        if (reader.TokenType == RTFTokenType.GroupStart)
        //        {
        //            continue;
        //        }
        //        if (reader.TokenType == RTFTokenType.GroupEnd)
        //        {
        //            continue;
        //        }
        //        if (reader.TokenType == RTFTokenType.Text)
        //        {
        //            txt.Accept(reader.CurrentToken, reader);
        //            continue;
        //        }
        //        switch (reader.Keyword)
        //        {
        //            case RTFConsts._nonshppict :
        //                // ignore group
        //                ReadToEndGround(reader);
        //                break;
        //            case RTFConsts._picscalex:
        //                img.ScaleX = reader.Parameter;
        //                break;
        //            case RTFConsts._picscaley:
        //                img.ScaleY = reader.Parameter;
        //                break;
        //            case RTFConsts._picwgoal:
        //                img.DesiredWidth = reader.Parameter;
        //                break;
        //            case RTFConsts._pichgoal:
        //                img.DesiredHeight = reader.Parameter;
        //                break;
        //            case RTFConsts._blipuid:
        //                img.ID = ReadInnerText(reader, true);
        //                break;
        //            case RTFConsts._emfblip:
        //                img.PicType = RTFPicType.Emfblip;
        //                break;
        //            case RTFConsts._pngblip:
        //                img.PicType = RTFPicType.Pngblip;
        //                break;
        //            case RTFConsts._jpegblip:
        //                img.PicType = RTFPicType.Jpegblip;
        //                break;
        //            case RTFConsts._macpict:
        //                img.PicType = RTFPicType.Macpict;
        //                break;
        //            case RTFConsts._pmmetafile:
        //                img.PicType = RTFPicType.Pmmetafile;
        //                break;
        //            case RTFConsts._wmetafile:
        //                img.PicType = RTFPicType.Wmetafile;
        //                break;
        //            case RTFConsts._dibitmap:
        //                img.PicType = RTFPicType.Dibitmap;
        //                break;
        //            case RTFConsts._wbitmap:
        //                img.PicType = RTFPicType.Wbitmap;
        //                break;
        //        }//switch
        //    }//while
        //    if (txt.HasContent)
        //    {
        //        img.Data = HexToBytes(txt.Text);
        //    }
        //    img.Format = format.Clone();
        //    img.Width = (int)(img.DesiredWidth * img.ScaleX / 100);
        //    img.Height = (int)(img.DesiredHeight * img.ScaleY / 100);
        //    img.Locked = true;
        //    return img;
        //}

        /// <summary>
        /// Read a rtf emb object
        /// </summary>
        /// <param name="reader">reader</param>
        /// <param name="format">format</param>
        /// <returns>rtf emb object instance</returns>
        private RTFDomObject ReadDomObject(RTFReader reader, DocumentFormatInfo format)
        {
            RTFDomObject obj = new RTFDomObject();
            obj.NativeLevel = reader.Level;
            AddContentElement(obj);
            int levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.Level < levelBack)
                {
                    break;
                }
                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                    continue;
                }
                if (reader.TokenType == RTFTokenType.GroupEnd)
                {
                    continue;
                }
                if (reader.Level == obj.NativeLevel + 1
                    && reader.Keyword.StartsWith("attribute_"))
                {
                    obj.CustomAttributes[reader.Keyword] = ReadInnerText(reader, true);
                }
                switch (reader.Keyword)
                {
                    case RTFConsts.Objautlink:
                        obj.Type = RTFObjectType.AutLink;
                        break;
                    case RTFConsts.Objclass:
                        obj.ClassName = ReadInnerText(reader, true);
                        break;
                    case RTFConsts.Objdata:
                        string data = ReadInnerText(reader, true);
                        obj.Content = HexToBytes(data);
                        break;
                    case RTFConsts.Objemb:
                        obj.Type = RTFObjectType.Emb;
                        break;
                    case RTFConsts.Objh:
                        obj.Height = reader.Parameter;
                        break;
                    case RTFConsts.Objhtml:
                        obj.Type = RTFObjectType.Html;
                        break;
                    case RTFConsts.Objicemb:
                        obj.Type = RTFObjectType.Icemb;
                        break;
                    case RTFConsts.Objlink:
                        obj.Type = RTFObjectType.Link;
                        break;
                    case RTFConsts.Objname:
                        obj.Name = ReadInnerText(reader, true);
                        break;
                    case RTFConsts.Objocx:
                        obj.Type = RTFObjectType.Ocx;
                        break;
                    case RTFConsts.Objpub:
                        obj.Type = RTFObjectType.Pub;
                        break;
                    case RTFConsts.Objsub:
                        obj.Type = RTFObjectType.Sub;
                        break;
                    case RTFConsts.Objtime:
                        break;
                    case RTFConsts.Objw:
                        obj.Width = reader.Parameter;
                        break;
                    case RTFConsts.Objscalex:
                        obj.ScaleX = reader.Parameter;
                        break;
                    case RTFConsts.Objscaley:
                        obj.ScaleY = reader.Parameter;
                        break;
                    case RTFConsts.Result:
                        RTFDomElementContainer result = new RTFDomElementContainer();
                        result.Name = RTFConsts.Result;
                        obj.AppendChild(result);
                        Load(reader, format);
                        result.Locked = true;
                        break;
                }
            }//while
            obj.Locked = true;
            return obj;
        }

        /// <summary>
        /// read field
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        private RTFDomField ReadDomField(
            RTFReader reader,
            DocumentFormatInfo format)
        {
            RTFDomField field = new RTFDomField();
            field.NativeLevel = reader.Level;
            AddContentElement(field);
            int levelBack = reader.Level;
            while (reader.ReadToken() != null)
            {
                if (reader.Level < levelBack)
                {
                    break;
                }

                if (reader.TokenType == RTFTokenType.GroupStart)
                {
                }
                else if (reader.TokenType == RTFTokenType.GroupEnd)
                {

                }
                else
                {
                    switch (reader.Keyword)
                    {
                        case RTFConsts.Flddirty:
                            {
                                field.Method = RTFDomFieldMethod.Dirty;
                                break;
                            }
                        case RTFConsts.Fldedit:
                            {
                                field.Method = RTFDomFieldMethod.Edit;
                                break;
                            }
                        case RTFConsts.Fldlock:
                            {
                                field.Method = RTFDomFieldMethod.Lock;
                                break;
                            }
                        case RTFConsts.Fldpriv:
                            {
                                field.Method = RTFDomFieldMethod.Priv;
                                break;
                            }
                        case RTFConsts.Fldrslt:
                            {
                                RTFDomElementContainer result = new RTFDomElementContainer();
                                result.Name = RTFConsts.Fldrslt;
                                field.AppendChild(result);
                                Load(reader, format);
                                result.Locked = true;
                                //field.Result = ReadInnerText(reader, true);
                                break;
                            }
                        case RTFConsts.Fldinst:
                            {
                                RTFDomElementContainer inst = new RTFDomElementContainer();
                                inst.Name = RTFConsts.Fldinst;
                                field.AppendChild(inst);
                                Load(reader, format);
                                inst.Locked = true;
                                string txt = inst.InnerText;
                                if (txt != null)
                                {
                                    int index = txt.IndexOf(RTFConsts.Hyperlink);
                                    if (index >= 0)
                                    {
                                        string link;
                                        int index1 = txt.IndexOf('\"', index);
                                        if (index1 > 0 && txt.Length > index1 + 2)
                                        {
                                            int index2 = txt.IndexOf('\"', index1 + 2);
                                            if (index2 > index1)
                                            {
                                                link = txt.Substring(index1 + 1, index2 - index1 - 1);
                                                if (format.Parent != null)
                                                {
                                                    if (link.StartsWith("_Toc"))
                                                        link = "#" + link;
                                                    format.Parent.Link = link;
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }

                                break;
                            }
                    }
                }
            }
            field.Locked = true;
            return field;

        }//private RTFDomField ReadDomField(RTFReader reader)


        private string ReadInnerText(RTFReader reader, bool deeply)
        {
            return ReadInnerText(
                reader,
                null,
                deeply,
                false,
                false);
        }

        /// <summary>
        /// read the following plain text in the current level
        /// </summary>
        /// <param name="reader">RTF reader</param>
        /// <param name="firstToken"></param>
        /// <param name="deeply">whether read the text in the sub level</param>
        /// <param name="breakMeetControlWord"></param>
        /// <param name="htmlMode"></param>
        /// <returns>text</returns>
        private string ReadInnerText(
            RTFReader reader,
            RTFToken firstToken,
            bool deeply,
            bool breakMeetControlWord,
            bool htmlMode)
        {
            int level = 0;
            RTFTextContainer container = new RTFTextContainer(this);
            container.Accept(firstToken, reader);
            while (true)
            {
                RTFTokenType type = reader.PeekTokenType();
                if (type == RTFTokenType.Eof)
                    break;
                if (type == RTFTokenType.GroupStart)
                {
                    level++;
                }
                else if (type == RTFTokenType.GroupEnd)
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                reader.ReadToken();

                if (deeply || level == 0)
                {
                    if (htmlMode)
                    {
                        if (reader.Keyword == "par")
                        {
                            container.Append(Environment.NewLine);
                            continue;
                        }
                    }
                    if (!container.Accept(reader.CurrentToken, reader) && breakMeetControlWord)
                    {
                        break;
                    }
                }
            }//while
            return container.Text;
        }

        public override string ToDomString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(ToString());
            builder.Append(Environment.NewLine + "   Info");
            foreach (string item in _myInfo.StringItems)
            {
                builder.Append(Environment.NewLine + "      " + item);
            }
            builder.Append(Environment.NewLine + "   ColorTable(" + ColorTable.Count + ")");
            for (int iCount = 0; iCount < ColorTable.Count; iCount++)
            {
                Color c = ColorTable[iCount];
                builder.Append(Environment.NewLine + "      " + iCount + ":" + c.R + " " + c.G + " " + c.B);
            }
            builder.Append(Environment.NewLine + "   FontTable(" + FontTable.Count + ")");
            foreach (RTFFont font in FontTable)
            {
                builder.Append(Environment.NewLine + "      " + font);
            }
            if (ListTable.Count > 0)
            {
                builder.Append(Environment.NewLine + "   ListTable(" + ListTable.Count + ")");
                foreach (RTFList list in ListTable)
                {
                    builder.Append(Environment.NewLine + "      " + list);
                }
            }
            if (ListOverrideTable.Count > 0)
            {
                builder.Append(Environment.NewLine + "   ListOverrideTable(" + ListOverrideTable.Count + ")");
                foreach (RTFListOverride list in ListOverrideTable)
                {
                    builder.Append(Environment.NewLine + "      " + list);
                }
            }
            builder.Append(Environment.NewLine + "   -----------------------");
            if (string.IsNullOrEmpty(HtmlContent) == false)
            {
                builder.Append(Environment.NewLine + "   HTMLContent:" + HtmlContent);
                builder.Append(Environment.NewLine + "   -----------------------");
            }
            ToDomString(Elements, builder, 1);
            return builder.ToString();
        }
    }
}
