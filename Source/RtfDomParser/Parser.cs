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
    /// RTF plain text container
    /// </summary>
    internal class RTFTextContainer
    {
        private readonly ByteBuffer _myBuffer = new ByteBuffer();

        private StringBuilder _myStr = new StringBuilder();

        /// <summary>
        /// initialize instance
        /// </summary>
        /// <param name="doc">owner document</param>
        public RTFTextContainer(RTFDomDocument doc)
        {
            Level = 0;
            Document = doc;
        }

        /// <summary>
        /// Owner document
        /// </summary>
        public RTFDomDocument Document { get; set; }

        /// <summary>
        /// this container has some text
        /// </summary>
        public bool HasContent
        {
            get
            {
                CheckBuffer();
                return _myStr.Length > 0;
            }
        }

        /// <summary>
        /// text value
        /// </summary>
        public string Text
        {
            get
            {
                CheckBuffer();
                return _myStr.ToString();
            }
        }

        public int Level { get; set; }

        /// <summary>
        /// Append text content
        /// </summary>
        /// <param name="text"></param>
        public void Append(string text)
        {
            if (string.IsNullOrEmpty(text) == false)
            {
                CheckBuffer();
                _myStr.Append(text);
            }
        }

        /// <summary>
        /// Accept rtf token
        /// </summary>
        /// <param name="token">RTF token</param>
        /// <param name="reader"></param>
        /// <returns>Is accept it?</returns>
        public bool Accept(RTFToken token, RTFReader reader)
        {
            if (token == null)
            {
                return false;
            }
            if (token.Type == RTFTokenType.Text)
            {
                if (reader != null)
                {
                    if (token.Key[0] == '?')
                    {
                        if (reader.LastToken != null)
                        {
                            if (reader.LastToken.Type == RTFTokenType.Keyword
                                && reader.LastToken.Key == "u"
                                && reader.LastToken.HasParam)
                            {
                                if (token.Key.Length > 0)
                                {
                                    CheckBuffer();
                                    //myStr.Append(token.Key.Substring(1));
                                }
                                return true;
                            }
                        }
                    }
                    //else if (token.Key == "\"")
                    //{
                    //    CheckBuffer();
                    //    while (true)
                    //    {
                    //        int v = reader.InnerReader.Read();
                    //        if (v > 0)
                    //        {
                    //            if (v == (int)'"')
                    //            {
                    //                break;
                    //            }
                    //            else
                    //            {
                    //                myStr.Append((char)v);
                    //            }
                    //        }
                    //        else
                    //        {
                    //            break;
                    //        }
                    //    }//while
                    //    return true;
                    //}
                }
                CheckBuffer();
                _myStr.Append(token.Key);
                return true;
            }
            if (token.Type == RTFTokenType.Control
                && token.Key == "'" && token.HasParam)
            {
                if (reader.CurrentLayerInfo.CheckUcValueCount())
                {
                    _myBuffer.Add((byte) token.Param);
                }
                return true;
            }
            if (token.Key == RTFConsts.U && token.HasParam)
            {
                // Unicode char
                CheckBuffer();
                _myStr.Append((char) token.Param);
                reader.CurrentLayerInfo.UcValueCount = reader.CurrentLayerInfo.UcValue;
                return true;
            }
            if (token.Key == "tab")
            {
                CheckBuffer();
                _myStr.Append("\t");
                return true;
            }
            if (token.Key == "emdash")
            {
                CheckBuffer();
                _myStr.Append('—');
                return true;
            }
            if (token.Key == "")
            {
                CheckBuffer();
                _myStr.Append('-');
                return true;
            }
            return false;
        }

        /// <summary>
        /// clear value
        /// </summary>
        public void Clear()
        {
            _myBuffer.Clear();
            _myStr = new StringBuilder();
        }

        private void CheckBuffer()
        {
            if (_myBuffer.Count > 0)
            {
                var txt = _myBuffer.GetString(Document.RuntimeEncoding);
                _myStr.Append(txt);
                _myBuffer.Clear();
            }
        }
    }

    public class RTFReader : IDisposable
    {
        //private RTFToken myLastToken = null;
        //public RTFToken LastToken
        //{
        //    get
        //    {
        //        return myLastToken;
        //    }
        //}

        private bool _enableDefaultProcess = true;

        private readonly Stack<RTFRawLayerInfo> _layerStack = new Stack<RTFRawLayerInfo>();

        //private System.Collections.ArrayList myTokenStack = new System.Collections.ArrayList();
        //public System.Collections.ArrayList TokenStack
        //{
        //    get
        //    {
        //        return myTokenStack ;
        //    }
        //}

        private Stream _myBaseStream;

        private RTFLex _myLex;
        //private RTFToken myToken = null ;

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFReader()
        {
        }


        public RTFReader(Stream stream)
        {
            var reader = new StreamReader(stream, Encoding.GetEncoding("us-ascii"));
            LoadReader(reader);
            _myBaseStream = stream;
        }

        public RTFReader(TextReader reader)
        {
            LoadReader(reader);
        }

        public TextReader InnerReader { get; private set; }

        /// <summary>
        /// current token
        /// </summary>
        public RTFToken CurrentToken { get; private set; }

        /// <summary>
        /// current token's type
        /// </summary>
        public RTFTokenType TokenType
        {
            get
            {
                if (CurrentToken == null)
                    return RTFTokenType.None;
                return CurrentToken.Type;
            }
        }

        /// <summary>
        /// current keyword
        /// </summary>
        public string Keyword
        {
            get
            {
                if (CurrentToken == null)
                    return null;
                return CurrentToken.Key;
            }
        }

        /// <summary>
        /// is current token has a parameter
        /// </summary>
        public bool HasParam
        {
            get
            {
                if (CurrentToken == null)
                    return false;
                return CurrentToken.HasParam;
            }
        }

        /// <summary>
        /// current parameter
        /// </summary>
        public int Parameter
        {
            get
            {
                if (CurrentToken == null)
                    return 0;
                return CurrentToken.Param;
            }
        }

        public int ContentPosition
        {
            get
            {
                if (_myBaseStream == null)
                    return 0;
                return (int) _myBaseStream.Position;
            }
        }

        public int ContentLength
        {
            get
            {
                if (_myBaseStream == null)
                    return 0;
                return (int) _myBaseStream.Length;
            }
        }

        /// <summary>
        /// Current token is the first token in owner group
        /// </summary>
        public bool FirstTokenInGroup { get; private set; }

        /// <summary>
        /// lost token
        /// </summary>
        public RTFToken LastToken { get; private set; }

        public int Level { get; private set; }

        /// <summary>
        /// total of this object handle tokens
        /// </summary>
        public int TokenCount { get; set; }

        public bool EnableDefaultProcess
        {
            get { return _enableDefaultProcess; }
            set { _enableDefaultProcess = value; }
        }

        public RTFRawLayerInfo CurrentLayerInfo
        {
            get
            {
                if (_layerStack.Count == 0)
                {
                    _layerStack.Push(new RTFRawLayerInfo());
                }
                return _layerStack.Peek();
            }
        }

        public void Dispose()
        {
            Close();
        }

        /// <summary>
        /// load rtf document
        /// </summary>
        /// <returns>is operation successful</returns>
        public void LoadStream(Stream stream)
        {
            //myTokenStack.Clear();
            CurrentToken = null;
            //var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(strFileName);
            //var randomAccessStream = await file.OpenReadAsync();

            //var stream = randomAccessStream.AsStreamForRead();

            //System.IO.FileStream stream = new System.IO.FileStream( strFileName , System.IO.FileMode.Open , System.IO.FileAccess.Read );
            InnerReader = new StreamReader(stream, Encoding.GetEncoding("us-ascii"));
            _myBaseStream = stream;
            _myLex = new RTFLex(InnerReader);
        }

        /// <summary>
        /// load rtf document
        /// </summary>
        /// <param name="reader">text reader</param>
        /// <returns>is operation successful</returns>
        public bool LoadReader(TextReader reader)
        {
            //myTokenStack.Clear();
            CurrentToken = null;
            if (reader != null)
            {
                InnerReader = reader;
                _myLex = new RTFLex(InnerReader);
                return true;
            }
            return false;
        }

        /// <summary>
        /// load rtf text
        /// </summary>
        /// <param name="strText">RTF text</param>
        /// <returns>is operation successful</returns>
        public bool LoadRTFText(string strText)
        {
            //myTokenStack.Clear();
            CurrentToken = null;
            if (strText != null && strText.Length > 3)
            {
                InnerReader = new StringReader(strText);
                _myLex = new RTFLex(InnerReader);
                return true;
            }
            return false;
        }

        public void Close()
        {
            if (InnerReader != null)
            {
                InnerReader.Dispose();
                InnerReader = null;
            }
        }

        /// <summary>
        /// next token type
        /// </summary>
        /// <returns></returns>
        public RTFTokenType PeekTokenType()
        {
            return _myLex.PeekTokenType();
        }

        public void DefaultProcess()
        {
            if (CurrentToken != null)
            {
                switch (CurrentToken.Key)
                {
                    case "uc":
                        CurrentLayerInfo.UcValue = Parameter;
                        break;
                }
            }
        }

        /// <summary>
        /// read token
        /// </summary>
        /// <returns>token readed</returns>
        public RTFToken ReadToken()
        {
            FirstTokenInGroup = false;
            LastToken = CurrentToken;
            if (LastToken != null && LastToken.Type == RTFTokenType.GroupStart)
            {
                FirstTokenInGroup = true;
            }
            CurrentToken = _myLex.NextToken();
            if (CurrentToken == null || CurrentToken.Type == RTFTokenType.Eof)
            {
                CurrentToken = null;
                return null;
            }
            TokenCount++;
            if (CurrentToken.Type == RTFTokenType.GroupStart)
            {
                if (_layerStack.Count == 0)
                {
                    _layerStack.Push(new RTFRawLayerInfo());
                }
                else
                {
                    var info = _layerStack.Peek();
                    _layerStack.Push(info.Clone());
                }
                Level++;
            }
            else if (CurrentToken.Type == RTFTokenType.GroupEnd)
            {
                if (_layerStack.Count > 0)
                {
                    _layerStack.Pop();
                }
                Level--;
            }
            if (EnableDefaultProcess)
            {
                DefaultProcess();
            }
            //if (myTokenStack.Count > 0)
            //{
            //    myCurrentToken.Parent = (RTFToken)myTokenStack[myTokenStack.Count - 1];
            //}
            //myTokenStack.Add( myCurrentToken );
            return CurrentToken;
        }

        /// <summary>
        /// read and ignore data , until just the end of current group,preserve the end.
        /// </summary>
        public void ReadToEndGround()
        {
            var level = 0;
            while (true)
            {
                var c = InnerReader.Peek();
                if (c == -1)
                {
                    break;
                }
                else if (c == '{')
                {
                    level++;
                }
                else if (c == '}')
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                c = InnerReader.Read();
            }
        }

        public override string ToString()
        {
            return "RTFReader Level:" + Level + " " + Keyword;
        }
    }

    public class RTFLex
    {
        private const int Eof = -1;

        private readonly TextReader _myReader;

        /// <summary>
        /// Initialize instance
        /// </summary>
        /// <param name="reader">reader</param>
        public RTFLex(TextReader reader)
        {
            _myReader = reader;
        }

        public RTFTokenType PeekTokenType()
        {
            var c = _myReader.Peek();

            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
            {
                _myReader.Read();
                c = _myReader.Peek();
            }
            if (c == Eof)
            {
                return RTFTokenType.Eof;
            }
            switch (c)
            {
                case '{':
                    return RTFTokenType.GroupStart;
                case '}':
                    return RTFTokenType.GroupEnd;
                case '\\':
                    return RTFTokenType.Control;
                default:
                    return RTFTokenType.Text;
            }
        }


        /// <summary>
        /// read next token
        /// </summary>
        /// <returns>token</returns>
        public RTFToken NextToken()
        {
            int c;
            var token = new RTFToken();

            //myReader.Read();

            c = _myReader.Read();
            if (c == '\"')
            {
                var str = new StringBuilder();
                while (true)
                {
                    c = _myReader.Read();
                    if (c < 0)
                    {
                        break;
                    }
                    if (c == '\"')
                    {
                        break;
                    }
                    str.Append((char) c);
                } //while
                token.Type = RTFTokenType.Text;
                token.Key = str.ToString();
                return token;
            }

            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
            {
                c = _myReader.Read();
                //c = myReader.Peek();
            }

            //
//			c = myReader.Read();
//			while( c == '\r'
//				|| c == '\n' 
//				|| c == '\t'
//				|| c == '\0')
//			{
//				c = myReader.Read();
//			}
            if (c != Eof)
            {
                switch (c)
                {
                    case '{':
                        token.Type = RTFTokenType.GroupStart;
                        break;
                    case '}':
                        token.Type = RTFTokenType.GroupEnd;
                        break;
                    case '\\':
                        ParseKeyword(token);
                        break;
                    default:
                        token.Type = RTFTokenType.Text;
                        ParseText(c, token);
                        break;
                }
            }
            else
            {
                token.Type = RTFTokenType.Eof;
            }
            return token;
        }

        private void ParseKeyword(RTFToken token)
        {
            int c;
            var ext = false;
            c = _myReader.Peek();
            if (char.IsLetter((char) c) == false)
            {
                _myReader.Read();
                if (c == '*')
                {
                    // expend keyword
                    token.Type = RTFTokenType.Keyword;
                    _myReader.Read();
                    //myReader.Read();
                    ext = true;
                    goto ReadKeywrod;
                }
                if (c == '\\' || c == '{' || c == '}')
                {
                    // special character
                    token.Type = RTFTokenType.Text;
                    token.Key = ((char) c).ToString();
                }
                else
                {
                    token.Type = RTFTokenType.Control;
                    token.Key = ((char) c).ToString();
                    if (token.Key == "\'")
                    {
                        // read 2 hex characters
                        var text = new StringBuilder();
                        text.Append((char) _myReader.Read());
                        text.Append((char) _myReader.Read());
                        token.HasParam = true;
                        token.Param = Convert.ToInt32(text.ToString().ToLower(), 16);
                        if (token.Param == 0)
                        {
                            token.Param = 0;
                        }
                    }
                }
                return;
            } //if( char.IsLetter( ( char ) c ) == false )

            ReadKeywrod :

            // read keyword
            var keyword = new StringBuilder();
            c = _myReader.Peek();
            while (char.IsLetter((char) c))
            {
                _myReader.Read();
                keyword.Append((char) c);
                c = _myReader.Peek();
            }

            if (ext)
                token.Type = RTFTokenType.ExtKeyword;
            else
                token.Type = RTFTokenType.Keyword;
            token.Key = keyword.ToString();

            // read a interger
            if (char.IsDigit((char) c) || c == '-')
            {
                token.HasParam = true;
                var negative = false;
                if (c == '-')
                {
                    negative = true;
                    _myReader.Read();
                }
                c = _myReader.Peek();
                var text = new StringBuilder();
                while (char.IsDigit((char) c))
                {
                    _myReader.Read();
                    text.Append((char) c);
                    c = _myReader.Peek();
                }
                var p = Convert.ToInt32(text.ToString());
                if (negative)
                    p = -p;
                token.Param = p;
            } //if( char.IsDigit( ( char ) c ) || c == '-' )

            if (c == ' ')
            {
                _myReader.Read();
            }
        }

        private void ParseText(int c, RTFToken token)
        {
            var myStr = new StringBuilder(((char) c).ToString());

            c = ClearWhiteSpace();

            while (c != '\\' && c != '}' && c != '{' && c != Eof)
            {
                _myReader.Read();
                myStr.Append((char) c);
                c = ClearWhiteSpace();
            }

            token.Key = myStr.ToString();
        }

        private int ClearWhiteSpace()
        {
            var c = _myReader.Peek();
            while (c == '\r'
                   || c == '\n'
                   || c == '\t'
                   || c == '\0')
            {
                _myReader.Read();
                c = _myReader.Peek();
            }
            return c;
        }
    }

    /// <summary>
    /// rtf token type
    /// </summary>
    public class RTFToken
    {
        private RTFTokenType _intType = RTFTokenType.None;

        public RTFToken()
        {
            Parent = null;
            Param = 0;
            HasParam = false;
            Key = null;
        }

        /// <summary>
        /// type
        /// </summary>
        public RTFTokenType Type
        {
            get { return _intType; }
            set { _intType = value; }
        }

        /// <summary>
        /// keyword
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public bool HasParam { get; set; }

        public int Param { get; set; }

        /// <summary>
        /// parent token
        /// </summary>
        public RTFToken Parent { get; set; }

        public bool IsTextToken
        {
            get
            {
                if (_intType == RTFTokenType.Text)
                    return true;
                if (_intType == RTFTokenType.Control && Key == "'" && HasParam)
                    return true;
                return false;
            }
        }

        public override string ToString()
        {
            if (_intType == RTFTokenType.Keyword)
            {
                return Key + Param;
            }
            if (_intType == RTFTokenType.GroupStart)
            {
                return "{";
            }
            if (_intType == RTFTokenType.GroupEnd)
            {
                return "}";
            }
            if (_intType == RTFTokenType.Text)
            {
                return "Text:" + Param;
            }
            return _intType + ":" + Key + " " + Param;
        }
    }

    /// <summary>
    /// rtf token type
    /// </summary>
    public enum RTFTokenType
    {
        None,
        Keyword,
        ExtKeyword,
        Control,
        Text,
        Eof,
        GroupStart,
        GroupEnd
    }
}