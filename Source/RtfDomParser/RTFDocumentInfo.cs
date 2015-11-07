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
using System.Linq;

namespace RtfDomParser
{
    public static class DictionaryExtension
    {
        public static TValue ValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key)
        {
            TValue value;
            dictionary.TryGetValue(key, out value);
            return value;
        }
    }

    /// <summary>
    /// Document information
    /// </summary>
    public class RTFDocumentInfo
    {
        private DateTime _dtmBuptim = DateTime.Now;

        private DateTime _dtmCreatim = DateTime.Now;

        private DateTime _dtmPrintim = DateTime.Now;

        private DateTime _dtmRevtim = DateTime.Now;

        private readonly Dictionary<string, string> _myInfo =
            new Dictionary<string, string>();

        /// <summary>
        /// document title
        /// </summary>
        public string Title
        {
            get { return _myInfo.ValueOrDefault("title"); }
            set { _myInfo["title"] = value; }
        }

        public string Subject
        {
            get { return _myInfo.ValueOrDefault("subject"); }
            set { _myInfo["subject"] = value; }
        }

        public string Author
        {
            get { return _myInfo.ValueOrDefault("author"); }
            set { _myInfo["author"] = value; }
        }

        public string Manager
        {
            get { return _myInfo.ValueOrDefault("manager"); }
            set { _myInfo["manager"] = value; }
        }

        public string Company
        {
            get { return _myInfo.ValueOrDefault("company"); }
            set { _myInfo["company"] = value; }
        }

        public string Operator
        {
            get { return _myInfo.ValueOrDefault("operator"); }
            set { _myInfo["operator"] = value; }
        }

        public string Category
        {
            get { return _myInfo.ValueOrDefault("category"); }
            set { _myInfo["categroy"] = value; }
        }

        public string Keywords
        {
            get { return _myInfo.ValueOrDefault("keywords"); }
            set { _myInfo["keywords"] = value; }
        }

        public string Comment
        {
            get { return _myInfo.ValueOrDefault("comment"); }
            set { _myInfo["comment"] = value; }
        }

        public string Doccomm
        {
            get { return _myInfo.ValueOrDefault("doccomm"); }
            set { _myInfo["doccomm"] = value; }
        }

        public string HLinkbase
        {
            get { return _myInfo.ValueOrDefault("hlinkbase"); }
            set { _myInfo["hlinkbase"] = value; }
        }

        /// <summary>
        /// total edit minutes
        /// </summary>
        public int Edmins
        {
            get
            {
                if (_myInfo.ContainsKey("edmins"))
                {
                    var v = Convert.ToString(_myInfo["edmins"]);
                    int result;
                    if (int.TryParse(v, out result))
                    {
                        return result;
                    }
                }
                return 0;
            }
            set { _myInfo["edmins"] = value.ToString(); }
        }

        /// <summary>
        /// version
        /// </summary>
        public string Vern
        {
            get { return _myInfo["vern"]; }
            set { _myInfo["vern"] = value; }
        }

        /// <summary>
        /// number of pages
        /// </summary>
        public string Nofpages
        {
            get { return _myInfo["nofpages"]; }
            set { _myInfo["nofpages"] = value; }
        }

        /// <summary>
        /// number of words
        /// </summary>
        public string Nofwords
        {
            get { return _myInfo["nofwords"]; }
            set { _myInfo["nofwords"] = value; }
        }

        /// <summary>
        /// number of characters , include whitespace
        /// </summary>
        public string Nofchars
        {
            get { return _myInfo["nofchars"]; }
            set { _myInfo["nofchars"] = value; }
        }

        /// <summary>
        /// number of characters , exclude white space
        /// </summary>
        public string Nofcharsws
        {
            get { return _myInfo["nofcharsws"]; }
            set { _myInfo["nofcharsws"] = value; }
        }

        /// <summary>
        /// inner id
        /// </summary>
        public string Id
        {
            get { return _myInfo["id"]; }
            set { _myInfo["id"] = value; }
        }

        /// <summary>
        /// creation time
        /// </summary>
        public DateTime Creatim
        {
            get { return _dtmCreatim; }
            set { _dtmCreatim = value; }
        }

        /// <summary>
        /// modified time
        /// </summary>
        public DateTime Revtim
        {
            get { return _dtmRevtim; }
            set { _dtmRevtim = value; }
        }

        /// <summary>
        /// last print time
        /// </summary>
        public DateTime Printim
        {
            get { return _dtmPrintim; }
            set { _dtmPrintim = value; }
        }

        /// <summary>
        /// back up time
        /// </summary>
        public DateTime Buptim
        {
            get { return _dtmBuptim; }
            set { _dtmBuptim = value; }
        }

        internal string[] StringItems
        {
            get
            {
                var list = _myInfo.Keys.Select(key => key + "=" + _myInfo[key]).ToList();
                list.Add("Creatim=" + Creatim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Revtim=" + Revtim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Printim=" + Printim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Buptim=" + Buptim.ToString("yyyy-MM-dd HH:mm:ss"));
                return list.ToArray();
            }
        }

        /// <summary>
        /// get information specify name
        /// </summary>
        /// <param name="strName">name</param>
        /// <returns>value</returns>
        public string GetInfo(string strName)
        {
            return _myInfo.ValueOrDefault(strName);
        }

        /// <summary>
        /// set information specify name
        /// </summary>
        /// <param name="strName">name</param>
        /// <param name="strValue">value</param>
        public void SetInfo(string strName, string strValue)
        {
            _myInfo[strName] = strValue;
        }

        public void Clear()
        {
            _myInfo.Clear();
            _dtmCreatim = DateTime.Now;
            _dtmRevtim = DateTime.Now;
            _dtmPrintim = DateTime.Now;
            _dtmBuptim = DateTime.Now;
        }

        public void Write(RTFWriter writer)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword("info");
            foreach (var strKey in _myInfo.Keys)
            {
                writer.WriteStartGroup();
                if (strKey == "edmins"
                    || strKey == "vern"
                    || strKey == "nofpages"
                    || strKey == "nofwords"
                    || strKey == "nofchars"
                    || strKey == "nofcharsws"
                    || strKey == "id")
                {
                    writer.WriteKeyword(strKey + _myInfo[strKey]);
                }
                else
                {
                    writer.WriteKeyword(strKey);
                    writer.WriteText(_myInfo[strKey]);
                }
                writer.WriteEndGroup();
            }
            writer.WriteStartGroup();

            WriteTime(writer, "creatim", _dtmCreatim);
            WriteTime(writer, "revtim", _dtmRevtim);
            WriteTime(writer, "printim", _dtmPrintim);
            WriteTime(writer, "buptim", _dtmBuptim);

            writer.WriteEndGroup();
        }

        private void WriteTime(RTFWriter writer, string name, DateTime value)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword(name);
            writer.WriteKeyword("yr" + value.Year);
            writer.WriteKeyword("mo" + value.Month);
            writer.WriteKeyword("dy" + value.Day);
            writer.WriteKeyword("hr" + value.Hour);
            writer.WriteKeyword("min" + value.Minute);
            writer.WriteKeyword("sec" + value.Second);
            writer.WriteEndGroup();
        }
    }
}