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
            TValue value = default(TValue);
            dictionary.TryGetValue(key, out value);
            return value;
        }
    }

	/// <summary>
	/// Document information
	/// </summary>
	public class RTFDocumentInfo 
	{
		private Dictionary<string,string> myInfo =
            new Dictionary<string, string>();
		/// <summary>
		/// get information specify name
		/// </summary>
		/// <param name="strName">name</param>
		/// <returns>value</returns>
		public string GetInfo( string strName )
		{
			return myInfo.ValueOrDefault(strName);
		}

		/// <summary>
		/// set information specify name
		/// </summary>
		/// <param name="strName">name</param>
		/// <param name="strValue">value</param>
		public void SetInfo( string strName , string strValue )
		{
			myInfo[ strName ] = strValue ;
		}
		/// <summary>
		/// document title
		/// </summary>
		public string Title
		{
			get{ return myInfo.ValueOrDefault("title") ;}
			set{ myInfo["title"] = value;}
		}
		public string Subject
		{
			get{ return myInfo.ValueOrDefault("subject") ;}
			set{ myInfo["subject"] = value;}
		}
		public string Author
		{
			get{ return myInfo.ValueOrDefault("author") ;}
			set{ myInfo["author"] = value;}
		}
		public string Manager
		{
			get{ return myInfo.ValueOrDefault("manager") ;}
			set{ myInfo["manager"] = value;}
		}
		public string Company
		{
			get{ return myInfo.ValueOrDefault("company") ;}
			set{ myInfo["company"] = value;}
		}
		public string Operator
		{
			get{ return myInfo.ValueOrDefault("operator") ;}
			set{ myInfo["operator"] = value;}
		}
		public string Category
		{
			get{ return myInfo.ValueOrDefault("category") ;}
			set{ myInfo["categroy"] = value;}
		}
		public string Keywords
		{
			get{ return myInfo.ValueOrDefault("keywords") ;}
			set{ myInfo["keywords"] = value;}
		}
		public string Comment
		{
			get{ return myInfo.ValueOrDefault("comment");}
			set{ myInfo["comment"] = value;}
		}
		public string Doccomm
		{
			get{ return myInfo.ValueOrDefault("doccomm") ;}
			set{ myInfo["doccomm"] = value;}
		}
		public string HLinkbase
		{
			get{ return myInfo.ValueOrDefault("hlinkbase") ;}
			set{ myInfo["hlinkbase"] = value;}
		}

        /// <summary>
        /// total edit minutes
        /// </summary>
        public int edmins
        {
            get 
            {
                if (myInfo.ContainsKey("edmins"))
                {
                    string v = Convert.ToString(myInfo["edmins"]);
                    int result = 0;
                    if (int.TryParse(v, out result))
                    {
                        return result;
                    }
                }
                return 0;
            }
            set
            {
                myInfo["edmins"] = value.ToString(); 
            }
        }

        /// <summary>
        /// version
        /// </summary>
        public string vern
        {
            get { return myInfo["vern"]; }
            set { myInfo["vern"] = value; }
        }

        /// <summary>
        /// number of pages
        /// </summary>
        public string nofpages
        {
            get { return myInfo["nofpages"]; }
            set { myInfo["nofpages"] = value; }
        }

        /// <summary>
        /// number of words
        /// </summary>
        public string nofwords
        {
            get { return myInfo["nofwords"]; }
            set { myInfo["nofwords"] = value; }
        }

        /// <summary>
        /// number of characters , include whitespace
        /// </summary>
        public string nofchars
        {
            get { return myInfo["nofchars"]; }
            set { myInfo["nofchars"] = value; }
        }

        /// <summary>
        /// number of characters , exclude white space
        /// </summary>
        public string nofcharsws
        {
            get { return myInfo["nofcharsws"]; }
            set { myInfo["nofcharsws"] = value; }
        }

        /// <summary>
        /// inner id
        /// </summary>
        public string id
        {
            get { return myInfo["id"]; }
            set { myInfo["id"] = value; }
        }

		private DateTime dtmCreatim = DateTime.Now ;
		/// <summary>
		/// creation time
		/// </summary>
		public DateTime Creatim
		{
			get{ return dtmCreatim ;}
			set{ dtmCreatim = value;}
		}

		private DateTime dtmRevtim = DateTime.Now ;
		/// <summary>
		/// modified time
		/// </summary>
		public DateTime Revtim
		{
			get{ return dtmRevtim ;}
			set{ dtmRevtim = value;}
		}

		private DateTime dtmPrintim = DateTime.Now ;
		/// <summary>
		/// last print time
		/// </summary>
		public DateTime Printim
		{
			get{ return dtmPrintim ;}
			set{ dtmPrintim = value;}
		}

		private DateTime dtmBuptim = DateTime.Now ;
		/// <summary>
		/// back up time
		/// </summary>
		public DateTime Buptim
		{
			get{ return dtmBuptim ;}
			set{ dtmBuptim = value;}
		}

        internal string[] StringItems
        {
            get
            {
                var list = new List<string>();
                foreach (string key in myInfo.Keys)
                {
                    list.Add(key + "=" + myInfo[key]);
                }
                list.Add("Creatim="+this.Creatim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Revtim="+ this.Revtim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Printim="+ this.Printim.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Buptim="+ this.Buptim.ToString("yyyy-MM-dd HH:mm:ss"));
                return list.ToArray();
            }
        }

		public void Clear()
		{
			myInfo.Clear();
			dtmCreatim = System.DateTime.Now ;
			dtmRevtim = DateTime.Now ;
			dtmPrintim = DateTime.Now ;
			dtmBuptim = DateTime.Now ;
		}

		public void Write(RTFWriter writer)
		{
			writer.WriteStartGroup( );
			writer.WriteKeyword("info");
			foreach( string strKey in myInfo.Keys )
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
                    writer.WriteKeyword(strKey + myInfo[strKey]);
                }
                else
                {
                    writer.WriteKeyword(strKey);
                    writer.WriteText(myInfo[strKey]);
                }
				writer.WriteEndGroup();
			}
			writer.WriteStartGroup();

			WriteTime( writer , "creatim" , dtmCreatim );
			WriteTime( writer , "revtim" , dtmRevtim );
			WriteTime( writer , "printim" , dtmPrintim );
			WriteTime( writer , "buptim" , dtmBuptim );

			writer.WriteEndGroup();
		}
		
		private void WriteTime( RTFWriter writer , string name , DateTime Value )
		{
			writer.WriteStartGroup();
			writer.WriteKeyword( name );
			writer.WriteKeyword( "yr" + Value.Year );
			writer.WriteKeyword( "mo" + Value.Month );
			writer.WriteKeyword( "dy" + Value.Day );
			writer.WriteKeyword( "hr" + Value.Hour );
			writer.WriteKeyword( "min" + Value.Minute );
			writer.WriteKeyword( "sec" + Value.Second );
			writer.WriteEndGroup();
		}
	}
}