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
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace RtfDomParser
{
    public class RTFDomObject : RTFDomElement
    {
        private Dictionary<string, string> _customAttributes = new Dictionary<string, string>();

        private int _scaleX = 100;

        private int _scaleY = 100;

        //private Dictionary<string, string> _Attributes = new Dictionary<string, string>();
        //public Dictionary<string, string> Attributes1
        //{
        //    get
        //    {
        //        if (_Attributes == null)
        //        {
        //            _Attributes = new Dictionary<string, string>();
        //        }
        //        return _Attributes; 
        //    }
        //    set
        //    {
        //        _Attributes = value; 
        //    }
        //}

        private RTFObjectType _type = RTFObjectType.Emb;


        public Dictionary<string, string> CustomAttributes
        {
            get { return _customAttributes ?? (_customAttributes = new Dictionary<string, string>()); }
            set { _customAttributes = value; }
        }

        public RTFObjectType Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public string ClassName { get; set; }

        public string Name { get; set; }

        public byte[] Content { get; set; }


        public string ContentText
        {
            get
            {
                if (Content == null || Content.Length == 0)
                {
                    return null;
                }
                return Encoding.UTF8.GetString(Content, 0, Content.Length);
            }
        }

        public int Width { get; set; }

        public int Height { get; set; }

        public int ScaleX
        {
            get { return _scaleX; }
            set { _scaleX = value; }
        }

        public int ScaleY
        {
            get { return _scaleY; }
            set { _scaleY = value; }
        }

        /// <summary>
        /// result
        /// </summary>
        [DefaultValue(null)]
        public RTFDomElementContainer Result
        {
            get
            { return Elements.OfType<RTFDomElementContainer>().FirstOrDefault(c => c.Name == RTFConsts.Result); }
            //set
            //{
            //    strResult = value;
            //}
        }

        public override string ToString()
        {
            var txt = "Object:" + Width + "*" + Height;
            if (Content != null && Content.Length > 0)
            {
                txt = txt + " " + Convert.ToDouble(Content.Length/1024.0).ToString("0.00") + "KB";
            }
            return txt;
        }
    }

    public enum RTFObjectType
    {
        Emb,
        Link,
        AutLink,
        Sub,
        Pub,
        Icemb,
        Html,
        Ocx
    }
}