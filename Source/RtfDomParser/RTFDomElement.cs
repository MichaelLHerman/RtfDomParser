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
using System.Text;
using System.Xml.Serialization;

namespace RtfDomParser
{
    /// <summary>
    /// RTF dom element
    /// </summary>
    /// <remarks>this is the most base element type</remarks>
    public abstract class RTFDomElement
    {
        private RTFAttributeList _myAttributes = new RTFAttributeList();

        private readonly List<RTFDomElement> _myElements = new List<RTFDomElement>();

        private RTFDomDocument _myOwnerDocument;

        /// <summary>
        /// Native level in RTF document
        /// </summary>
        internal int NativeLevel = -1;

        public RTFDomElement()
        {
            Locked = false;
            Parent = null;
        }

        /// <summary>
        /// RTF native attribute
        /// </summary>
        public RTFAttributeList Attributes
        {
            get { return _myAttributes; }
            set { _myAttributes = value; }
        }

        /// <summary>
        /// child elements list
        /// </summary>
        public List<RTFDomElement> Elements
        {
            get { return _myElements; }
        }

        /// <summary>
        /// the docuemnt which owned this element
        /// </summary>
        [XmlIgnore]
        public RTFDomDocument OwnerDocument
        {
            get { return _myOwnerDocument; }
            set
            {
                _myOwnerDocument = value;
                foreach (var element in Elements)
                {
                    element.OwnerDocument = value;
                }
            }
        }

        /// <summary>
        /// parent element
        /// </summary>
        public RTFDomElement Parent { get; private set; }


        public virtual string InnerText
        {
            get
            {
                var str = new StringBuilder();
                if (_myElements != null)
                {
                    foreach (var element in _myElements)
                    {
                        str.Append(element.InnerText);
                    }
                }
                return str.ToString();
            }
        }

        /// <summary>
        /// Whether element is locked , if element is lock , it can not append chidl element
        /// </summary>
        [XmlIgnore]
        public bool Locked { get; set; }

        public bool HasAttribute(string name)
        {
            return _myAttributes.Contains(name);
        }

        public int GetAttributeValue(string name, int defaultValue)
        {
            if (_myAttributes.Contains(name))
                return _myAttributes[name];
            return defaultValue;
        }

        /// <summary>
        /// append child element
        /// </summary>
        /// <param name="element">child element</param>
        /// <returns>index of element</returns>
        public void AppendChild(RTFDomElement element)
        {
            CheckLocked();
            element.Parent = this;
            element.OwnerDocument = _myOwnerDocument;
            _myElements.Add(element);
        }

        /// <summary>
        /// set attribute
        /// </summary>
        /// <param name="name">name</param>
        /// <param name="value">value</param>
        public void SetAttribute(string name, int value)
        {
            CheckLocked();
            _myAttributes[name] = value;
        }


        private void CheckLocked()
        {
            if (Locked)
            {
                throw new InvalidOperationException("Element locked");
            }
        }

        public void SetLockedDeeply(bool locked)
        {
            Locked = locked;
            if (_myElements != null)
            {
                foreach (var element in _myElements)
                {
                    element.SetLockedDeeply(locked);
                }
            }
        }


        public virtual string ToDomString()
        {
            var builder = new StringBuilder();
            builder.Append(ToString());
            ToDomString(Elements, builder, 1);
            return builder.ToString();
        }

        protected void ToDomString(List<RTFDomElement> elements, StringBuilder builder, int level)
        {
            foreach (var element in elements)
            {
                builder.Append(Environment.NewLine);
                builder.Append(new string(' ', level*4));
                builder.Append(element);
                ToDomString(element.Elements, builder, level + 1);
            }
        }
    }
}