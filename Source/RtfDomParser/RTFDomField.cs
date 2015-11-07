/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System.ComponentModel;
using System.Linq;

namespace RtfDomParser
{
    public class RTFDomField : RTFDomElement
    {
        private RTFDomFieldMethod _intMethod = RTFDomFieldMethod.None;

        /// <summary>
        /// method
        /// </summary>
        [DefaultValue(RTFDomFieldMethod.None)]
        public RTFDomFieldMethod Method
        {
            get { return _intMethod; }
            set { _intMethod = value; }
        }

        //private string strInstructions = null;
        /// <summary>
        /// instructions
        /// </summary>
        [DefaultValue(null)]
        public string Instructions
        {
            get
            {
                return (from c in Elements.OfType<RTFDomElementContainer>() where c.Name == RTFConsts.Fldinst select c.InnerText).FirstOrDefault();
            }
            //set
            //{
            //    strInstructions = value;
            //}
        }

        /// <summary>
        /// result
        /// </summary>
        [DefaultValue(null)]
        public RTFDomElementContainer Result
        {
            get
            {
                return Elements.OfType<RTFDomElementContainer>().FirstOrDefault(c => c.Name == RTFConsts.Fldrslt);
            }
            //set
            //{
            //    strResult = value;
            //}
        }

        public string ResultString
        {
            get
            {
                var c = Result;
                return c != null ? c.InnerText : null;
            }
        }

        public override string ToString()
        {
            return "Field"; // +strInstructions + " Result:" + this.ResultString;
        }
    } 


    public enum RTFDomFieldMethod
    {
        None,
        Dirty,
        Edit,
        Lock,
        Priv
    }
}