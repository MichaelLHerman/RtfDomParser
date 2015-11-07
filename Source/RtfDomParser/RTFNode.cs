/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


namespace RtfDomParser
{
    /// <summary>
    /// RTF parser node, this source code evolution from other software.
    /// </summary>
    public class RTFNode
    {
        /// <summary>
        /// Whether this node has parameter
        /// </summary>
        protected bool BolHasParameter;

        protected int IntParameter;

        protected RTFNodeType IntType = RTFNodeType.None;

        protected RTFRawDocument MyOwnerDocument;

        protected RTFNodeGroup MyParent;

        /// <summary>
        /// key word
        /// </summary>
        protected string StrKeyword;

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFNode()
        {
        }

        public RTFNode(RTFNodeType type, string key)
        {
            IntType = type;
            StrKeyword = key;
        }

        public RTFNode(RTFToken token)
        {
            StrKeyword = token.Key;
            BolHasParameter = token.HasParam;
            IntParameter = token.Param;
            if (token.Type == RTFTokenType.Control)
                IntType = RTFNodeType.Control;
            else if (token.Type == RTFTokenType.Control)
                IntType = RTFNodeType.Control;
            else if (token.Type == RTFTokenType.Keyword)
                IntType = RTFNodeType.Keyword;
            else if (token.Type == RTFTokenType.ExtKeyword)
                IntType = RTFNodeType.ExtKeyword;
            else if (token.Type == RTFTokenType.Text)
                IntType = RTFNodeType.Text;
            else
                IntType = RTFNodeType.Text;
        }

        /// <summary>
        /// parent node
        /// </summary>
        public virtual RTFNodeGroup Parent
        {
            get { return MyParent; }
            set { MyParent = value; }
        }

        /// <summary>
        /// raw document which owner this node
        /// </summary>
        public virtual RTFRawDocument OwnerDocument
        {
            get { return MyOwnerDocument; }
            set
            {
                MyOwnerDocument = value;
                if (Nodes != null)
                {
                    foreach (var node in Nodes)
                    {
                        node.OwnerDocument = value;
                    }
                }
            }
        }

        /// <summary>
        /// Keyword
        /// </summary>
        public virtual string Keyword
        {
            get { return StrKeyword; }
            set { StrKeyword = value; }
        }

        /// <summary>
        /// Whether this node has parameter
        /// </summary>
        public virtual bool HasParameter
        {
            get { return BolHasParameter; }
            set { BolHasParameter = value; }
        }

        /// <summary>
        /// paramter value
        /// </summary>
        public virtual int Parameter
        {
            get { return IntParameter; }
        }

        /// <summary>
        /// child nodes
        /// </summary>
        public virtual RTFNodeList Nodes
        {
            get { return null; }
        }


        /// <summary>
        /// index
        /// </summary>
        public int Index
        {
            get
            {
                if (MyParent == null)
                    return 0;
                return MyParent.Nodes.IndexOf(this);
            }
        }

        /// <summary>
        /// node type
        /// </summary>
        public RTFNodeType Type
        {
            get { return IntType; }
        }

        /// <summary>
        /// previous node in parent nodes list
        /// </summary>
        public RTFNode PreviousNode
        {
            get
            {
                if (MyParent != null)
                {
                    var index = MyParent.Nodes.IndexOf(this);
                    if (index > 0)
                    {
                        return MyParent.Nodes[index - 1];
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// next node in parent nodes list
        /// </summary>
        public RTFNode NextNode
        {
            get
            {
                if (MyParent != null)
                {
                    var index = MyParent.Nodes.IndexOf(this);
                    if (index >= 0 && index < MyParent.Nodes.Count - 1)
                        return MyParent.Nodes[index + 1];
                }
                return null;
            }
        }

        /// <summary>
        /// write to rtf document
        /// </summary>
        /// <param name="writer">RTF text writer</param>
        public virtual void Write(RTFWriter writer)
        {
            if (IntType == RTFNodeType.Control
                || IntType == RTFNodeType.Keyword
                || IntType == RTFNodeType.ExtKeyword)
            {
                if (BolHasParameter)
                {
                    writer.WriteKeyword(
                        StrKeyword + IntParameter,
                        IntType == RTFNodeType.ExtKeyword);
                }
                else
                {
                    writer.WriteKeyword(
                        StrKeyword,
                        IntType == RTFNodeType.ExtKeyword);
                }
            }
            else if (IntType == RTFNodeType.Text)
            {
                writer.WriteText(StrKeyword);
            }
        }
    }
}