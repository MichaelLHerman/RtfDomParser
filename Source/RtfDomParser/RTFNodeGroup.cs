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

namespace RtfDomParser
{
    /// <summary>
    /// RTF node group , this source code evolution from other software.
    /// </summary>
    public class RTFNodeGroup : RTFNode
    {
        /// <summary>
        /// child node list
        /// </summary>
        protected RTFNodeList MyNodes = new RTFNodeList();

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFNodeGroup()
        {
            IntType = RTFNodeType.Group;
        }

        /// <summary>
        /// child node list
        /// </summary>
        public override RTFNodeList Nodes
        {
            get { return MyNodes; }
        }

        /// <summary>
        /// first child node
        /// </summary>
        public RTFNode FirstNode
        {
            get
            {
                if (MyNodes.Count > 0)
                    return MyNodes[0];
                return null;
            }
        }

        public override string Keyword
        {
            get
            {
                if (MyNodes.Count > 0)
                    return MyNodes[0].Keyword;
                return null;
            }
            set { }
        }

        public override bool HasParameter
        {
            get
            {
                if (MyNodes.Count > 0)
                    return MyNodes[0].HasParameter;
                return false;
            }
            set { }
        }

        public override int Parameter
        {
            get
            {
                if (MyNodes.Count > 0)
                    return MyNodes[0].Parameter;
                return 0;
            }
        }


        public virtual string Text
        {
            get
            {
                var myStr = new StringBuilder();
                foreach (var node in MyNodes)
                {
                    if (node is RTFNodeGroup)
                    {
                        myStr.Append(((RTFNodeGroup) node).Text);
                    }
                    if (node.Type == RTFNodeType.Text)
                        myStr.Append(node.Keyword);
                }
                return myStr.ToString();
            }
        }

        /// <summary>
        /// get all child node deeply
        /// </summary>
        /// <param name="includeGroupNode">contains group type node</param>
        /// <returns>child nodes list</returns>
        public RTFNodeList GetAllNodes(bool includeGroupNode)
        {
            var list = new RTFNodeList();
            AddAllNodes(list, includeGroupNode);
            return list;
        }

        private void AddAllNodes(RTFNodeList list, bool includeGroupNode)
        {
            foreach (var node in MyNodes)
            {
                if (node is RTFNodeGroup)
                {
                    if (includeGroupNode)
                        list.Add(node);
                    ((RTFNodeGroup) node).AddAllNodes(list, includeGroupNode);
                }
                else
                    list.Add(node);
            }
        }

        internal void MergeText()
        {
            var list = new RTFNodeList();
            var myStr = new StringBuilder();
            var buffer = new ByteBuffer();
            //System.IO.MemoryStream ms = new System.IO.MemoryStream();
            //System.Text.Encoding encode = myOwnerDocument.Encoding ;
            foreach (var node in MyNodes)
            {
                if (node.Type == RTFNodeType.Text)
                {
                    AddString(myStr, buffer);
                    myStr.Append(node.Keyword);
                    continue;
                }
                if (node.Type == RTFNodeType.Control
                    && node.Keyword == "\'"
                    && node.HasParameter)
                {
                    buffer.Add((byte) node.Parameter);
                    continue;
                }
                if (node.Type == RTFNodeType.Control || node.Type == RTFNodeType.Keyword)
                {
                    if (node.Keyword == "tab")
                    {
                        AddString(myStr, buffer);
                        myStr.Append('\t');
                        continue;
                    }
                    if (node.Keyword == "emdash")
                    {
                        AddString(myStr, buffer);
                        // notice!! This code may cause compiler error in OS which not support chinese character.
                        // you can change to myStr.Append('-');
                        myStr.Append('j');
                        continue;
                    }
                    if (node.Keyword == "")
                    {
                        AddString(myStr, buffer);
                        // notice!! This code may cause compiler error in OS which not support chinese character.
                        // you can change to myStr.Append('-');
                        myStr.Append('Ƀ');
                        continue;
                    }
                }
                AddString(myStr, buffer);
                if (myStr.Length > 0)
                {
                    list.Add(new RTFNode(RTFNodeType.Text, myStr.ToString()));
                    myStr = new StringBuilder();
                }
                list.Add(node);
            } //foreach( RTFNode node in myNodes )

            AddString(myStr, buffer);
            if (myStr.Length > 0)
            {
                list.Add(new RTFNode(RTFNodeType.Text, myStr.ToString()));
            }
            MyNodes.Clear();
            foreach (var node in list)
            {
                node.Parent = this;
                node.OwnerDocument = MyOwnerDocument;
                MyNodes.Add(node);
            }
        }

        private void AddString(StringBuilder myStr, ByteBuffer buffer)
        {
            if (buffer.Count > 0)
            {
                //if( buffer.Count == 1 )
                //{
                //    myStr.Append( ( char ) buffer[0] );
                //}
                //else
                {
                    var txt = buffer.GetString(MyOwnerDocument.RuntimeEncoding);
                    myStr.Append(txt);
                }
                buffer.Reset();
            }
        }

        /// <summary>
        /// write content to rtf document
        /// </summary>
        /// <param name="writer">RTF text writer</param>
        public override void Write(RTFWriter writer)
        {
            writer.WriteStartGroup();
            foreach (var node in MyNodes)
            {
                node.Write(writer);
            }
            writer.WriteEndGroup();
        }

        /// <summary>
        /// search child node special keyword
        /// </summary>
        /// <param name="key">special keyword</param>
        /// <param name="deeply">whether search deeplyl</param>
        /// <returns>node find</returns>
        public RTFNode SearchKey(string key, bool deeply)
        {
            foreach (var node in MyNodes)
            {
                if (node.Type == RTFNodeType.Keyword
                    || node.Type == RTFNodeType.ExtKeyword
                    || node.Type == RTFNodeType.Control)
                {
                    if (node.Keyword == key)
                        return node;
                }
                if (deeply)
                {
                    if (node is RTFNodeGroup)
                    {
                        var g = (RTFNodeGroup) node;
                        var n = g.SearchKey(key, true);
                        if (n != null)
                            return n;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// append child node
        /// </summary>
        /// <param name="node">node</param>
        public void AppendChild(RTFNode node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException("node");
            if (node == this)
                throw new ArgumentException("node != this");
            node.Parent = this;
            node.OwnerDocument = MyOwnerDocument;
            Nodes.Add(node);
        }

        /// <summary>
        /// delete node
        /// </summary>
        /// <param name="node">node</param>
        public void RemoveChild(RTFNode node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException("node");
            if (node == this)
                throw new ArgumentException("node != this");
            Nodes.Remove(node);
        }

        /// <summary>
        /// insert node
        /// </summary>
        /// <param name="index">index</param>
        /// <param name="node">node</param>
        public void InsertNode(int index, RTFNode node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException("node");
            if (node == this)
                throw new ArgumentException("node != this");
            node.Parent = this;
            node.OwnerDocument = MyOwnerDocument;
            Nodes.Insert(index, node);
        }

        private void CheckNodes()
        {
            if (Nodes == null)
                throw new Exception("child node is invalidate");
        }
    }
}