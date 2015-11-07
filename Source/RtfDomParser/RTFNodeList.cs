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

namespace RtfDomParser
{
	/// <summary>
    /// RTF node list , this source code evolution from other software.
	/// </summary>
    [System.Diagnostics.DebuggerDisplay("Count={ Count }")]
    public class RTFNodeList : List<RTFNode> 
	{

		/// <summary>
		/// get node special keyword
		/// </summary>
		public RTFNode this[ string KeyWord ]
		{
			get
			{
				foreach( RTFNode node in this )
				{
					if( node.Keyword == KeyWord )
						return node ;
				}
				return null ;
			}
		}
		
		/// <summary>
		/// get node special type
		/// </summary>
		public RTFNode this[ Type t ]
		{
			get
			{
				foreach( RTFNode node in this )
				{
					if( t.Equals( node.GetType()))
						return node ;
				}
				return null;
			}
		}

		/// <summary>
		/// get node's parameter value special keyword, if can not find , return default value.
		/// </summary>
		/// <param name="key">keyword</param>
		/// <param name="DefaultValue">default value</param>
		/// <returns>parameter value</returns>
		public int GetParameter( string key , int DefaultValue )
		{
			foreach( RTFNode node in this )
			{
				if( node.Keyword == key && node.HasParameter )
					return node.Parameter ;
			}
			return DefaultValue ;
		}

		public string Text
		{
			get
			{
				System.Text.StringBuilder myStr = new System.Text.StringBuilder();
				foreach( RTFNode node in this )
				{
					if( node.Type == RTFNodeType.Text )
					{
						myStr.Append( node.Keyword );
					}
					else if( node is RTFNodeGroup )
					{
						string txt = node.Nodes.Text ;
						if( txt != null )
							myStr.Append( txt );
					}
				}
				return myStr.ToString();
			}
		}

		/// <summary>
		/// detect whether exist node special keyword in this list
		/// </summary>
		/// <param name="Key">keyword</param>
		/// <returns>exist or not</returns>
		public bool ContainsKey( string Key )
		{
			return this[ Key ] != null;
		}

	}
}