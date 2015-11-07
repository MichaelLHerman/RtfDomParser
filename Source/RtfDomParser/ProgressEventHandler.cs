/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */

using System;

namespace RtfDomParser
{
    /// <summary>
    /// progress event handler type
    /// </summary>
    /// <param name="sender">sender</param>
    /// <param name="args">event arguments</param>
    public delegate void ProgressEventHandler(object sender, ProgressEventArgs args);

    /// <summary>
    /// porgress event arguments
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        private readonly int _intMaxValue;

        private readonly int _intValue;

        private readonly string _strMessage;

        public ProgressEventArgs(int max, int value, string message)
        {
            Cancel = false;
            _intMaxValue = max;
            _intValue = value;
            _strMessage = message;
        }

        /// <summary>
        /// progress max value
        /// </summary>
        public int MaxValue
        {
            get { return _intMaxValue; }
        }

        /// <summary>
        /// current value
        /// </summary>
        public int Value
        {
            get { return _intValue; }
        }

        /// <summary>
        /// progress message
        /// </summary>
        public string Message
        {
            get { return _strMessage; }
        }

        /// <summary>
        /// cancel operation
        /// </summary>
        public bool Cancel { get; set; }
    }
}