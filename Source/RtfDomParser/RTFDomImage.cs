/*
 * 
 *   DCSoft RTF DOM v1.0
 *   Author : Yuan yong fu.
 *   Email  : yyf9989@hotmail.com
 *   blog site:http://www.cnblogs.com/xdesigner.
 * 
 */


using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace RtfDomParser
{
    /// <summary>
    /// image element
    /// </summary>
    public class RTFDomImage : RTFDomElement
    {
        private int _intScaleX = 100;

        private int _intScaleY = 100;

        private DocumentFormatInfo _myFormat = new DocumentFormatInfo();

        private RTFPicType _picType = RTFPicType.Jpegblip;

        /// <summary>
        /// initialize instance
        /// </summary>
        public RTFDomImage()
        {
            Height = 0;
            Width = 0;
            DesiredHeight = 0;
            Data = null;
            Id = null;
            DesiredWidth = 0;
        }

        /// <summary>
        /// id
        /// </summary>
        [DefaultValue(null)]
        public string Id { get; set; }

        /// <summary>
        /// data
        /// </summary>
        [XmlIgnore]
        public byte[] Data { get; set; }


        [XmlElement]
        public string Base64Data
        {
            get
            {
                return Data == null ? null : Convert.ToBase64String(Data);
            }
            set {
                Data = !string.IsNullOrEmpty(value) ? Convert.FromBase64String(value) : null;
            }
        }

        /// <summary>
        /// scale rate at the X coordinate, in percent unit.
        /// </summary>
        [DefaultValue(100)]
        public int ScaleX
        {
            get { return _intScaleX; }
            set { _intScaleX = value; }
        }

        /// <summary>
        /// scale rate at the Y coordinate , in percent unit.
        /// </summary>
        [DefaultValue(100)]
        public int ScaleY
        {
            get { return _intScaleY; }
            set { _intScaleY = value; }
        }

        /// <summary>
        /// desired width
        /// </summary>
        [DefaultValue(0)]
        public int DesiredWidth { get; set; }

        /// <summary>
        /// desired height
        /// </summary>
        [DefaultValue(0)]
        public int DesiredHeight { get; set; }

        /// <summary>
        /// width
        /// </summary>
        [DefaultValue(0)]
        public int Width { get; set; }

        /// <summary>
        /// height
        /// </summary>
        [DefaultValue(0)]
        public int Height { get; set; }

        public RTFPicType PicType
        {
            get { return _picType; }
            set { _picType = value; }
        }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format
        {
            get { return _myFormat; }
            set { _myFormat = value; }
        }

        public override string ToString()
        {
            var txt = "Image:" + Width + "*" + Height;
            if (Data != null && Data.Length > 0)
                txt = txt + " " + Convert.ToDouble(Data.Length/1024.0).ToString("0.00") + "KB";
            return txt;
        }
    }

    public enum RTFPicType
    {
        /// <summary>
        /// Source of the picture is an EMF (enhanced metafile).
        /// </summary>
        Emfblip,

        /// <summary>
        /// Source of the picture is a PNG.
        /// </summary>
        Pngblip,

        /// <summary>
        /// Source of the picture is a JPEG.
        /// </summary>
        Jpegblip,

        /// <summary>
        /// ource of the picture is QuickDraw.
        /// </summary>
        Macpict,

        /// <summary>
        /// Source of the picture is an OS/2 metafile. The N argument identifies the metafile type. The N values are described in the \pmmetafile table further on in this section.
        /// </summary>
        Pmmetafile,

        /// <summary>
        /// Source of the picture is a Windows metafile. The N argument identifies the metafile type (the default type is 1).
        /// </summary>
        Wmetafile,

        /// <summary>
        /// Source of the picture is a Windows device-independent bitmap. The N argument identifies the bitmap type, which must equal 0.
        /// The information to be included in RTF from a Windows device-independent bitmap is the concatenation of the BITMAPINFO structure followed by the actual pixel data.
        /// </summary>
        Dibitmap,

        /// <summary>
        /// Source of the picture is a Windows device-dependent bitmap. The N argument identifies the bitmap type (must equal 0).
        /// The information to be included in RTF from a Windows device-dependent bitmap is the result of the GetBitmapBits function.
        /// </summary>
        Wbitmap
    }
}