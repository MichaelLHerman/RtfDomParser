using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfDomParser
{
    public class Font
    {
        public string Name { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        // Summary:
        //     Gets the em-size of this System.Drawing.Font measured in the units specified
        //     by the System.Drawing.Font.Unit property.
        //
        // Returns:
        //     The em-size of this System.Drawing.Font.
        public float Size { get; set; }
        public bool Strikeout { get; set; }
        public bool Underline { get; set; }
    }

}
