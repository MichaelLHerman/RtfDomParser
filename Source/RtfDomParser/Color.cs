using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RtfDomParser
{
    public struct Color
    {
        public static readonly Color Empty;

        public Color(byte a, byte r, byte g, byte b)
            : this()
        {
            A = a;
            R = r;
            G = g;
            B = b;
            _isInitialized = true;
        }

        public byte A { get; set; }
        public byte R { get; set; }
        public byte G { get; set; }
        public byte B { get; set; }

        // Summary:
        //     Tests whether two specified System.Drawing.Color structures are different.
        //
        // Parameters:
        //   left:
        //     The System.Drawing.Color that is to the left of the inequality operator.
        //
        //   right:
        //     The System.Drawing.Color that is to the right of the inequality operator.
        //
        // Returns:
        //     true if the two System.Drawing.Color structures are different; otherwise,
        //     false.
        public static bool operator !=(Color left, Color right)
        {
            return left.A != right.A || left.R != right.R || left.G != right.G || left.B == right.B;
        }
        //
        // Summary:
        //     Tests whether two specified System.Drawing.Color structures are equivalent.
        //
        // Parameters:
        //   left:
        //     The System.Drawing.Color that is to the left of the equality operator.
        //
        //   right:
        //     The System.Drawing.Color that is to the right of the equality operator.
        //
        // Returns:
        //     true if the two System.Drawing.Color structures are equal; otherwise, false.
        public static bool operator ==(Color left, Color right)
        {
            return left.A == right.A && left.R == right.R && left.G == right.G && left.B == right.B;
        }



        //
        // Summary:
        //     Gets a system-defined color that has an ARGB value of #FF000000.
        //
        // Returns:
        //     A System.Drawing.Color representing a system-defined color.
        public static Color Black { get { return new Color(255, 0, 0, 0); } }
        public static Color Transparent { get { return new Color(0, 0, 0, 0); } }

        private bool _isInitialized;
        public bool IsEmpty
        {
            get
            {
                return !_isInitialized;
            }
        }

        // Summary:
        //     Tests whether the specified object is a System.Drawing.Color structure and
        //     is equivalent to this System.Drawing.Color structure.
        //
        // Parameters:
        //   obj:
        //     The object to test.
        //
        // Returns:
        //     true if obj is a System.Drawing.Color structure equivalent to this System.Drawing.Color
        //     structure; otherwise, false.
        public override bool Equals(object obj)
        {
            return obj is Color && (Color)obj == this;
        }
        //
        // Summary:
        //     Creates a System.Drawing.Color structure from a 32-bit ARGB value.
        //
        // Parameters:
        //   argb:
        //     A value specifying the 32-bit ARGB value.
        //
        // Returns:
        //     The System.Drawing.Color structure that this method creates.
        public static Color FromArgb(int argb)
        {
            return new Color((byte)(argb >> 24),
                (byte)(argb >> 16),
                (byte)(argb >> 8),
                (byte)argb);
        }

        public static Color FromArgb(int alpha, Color baseColor)
        {
            return new Color((byte)alpha, baseColor.R, baseColor.G, baseColor.B);
        }

        //
        // Summary:
        //     Creates a System.Drawing.Color structure from the specified 8-bit color values
        //     (red, green, and blue). The alpha value is implicitly 255 (fully opaque).
        //     Although this method allows a 32-bit value to be passed for each color component,
        //     the value of each component is limited to 8 bits.
        //
        // Parameters:
        //   red:
        //     The red component value for the new System.Drawing.Color. Valid values are
        //     0 through 255.
        //
        //   green:
        //     The green component value for the new System.Drawing.Color. Valid values
        //     are 0 through 255.
        //
        //   blue:
        //     The blue component value for the new System.Drawing.Color. Valid values are
        //     0 through 255.
        //
        // Returns:
        //     The System.Drawing.Color that this method creates.
        //
        // Exceptions:
        //   System.ArgumentException:
        //     red, green, or blue is less than 0 or greater than 255.
        public static Color FromArgb(int red, int green, int blue)
        {
            return new Color(255, (byte)red, (byte)green, (byte)blue);
        }
        //
        // Summary:
        //     Creates a System.Drawing.Color structure from the four ARGB component (alpha,
        //     red, green, and blue) values. Although this method allows a 32-bit value
        //     to be passed for each component, the value of each component is limited to
        //     8 bits.
        //
        // Parameters:
        //   alpha:
        //     The alpha component. Valid values are 0 through 255.
        //
        //   red:
        //     The red component. Valid values are 0 through 255.
        //
        //   green:
        //     The green component. Valid values are 0 through 255.
        //
        //   blue:
        //     The blue component. Valid values are 0 through 255.
        //
        // Returns:
        //     The System.Drawing.Color that this method creates.
        //
        // Exceptions:
        //   System.ArgumentException:
        //     alpha, red, green, or blue is less than 0 or greater than 255.
        public static Color FromArgb(int alpha, int red, int green, int blue)
        {
            return new Color((byte)alpha, (byte)red, (byte)green, (byte)blue);
        }
        //
        // Summary:
        //     Returns a hash code for this System.Drawing.Color structure.
        //
        // Returns:
        //     An integer value that specifies the hash code for this System.Drawing.Color.
        public override int GetHashCode()
        {
            return A << 24 ^ R << 16 ^ G << 8 ^ R;
        }
        //
        // Summary:
        //     Gets the 32-bit ARGB value of this System.Drawing.Color structure.
        //
        // Returns:
        //     The 32-bit ARGB value of this System.Drawing.Color.
        public int ToArgb()
        {
            return A << 24 ^ R << 16 ^ G << 8 ^ R;
        }

    }
}
