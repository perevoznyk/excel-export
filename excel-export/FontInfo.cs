// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace Export.XLS
{
    internal class FontInfo
    {
        public FontInfo(Font font, ExcelColor color)
        {
            this.Font = font;
            this.Color = color;
        }

        public Font Font { get; set; }
        public ExcelColor Color { get; set; }
        
        public override bool Equals(object obj)
        {
            if (obj is FontInfo)
            {
                FontInfo info = (FontInfo)obj;
                return (this.Font.Equals(info.Font) && (this.Color.Equals(info.Color)));
            }
            return false;
        }

        public override int GetHashCode()
        {
            return this.Font.GetHashCode() ^ this.Color.GetHashCode();
        }
    }
}
