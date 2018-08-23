// C# Excel Writer library v2.1
// by Serhiy Perevoznyk, 2008-2018

namespace Export.XLS
{
    public class Font
    {
        private readonly string name;
        private readonly float size;
        private FontStyle fontStyle;

        public string Name { get { return this.name; } }
        public float Size { get { return this.size; } }

        public bool Bold
        {
            get
            {
                return (Style & FontStyle.Bold) != 0;
            }
        }

        public bool Italic
        {
            get
            {
                return (Style & FontStyle.Italic) != 0;
            }
        }

        public bool Underline
        {
            get
            {
                return (Style & FontStyle.Underline) != 0;
            }
        }

        public bool Strikeout
        {
            get
            {
                return (Style & FontStyle.Strikeout) != 0;
            }
        }

        public Font(string familyName, float emSize)
        {
            this.name = familyName;
            this.size = emSize;
        }

        public Font(string familyName, float emSize, FontStyle style)
        {
            this.name = familyName;
            this.size = emSize;
            this.fontStyle = style;
        }

        public FontStyle Style
        {
            get
            {
                return fontStyle;
            }
        }

    }
}
