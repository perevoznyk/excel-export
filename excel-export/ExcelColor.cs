// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018

using System.Drawing;

namespace Export.XLS
{
    public struct ExcelColor
    {
        private readonly Color color;
        private ushort index;

        
        internal ExcelColor(Color color, ushort index)
        {
            this.color = color;
            this.index = index;
        }

        public ushort Index
        {
            get { return this.index; }
        }

        public Color Color
        {
            get { return this.color; }
        }

        public static ExcelColor Black
        {
            get { return new ExcelColor(Color.Black, 0); }
        }

        public static ExcelColor White
        {
            get { return new ExcelColor(Color.White, 1); }
        }

        public static ExcelColor Red
        {
            get { return new ExcelColor(Color.Red, 2); }
        }

        public static ExcelColor Green
        {
            get { return new ExcelColor(Color.Green, 3); }
        }

        public static ExcelColor Blue
        {
            get { return new ExcelColor(Color.Blue, 4); }
        }

        public static ExcelColor Yellow
        {
            get { return new ExcelColor(Color.Yellow, 5); }
        }

        public static ExcelColor Magenta
        {
            get { return new ExcelColor(Color.Magenta, 6); }
        }

        public static ExcelColor Cyan
        {
            get { return new ExcelColor(Color.Cyan, 7); }
        }

        public static ExcelColor DarkRed
        {
            get { return new ExcelColor(Color.DarkRed, 0x10); }
        }

        public static ExcelColor DarkGreen
        {
            get { return new ExcelColor(Color.DarkGreen, 0x11); }
        }

        public static ExcelColor DarkBlue
        {
            get { return new ExcelColor(Color.DarkBlue, 0x12); }
        }

        public static ExcelColor Olive
        {
            get { return new ExcelColor(Color.Olive, 0x13); }
        }

        public static ExcelColor Purple
        {
            get { return new ExcelColor(Color.Purple, 0x14); }
        }

        public static ExcelColor Teal
        {
            get { return new ExcelColor(Color.Teal, 0x15); }
        }

        public static ExcelColor Silver
        {
            get { return new ExcelColor(Color.Silver, 0x16); }
        }

        public static ExcelColor Gray
        {
            get { return new ExcelColor(Color.Gray, 0x17); }
        }

        public static ExcelColor WindowText
        {
            get { return new ExcelColor(Color.Black, 0x18); }
        }

        public static ExcelColor WindowBackground
        {
            get { return new ExcelColor(Color.White, 0x19); }
        }

        public static ExcelColor Automatic
        {
            get { return new ExcelColor(Color.Black, BIFF.DefaultColor); }
        }

        public override bool Equals(object obj)
        {
            if (obj is ExcelColor)
            {
                return (this.index == ((ExcelColor)obj).index);
            }
            else
                return false;
        }

        public override int GetHashCode()
        {
            return this.index;
        }
    }
}
