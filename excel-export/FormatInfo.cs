// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018


namespace Export.XLS
{
    internal class FormatInfo
    {
        private CellInfo cell;
        private int formatIndex;
        private int fontIndex;

        public FormatInfo(CellInfo cell)
        {
            this.cell = cell;
            if (string.IsNullOrEmpty(cell.Format))
                formatIndex = 0;
            else
                formatIndex = cell.Document.Formats.IndexOf(cell.Format);

            FontInfo fontInfo = new FontInfo(cell.Font, cell.ForeColor);
            fontIndex = cell.Document.Fonts.IndexOf(fontInfo);
            if (fontIndex == -1)
            {
                cell.Document.Fonts.Add(fontInfo);
                fontIndex = cell.Document.Fonts.IndexOf(fontInfo);
            }

            if (fontIndex > 3)
                fontIndex++;

        }

        public override bool Equals(object obj)
        {
            if (obj is FormatInfo)
            {
                FormatInfo info = (FormatInfo)obj;
                return ((this.fontIndex == info.fontIndex) && (this.formatIndex == info.formatIndex)
                    && (this.ForeColor.Index == info.ForeColor.Index)  && (this.BackColor.Index == info.BackColor.Index) 
                    && (this.HorizontalAlignment == info.HorizontalAlignment) );

            }
            else
                return false;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public ExcelColor BackColor
        {
            get { return cell.BackColor; }
        }

        public ExcelColor ForeColor
        {
            get { return cell.ForeColor; }
        }

        public Font Font
        {
            get { return cell.Font; }
        }

        public string Format
        {
            get { return cell.Format; }
        }

        public Alignment HorizontalAlignment
        {
            get { return cell.Alignment; }
        }

        public int FormatIndex
        {
            get { return formatIndex; }
        }

        public int FontIndex
        {
            get { return fontIndex; }
        }
    }
}
