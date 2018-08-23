// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018


namespace Export.XLS
{
    public class Cell
    {
        private ExcelDocument document;
        private CellInfo cellInfo;

        internal Cell(int row, int column, ExcelDocument document)
        {
            this.document = document;
            cellInfo = document.GetCellInfo(row, column);
        }

        internal ExcelDocument Document
        {
            get { return this.document; }
        }

        public object Value
        {
            get { return cellInfo.Value; }
            set { cellInfo.Value = value; }
        }

        public string Format
        {
            get { return cellInfo.Format; }
            set
            {
                cellInfo.Format = value;
                if (!document.Formats.Contains(value))
                    document.Formats.Add(value);
            }
        }

        public ExcelColor BackColor
        {
            get { return cellInfo.BackColor; }
            set { cellInfo.BackColor = value; }
        }

        public ExcelColor ForeColor
        {
            get { return cellInfo.ForeColor; }
            set { cellInfo.ForeColor = value; }
        }

        public Font Font
        {
            get { return cellInfo.Font; }
            set { cellInfo.Font = value; }
        }

        public Alignment Alignment
        {
            get { return cellInfo.Alignment; }
            set { cellInfo.Alignment = value; }
        }

        public int Row
        {
            get { return cellInfo.Row; }
        }

        public int Column
        {
            get { return cellInfo.Column; }
        }
    }
}
