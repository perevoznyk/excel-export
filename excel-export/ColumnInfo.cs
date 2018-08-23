// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018

using System;
using System.Collections.Generic;
using System.Text;

namespace Export.XLS
{
    internal class ColumnInfo
    {
        public int Width { get; set; }
        public int Index { get; set; }

        public override bool Equals(object obj)
        {
            if (obj is ColumnInfo)
            {
                return (this.Index == ((ColumnInfo)obj).Index);
            }
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
