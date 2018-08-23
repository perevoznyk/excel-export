// C# Excel Writer library v2.1
// by Serhiy Perevoznyk, 2008-2018

using System;
using System.Collections.Generic;
using System.Text;

namespace Export.XLS
{
    [Flags]
    public enum FontStyle
    {
        Regular = 0,
        Bold = 1,
        Italic = 2,
        Underline = 4,
        Strikeout = 8
    }
}
