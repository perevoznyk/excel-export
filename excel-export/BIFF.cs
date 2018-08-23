// C# Excel Writer library v2.0
// by Serhiy Perevoznyk, 2008-2018
// .NET Core


namespace Export.XLS
{
    internal sealed class BIFF
    {
        public const ushort DefaultColor = 0x7fff;

        public const ushort BOFRecord = 0x0209;
        public const ushort EOFRecord = 0x0A;

        public const ushort FontRecord = 0x0231;
        public const ushort FormatRecord = 0x001E;
        public const ushort LabelRecord = 0x0204;
        public const ushort WindowProtectRecord = 0x0019;
        public const ushort XFRecord = 0x0243;
        public const ushort HeaderRecord = 0x0014;
        public const ushort FooterRecord = 0x0015;
        public const ushort ExtendedRecord = 0x0243;
        public const ushort StyleRecord = 0x0293;
        public const ushort CodepageRecord = 0x0042;
        public const ushort NumberRecord = 0x0203;
        public const ushort ColumnInfoRecord = 0x007D;

    }
}
