using CloudyWing.Enumeration.Abstractions;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a standard paper size with its numeric identifier and dimensions.
/// </summary>
public sealed class PaperSize : IntEnumeration<PaperSize> {
    /// <summary>
    /// Initializes a new instance of the <see cref="PaperSize"/> class with a value only.
    /// </summary>
    /// <param name="value">The numeric paper size identifier.</param>
    public PaperSize(int value) : base(value) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="PaperSize"/> class with a value and display name.
    /// </summary>
    /// <param name="value">The numeric paper size identifier.</param>
    /// <param name="name">The display name.</param>
    public PaperSize(int value, string name) : base(value, name) { }

    private PaperSize(int value, string name, int width, int height) : base(value, name) {
        Width = width;
        Height = height;
    }

    /// <summary>
    /// Gets the width in hundredths of an inch.
    /// </summary>
    public int Width { get; }

    /// <summary>
    /// Gets the height in hundredths of an inch.
    /// </summary>
    public int Height { get; }

    /// <summary>
    /// Gets the Letter paper size (8.5 in. by 11 in.). Value: 1.
    /// </summary>
    public static PaperSize Letter { get; } = new(1, "Letter", 850, 1100);

    /// <summary>
    /// Gets the Letter Small paper size (8.5 in. by 11 in.). Value: 2.
    /// </summary>
    public static PaperSize LetterSmall { get; } = new(2, "Letter Small", 850, 1100);

    /// <summary>
    /// Gets the Tabloid paper (11 in. by 17 in.). Value: 3.
    /// </summary>
    public static PaperSize Tabloid { get; } = new(3, "Tabloid", 1100, 1700);

    /// <summary>
    /// Gets the Ledger paper (17 in. by 11 in.). Value: 4.
    /// </summary>
    public static PaperSize Ledger { get; } = new(4, "Ledger", 1700, 1100);

    /// <summary>
    /// Gets the Legal paper (8.5 in. by 14 in.). Value: 5
    /// </summary>
    public static PaperSize Legal { get; } = new(5, "Legal", 850, 1400);

    /// <summary>
    /// Gets the Statement paper (5.5 in. by 8.5 in.). Value: 6
    /// </summary>
    public static PaperSize Statement { get; } = new(6, "Statement", 550, 850);

    /// <summary>
    /// Gets the Executive paper (7.25 in. by 10.5 in.). Value: 7
    /// </summary>
    public static PaperSize Executive { get; } = new(7, "Executive", 725, 1050);

    /// <summary>
    /// Gets the A3 paper (297 mm by 420 mm). Value: 8
    /// </summary>
    public static PaperSize A3 { get; } = new(8, "A3", 1169, 1654);

    /// <summary>
    /// Gets the A4 paper (210 mm by 297 mm). Value: 9
    /// </summary>
    public static PaperSize A4 { get; } = new(9, "A4", 827, 1169);

    /// <summary>
    /// Gets the A4 Small paper (210 mm by 297 mm). Value: 10
    /// </summary>
    public static PaperSize A4Small { get; } = new(10, "A4 Small", 827, 1169);

    /// <summary>
    /// Gets the A5 paper (148 mm by 210 mm). Value: 11
    /// </summary>
    public static PaperSize A5 { get; } = new(11, "A5", 583, 827);

    /// <summary>
    /// Gets the B4 (JIS) paper (250 mm by 353 mm). Value: 12
    /// </summary>
    public static PaperSize B4 { get; } = new(12, "B4 (JIS)", 1012, 1433);

    /// <summary>
    /// Gets the B5 (JIS) paper (176 mm by 250 mm). Value: 13
    /// </summary>
    public static PaperSize B5 { get; } = new(13, "B5 (JIS)", 717, 1012);

    /// <summary>
    /// Gets the Folio paper (8.5 in. by 13 in.). Value: 14
    /// </summary>
    public static PaperSize Folio { get; } = new(14, "Folio", 850, 1300);

    /// <summary>
    /// Gets the Quarto paper (215 mm by 275 mm). Value: 15
    /// </summary>
    public static PaperSize Quarto { get; } = new(15, "Quarto", 846, 1083);

    /// <summary>
    /// Gets the standard 10×14 paper. Value: 16
    /// </summary>
    public static PaperSize Standard10x14 { get; } = new(16, "10×14", 1000, 1400);

    /// <summary>
    /// Gets the standard 11×17 paper. Value: 17
    /// </summary>
    public static PaperSize Standard11x17 { get; } = new(17, "11×17", 1100, 1700);

    /// <summary>
    /// Gets the Note paper (8.5 in. by 11 in.). Value: 18
    /// </summary>
    public static PaperSize Note { get; } = new(18, "Note", 850, 1100);

    /// <summary>
    /// Gets Envelope #9 (3.875 in. by 8.875 in.). Value: 19
    /// </summary>
    public static PaperSize Number9Envelope { get; } = new(19, "Envelope #9", 387, 887);

    /// <summary>
    /// Gets Envelope #10 (4.125 in. by 9.5 in.). Value: 20
    /// </summary>
    public static PaperSize Number10Envelope { get; } = new(20, "Envelope #10", 412, 950);

    /// <summary>
    /// Gets Envelope #11 (4.5 in. by 10.375 in.). Value: 21
    /// </summary>
    public static PaperSize Number11Envelope { get; } = new(21, "Envelope #11", 450, 1037);

    /// <summary>
    /// Gets Envelope #12 (4.75 in. by 11 in.). Value: 22
    /// </summary>
    public static PaperSize Number12Envelope { get; } = new(22, "Envelope #12", 475, 1100);

    /// <summary>
    /// Gets Envelope #14 (5 in. by 11.5 in.). Value: 23
    /// </summary>
    public static PaperSize Number14Envelope { get; } = new(23, "Envelope #14", 500, 1150);

    /// <summary>
    /// Gets C sheet paper (17 in. by 22 in.). Value: 24
    /// </summary>
    public static PaperSize CSheet { get; } = new(24, "C size sheet", 1700, 2200);

    /// <summary>
    /// Gets D sheet paper (22 in. by 34 in.). Value: 25
    /// </summary>
    public static PaperSize DSheet { get; } = new(25, "D size sheet", 2200, 3400);

    /// <summary>
    /// Gets E sheet paper (34 in. by 44 in.). Value: 26
    /// </summary>
    public static PaperSize ESheet { get; } = new(26, "E size sheet", 3400, 4400);

    /// <summary>
    /// Gets DL Envelope (110 mm by 220 mm). Value: 27
    /// </summary>
    public static PaperSize DLEnvelope { get; } = new(27, "Envelope DL", 433, 866);

    /// <summary>
    /// Gets C5 Envelope (162 mm by 229 mm). Value: 28
    /// </summary>
    public static PaperSize C5Envelope { get; } = new(28, "Envelope C5", 638, 902);

    /// <summary>
    /// Gets C3 Envelope (324 mm by 458 mm). Value: 29
    /// </summary>
    public static PaperSize C3Envelope { get; } = new(29, "Envelope C3", 1276, 1803);

    /// <summary>
    /// Gets C4 Envelope (229 mm by 324 mm). Value: 30
    /// </summary>
    public static PaperSize C4Envelope { get; } = new(30, "Envelope C4", 902, 1276);

    /// <summary>
    /// Gets C6 Envelope (114 mm by 162 mm). Value: 31
    /// </summary>
    public static PaperSize C6Envelope { get; } = new(31, "Envelope C6", 449, 638);

    /// <summary>
    /// Gets C65 Envelope (114 mm by 229 mm). Value: 32
    /// </summary>
    public static PaperSize C65Envelope { get; } = new(32, "Envelope C65", 449, 902);

    /// <summary>
    /// Gets B4 Envelope (250 mm by 353 mm). Value: 33
    /// </summary>
    public static PaperSize B4Envelope { get; } = new(33, "Envelope B4", 984, 1390);

    /// <summary>
    /// Gets B5 Envelope (176 mm by 250 mm). Value: 34
    /// </summary>
    public static PaperSize B5Envelope { get; } = new(34, "Envelope B5", 693, 984);

    /// <summary>
    /// Gets B6 Envelope (176 mm by 125 mm). Value: 35
    /// </summary>
    public static PaperSize B6Envelope { get; } = new(35, "Envelope B6", 693, 492);

    /// <summary>
    /// Gets Italy Envelope (110 mm by 230 mm). Value: 36
    /// </summary>
    public static PaperSize ItalyEnvelope { get; } = new(36, "Envelope", 433, 906);

    /// <summary>
    /// Gets Monarch Envelope (3.875 in. by 7.5 in.). Value: 37
    /// </summary>
    public static PaperSize MonarchEnvelope { get; } = new(37, "Envelope Monarch", 387, 750);

    /// <summary>
    /// Gets 6¾ Envelope (3.625 in. by 6.5 in.). Value: 38
    /// </summary>
    public static PaperSize PersonalEnvelope { get; } = new(38, "6 3/4 Envelope", 362, 650);

    /// <summary>
    /// Gets US Standard Fanfold (14.875 in. by 11 in.). Value: 39
    /// </summary>
    public static PaperSize USStandardFanfold { get; } = new(39, "US Std Fanfold", 1487, 1100);

    /// <summary>
    /// Gets German Standard Fanfold (8.5 in. by 12 in.). Value: 40
    /// </summary>
    public static PaperSize GermanStandardFanfold { get; } = new(40, "German Std Fanfold", 850, 1200);

    /// <summary>
    /// Gets German Legal Fanfold (8.5 in. by 13 in.). Value: 41
    /// </summary>
    public static PaperSize GermanLegalFanfold { get; } = new(41, "German Legal Fanfold", 850, 1300);

    /// <summary>
    /// Gets ISO B4 (250 mm by 353 mm). Value: 42
    /// </summary>
    public static PaperSize IsoB4 { get; } = new(42, "B4 (ISO)", 984, 1390);

    /// <summary>
    /// Gets Japanese Postcard (200 mm by 148 mm). Value: 43
    /// </summary>
    public static PaperSize JapanesePostcard { get; } = new(43, "Japanese Postcard", 394, 583);

    /// <summary>
    /// Gets standard 9×11 paper. Value: 44
    /// </summary>
    public static PaperSize Standard9x11 { get; } = new(44, "9×11", 900, 1100);

    /// <summary>
    /// Gets standard 10×11 paper. Value: 45
    /// </summary>
    public static PaperSize Standard10x11 { get; } = new(45, "10×11", 1000, 1100);

    /// <summary>
    /// Gets standard 15×11 paper. Value: 46
    /// </summary>
    public static PaperSize Standard15x11 { get; } = new(46, "15×11", 1500, 1100);

    /// <summary>
    /// Gets Invite Envelope (220 mm by 220 mm). Value: 47
    /// </summary>
    public static PaperSize InviteEnvelope { get; } = new(47, "Envelope Invite", 866, 866);

    /// <summary>
    /// Gets Letter Extra paper (9.275 in. by 12 in.). Value: 50
    /// </summary>
    public static PaperSize LetterExtra { get; } = new(50, "Letter Extra", 950, 1200);

    /// <summary>
    /// Gets Legal Extra paper (9.275 in. by 15 in.). Value: 51
    /// </summary>
    public static PaperSize LegalExtra { get; } = new(51, "Legal Extra", 950, 1500);

    /// <summary>
    /// Gets A4 Extra paper (236 mm by 322 mm). Value: 53
    /// </summary>
    public static PaperSize A4Extra { get; } = new(53, "A4 Extra", 927, 1269);

    /// <summary>
    /// Gets Letter Transverse paper (8.275 in. by 11 in.). Value: 54
    /// </summary>
    public static PaperSize LetterTransverse { get; } = new(54, "Letter Transverse", 850, 1100);

    /// <summary>
    /// Gets A4 Transverse paper (210 mm by 297 mm). Value: 55
    /// </summary>
    public static PaperSize A4Transverse { get; } = new(55, "A4 Transverse", 827, 1169);

    /// <summary>
    /// Gets Letter Extra Transverse paper (9.275 in. by 12 in.). Value: 56
    /// </summary>
    public static PaperSize LetterExtraTransverse { get; } = new(56, "Letter Extra Transverse", 950, 1200);

    /// <summary>
    /// Gets Super A paper (227 mm by 356 mm). Value: 57
    /// </summary>
    public static PaperSize APlus { get; } = new(57, "Super A", 894, 1402);

    /// <summary>
    /// Gets Super B paper (305 mm by 487 mm). Value: 58
    /// </summary>
    public static PaperSize BPlus { get; } = new(58, "Super B", 1201, 1917);

    /// <summary>
    /// Gets Letter Plus paper (8.5 in. by 12.69 in.). Value: 59
    /// </summary>
    public static PaperSize LetterPlus { get; } = new(59, "Letter Plus", 850, 1269);

    /// <summary>
    /// Gets A4 Plus paper (210 mm by 330 mm). Value: 60
    /// </summary>
    public static PaperSize A4Plus { get; } = new(60, "A4 Plus", 827, 1299);

    /// <summary>
    /// Gets A5 Transverse paper (148 mm by 210 mm). Value: 61
    /// </summary>
    public static PaperSize A5Transverse { get; } = new(61, "A5 Transverse", 583, 827);

    /// <summary>
    /// Gets B5 (JIS) Transverse paper (182 mm by 257 mm). Value: 62
    /// </summary>
    public static PaperSize B5Transverse { get; } = new(62, "B5 (JIS) Transverse", 717, 1012);

    /// <summary>
    /// Gets A3 Extra paper (322 mm by 445 mm). Value: 63
    /// </summary>
    public static PaperSize A3Extra { get; } = new(63, "A3 Extra", 1268, 1752);

    /// <summary>
    /// Gets A5 Extra paper (174 mm by 235 mm). Value: 64
    /// </summary>
    public static PaperSize A5Extra { get; } = new(64, "A5 Extra", 685, 925);

    /// <summary>
    /// Gets ISO B5 Extra paper (201 mm by 276 mm). Value: 65
    /// </summary>
    public static PaperSize B5Extra { get; } = new(65, "B5 (ISO) Extra", 791, 1087);

    /// <summary>
    /// Gets A2 paper (420 mm by 594 mm). Value: 66
    /// </summary>
    public static PaperSize A2 { get; } = new(66, "A2", 1654, 2339);

    /// <summary>
    /// Gets A3 Transverse paper (297 mm by 420 mm). Value: 67
    /// </summary>
    public static PaperSize A3Transverse { get; } = new(67, "A3 Transverse", 1169, 1654);

    /// <summary>
    /// Gets A3 Extra Transverse paper (322 mm by 445 mm). Value: 68
    /// </summary>
    public static PaperSize A3ExtraTransverse { get; } = new(68, "A3 Extra Transverse", 1268, 1752);
}
