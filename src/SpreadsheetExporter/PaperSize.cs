using System;
using System.Collections.Generic;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The paper size.
    /// </summary>
    public class PaperSize : IEquatable<PaperSize>, IEquatable<int>, IComparable<PaperSize>, IComparable<int>, IComparable {
        private static readonly Lazy<PaperSize> letter = new(() => new PaperSize(1, "Letter", 850, 1100));
        private static readonly Lazy<PaperSize> letterSmall = new(() => new PaperSize(2, "Letter Small", 850, 1100));
        private static readonly Lazy<PaperSize> tabloid = new(() => new PaperSize(3, "Tabloid", 1100, 1700));
        private static readonly Lazy<PaperSize> ledger = new(() => new PaperSize(4, "Ledger", 1700, 1100));
        private static readonly Lazy<PaperSize> legal = new(() => new PaperSize(5, "Legal", 850, 1400));
        private static readonly Lazy<PaperSize> statement = new(() => new PaperSize(6, "Statement", 550, 850));
        private static readonly Lazy<PaperSize> executive = new(() => new PaperSize(7, "Executive", 725, 1050));
        private static readonly Lazy<PaperSize> a3 = new(() => new PaperSize(8, "A3", 1169, 1654));
        private static readonly Lazy<PaperSize> a4 = new(() => new PaperSize(9, "A4", 827, 1169));
        private static readonly Lazy<PaperSize> a4Small = new(() => new PaperSize(10, "A4 Small", 827, 1169));
        private static readonly Lazy<PaperSize> a5 = new(() => new PaperSize(11, "A5", 583, 827));
        private static readonly Lazy<PaperSize> b4 = new(() => new PaperSize(12, "B4 (JIS)", 1012, 1433));
        private static readonly Lazy<PaperSize> b5 = new(() => new PaperSize(13, "B5 (JIS)", 717, 1012));
        private static readonly Lazy<PaperSize> folio = new(() => new PaperSize(14, "Folio", 850, 1300));
        private static readonly Lazy<PaperSize> quarto = new(() => new PaperSize(15, "Quarto", 846, 1083));
        private static readonly Lazy<PaperSize> standard10x14 = new(() => new PaperSize(16, "10×14", 1000, 1400));
        private static readonly Lazy<PaperSize> standard11x17 = new(() => new PaperSize(17, "11×17", 1100, 1700));
        private static readonly Lazy<PaperSize> note = new(() => new PaperSize(18, "Note", 850, 1100));
        private static readonly Lazy<PaperSize> number9Envelope = new(() => new PaperSize(19, "Envelope #9", 387, 887));
        private static readonly Lazy<PaperSize> number10Envelope = new(() => new PaperSize(20, "Envelope #10", 412, 950));
        private static readonly Lazy<PaperSize> number11Envelope = new(() => new PaperSize(21, "Envelope #11", 450, 1037));
        private static readonly Lazy<PaperSize> number12Envelope = new(() => new PaperSize(22, "Envelope #12", 475, 1100));
        private static readonly Lazy<PaperSize> number14Envelope = new(() => new PaperSize(23, "Envelope #14", 500, 1150));
        private static readonly Lazy<PaperSize> cSheet = new(() => new PaperSize(24, "C size sheet", 1700, 2200));
        private static readonly Lazy<PaperSize> dSheet = new(() => new PaperSize(25, "D size sheet", 2200, 3400));
        private static readonly Lazy<PaperSize> eSheet = new(() => new PaperSize(26, "E size sheet", 3400, 4400));
        private static readonly Lazy<PaperSize> dLEnvelope = new(() => new PaperSize(27, "Envelope DL", 433, 866));
        private static readonly Lazy<PaperSize> c5Envelope = new(() => new PaperSize(28, "Envelope C5", 638, 902));
        private static readonly Lazy<PaperSize> c3Envelope = new(() => new PaperSize(29, "Envelope C3", 1276, 1803));
        private static readonly Lazy<PaperSize> c4Envelope = new(() => new PaperSize(30, "Envelope C4", 902, 1276));
        private static readonly Lazy<PaperSize> c6Envelope = new(() => new PaperSize(31, "Envelope C6", 449, 638));
        private static readonly Lazy<PaperSize> c65Envelope = new(() => new PaperSize(32, "Envelope C65", 449, 902));
        private static readonly Lazy<PaperSize> b4Envelope = new(() => new PaperSize(33, "Envelope B4", 984, 1390));
        private static readonly Lazy<PaperSize> b5Envelope = new(() => new PaperSize(34, "Envelope B5", 693, 984));
        private static readonly Lazy<PaperSize> b6Envelope = new(() => new PaperSize(35, "Envelope B6", 693, 492));
        private static readonly Lazy<PaperSize> italyEnvelope = new(() => new PaperSize(36, "Envelope", 433, 906));
        private static readonly Lazy<PaperSize> monarchEnvelope = new(() => new PaperSize(37, "Envelope Monarch", 387, 750));
        private static readonly Lazy<PaperSize> personalEnvelope = new(() => new PaperSize(38, "6 3/4 Envelope", 362, 650));
        private static readonly Lazy<PaperSize> uSStandardFanfold = new(() => new PaperSize(39, "US Std Fanfold", 1487, 1100));
        private static readonly Lazy<PaperSize> germanStandardFanfold = new(() => new PaperSize(40, "German Std Fanfold", 850, 1200));
        private static readonly Lazy<PaperSize> germanLegalFanfold = new(() => new PaperSize(41, "German Legal Fanfold", 850, 1300));
        private static readonly Lazy<PaperSize> isoB4 = new(() => new PaperSize(42, "B4 (ISO)", 984, 1390));
        private static readonly Lazy<PaperSize> japanesePostcard = new(() => new PaperSize(43, "Japanese Postcard", 394, 583));
        private static readonly Lazy<PaperSize> standard9x11 = new(() => new PaperSize(44, "9×11", 900, 1100));
        private static readonly Lazy<PaperSize> standard10x11 = new(() => new PaperSize(45, "10×11", 1000, 1100));
        private static readonly Lazy<PaperSize> standard15x11 = new(() => new PaperSize(46, "15×11", 1500, 1100));
        private static readonly Lazy<PaperSize> inviteEnvelope = new(() => new PaperSize(47, "Envelope Invite", 866, 866));
        private static readonly Lazy<PaperSize> letterExtra = new(() => new PaperSize(50, "Letter Extra", 950, 1200));
        private static readonly Lazy<PaperSize> legalExtra = new(() => new PaperSize(51, "Legal Extra", 950, 1500));
        private static readonly Lazy<PaperSize> a4Extra = new(() => new PaperSize(53, "A4 Extra", 927, 1269));
        private static readonly Lazy<PaperSize> letterTransverse = new(() => new PaperSize(54, "Letter Transverse", 850, 1100));
        private static readonly Lazy<PaperSize> a4Transverse = new(() => new PaperSize(55, "A4 Transverse", 827, 1169));
        private static readonly Lazy<PaperSize> letterExtraTransverse = new(() => new PaperSize(56, "Letter Extra Transverse", 950, 1200));
        private static readonly Lazy<PaperSize> aPlus = new(() => new PaperSize(57, "Super A", 894, 1402));
        private static readonly Lazy<PaperSize> bPlus = new(() => new PaperSize(58, "Super B", 1201, 1917));
        private static readonly Lazy<PaperSize> letterPlus = new(() => new PaperSize(59, "Letter Plus", 850, 1269));
        private static readonly Lazy<PaperSize> a4Plus = new(() => new PaperSize(60, "A4 Plus", 827, 1299));
        private static readonly Lazy<PaperSize> a5Transverse = new(() => new PaperSize(61, "A5 Transverse", 583, 827));
        private static readonly Lazy<PaperSize> b5Transverse = new(() => new PaperSize(62, "B5 (JIS) Transverse", 717, 1012));
        private static readonly Lazy<PaperSize> a3Extra = new(() => new PaperSize(63, "A3 Extra", 1268, 1752));
        private static readonly Lazy<PaperSize> a5Extra = new(() => new PaperSize(64, "A5 Extra", 685, 925));
        private static readonly Lazy<PaperSize> b5Extra = new(() => new PaperSize(65, "B5 (ISO) Extra", 791, 1087));
        private static readonly Lazy<PaperSize> a2 = new(() => new PaperSize(66, "A2", 1654, 2339));
        private static readonly Lazy<PaperSize> a3Transverse = new(() => new PaperSize(67, "A3 Transverse", 1169, 1654));
        private static readonly Lazy<PaperSize> a3ExtraTransverse = new(() => new PaperSize(68, "A3 Extra Transverse", 1268, 1752));

        private PaperSize(int value, string name, int width, int height) {
            Value = value;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Width = width;
            Height = height;
        }

        /// <summary>
        /// Performs an implicit conversion from <see cref="PaperSize" /> to <see cref="int" />.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The result of the conversion.</returns>
        public static implicit operator int(PaperSize value) {
            return value.Value;
        }

        /// <summary>
        /// Performs an explicit conversion from <see cref="int" /> to <see cref="PaperSize" />.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The result of the conversion.</returns>
        /// <exception cref="InvalidCastException"></exception>
        public static explicit operator PaperSize(int value) {
            return GetAll().SingleOrDefault(x => x.Value == value)
                ?? throw new InvalidCastException();
        }

        /// <summary>
        /// Implements the operator ==.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator ==(PaperSize left, PaperSize right) {
            return (left is null && right is null) || left?.Equals(right) == true;
        }

        /// <summary>
        /// Implements the operator ==.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator ==(PaperSize left, int right) {
            return left?.Value == right;
        }

        /// <summary>
        /// Implements the operator ==.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator ==(int left, PaperSize right) {
            return right?.Value == left;
        }

        /// <summary>
        /// Implements the operator !=.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator !=(PaperSize left, PaperSize right) {
            return !(left == right);
        }

        /// <summary>
        /// Implements the operator !=.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator !=(PaperSize left, int right) {
            return !(left == right);
        }

        /// <summary>
        /// Implements the operator !=.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>The result of the operator.</returns>
        public static bool operator !=(int left, PaperSize right) {
            return !(left == right);
        }

        /// <summary>
        /// Gets the letter paper (8.5 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 1
        /// </value>
        public static PaperSize Letter => letter.Value;

        /// <summary>
        /// Gets the letter small paper (8.5 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 2
        /// </value>
        public static PaperSize LetterSmall => letterSmall.Value;

        /// <summary>
        /// Gets the tabloid paper (11 in. by 17 in.).
        /// </summary>
        /// <value>
        /// 3
        /// </value>
        public static PaperSize Tabloid => tabloid.Value;

        /// <summary>
        /// Gets the ledger paper (17 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 4
        /// </value>
        public static PaperSize Ledger => ledger.Value;

        /// <summary>
        /// Gets the legal paper (8.5 in. by 14 in.).
        /// </summary>
        /// <value>
        /// 5
        /// </value>
        public static PaperSize Legal => legal.Value;

        /// <summary>
        /// Gets the statement paper (5.5 in. by 8.5 in.).
        /// </summary>
        /// <value>
        /// 6
        /// </value>
        public static PaperSize Statement => statement.Value;

        /// <summary>
        /// Gets the executive paper (7.25 in. by 10.5 in.).
        /// </summary>
        /// <value>
        /// 7
        /// </value>
        public static PaperSize Executive => executive.Value;

        /// <summary>
        /// Gets the A3 paper (297 mm by 420 mm).
        /// </summary>
        /// <value>
        /// 8
        /// </value>
        public static PaperSize A3 => a3.Value;

        /// <summary>
        /// Gets the A4 paper (210 mm by 297 mm).
        /// </summary>
        /// <value>
        /// 9
        /// </value>
        public static PaperSize A4 => a4.Value;

        /// <summary>
        /// Gets the A4 small paper (210 mm by 297 mm).
        /// </summary>
        /// <value>
        /// 10
        /// </value>
        public static PaperSize A4Small => a4Small.Value;

        /// <summary>
        /// Gets the A5 paper (148 mm by 210 mm).
        /// </summary>
        /// <value>
        /// 11
        /// </value>
        public static PaperSize A5 => a5.Value;

        /// <summary>
        /// Gets the B4 paper (250 mm by 353 mm).
        /// </summary>
        /// <value>
        /// 12
        /// </value>
        public static PaperSize B4 => b4.Value;

        /// <summary>
        /// Gets the B5 paper (176 mm by 250 mm).
        /// </summary>
        /// <value>
        /// 13
        /// </value>
        public static PaperSize B5 => b5.Value;

        /// <summary>
        /// Gets the folio paper (8.5 in. by 13 in.).
        /// </summary>
        /// <value>
        /// 14
        /// </value>
        public static PaperSize Folio => folio.Value;

        /// <summary>
        /// Gets the quarto paper (215 mm by 275 mm).
        /// </summary>
        /// <value>
        /// 15
        /// </value>
        public static PaperSize Quarto => quarto.Value;

        /// <summary>
        /// Gets the standard paper (10 in. by 14 in.)
        /// </summary>
        /// <value>
        /// 16
        /// </value>
        public static PaperSize Standard10x14 => standard10x14.Value;

        /// <summary>
        /// Gets the standard paper (11 in. by 17 in.).
        /// </summary>
        /// <value>
        /// 17
        /// </value>
        public static PaperSize Standard11x17 => standard11x17.Value;

        /// <summary>
        /// Gets the note paper (8.5 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 18
        /// </value>
        public static PaperSize Note => note.Value;

        /// <summary>
        /// Gets the #9 envelope (3.875 in. by 8.875 in.).
        /// </summary>
        /// <value>
        /// 19
        /// </value>
        public static PaperSize Number9Envelope => number9Envelope.Value;

        /// <summary>
        /// Gets the #10 envelope (4.125 in. by 9.5 in.).
        /// </summary>
        /// <value>
        /// 20
        /// </value>
        public static PaperSize Number10Envelope => number10Envelope.Value;

        /// <summary>
        /// Gets the #11 envelope (4.5 in. by 10.375 in.).
        /// </summary>
        /// <value>
        /// 21
        /// </value>
        public static PaperSize Number11Envelope => number11Envelope.Value;

        /// <summary>
        /// Gets the #12 envelope (4.75 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 22
        /// </value>
        public static PaperSize Number12Envelope => number12Envelope.Value;

        /// <summary>
        /// Gets the #14 envelope (5 in. by 11.5 in.).
        /// </summary>
        /// <value>
        /// 23
        /// </value>
        public static PaperSize Number14Envelope => number14Envelope.Value;

        /// <summary>
        /// Gets the C sheet paper (17 in. by 22 in.).
        /// </summary>
        /// <value>
        /// 24
        /// </value>
        public static PaperSize CSheet => cSheet.Value;

        /// <summary>
        /// Gets the D sheet paper (22 in. by 34 in.).
        /// </summary>
        /// <value>
        /// 25
        /// </value>
        public static PaperSize DSheet => dSheet.Value;

        /// <summary>
        /// Gets the E sheet paper (34 in. by 44 in.).
        /// </summary>
        /// <value>
        /// 26
        /// </value>
        public static PaperSize ESheet => eSheet.Value;

        /// <summary>
        /// Gets the DL envelope (110 mm by 220 mm).
        /// </summary>
        /// <value>
        /// 27
        /// </value>
        public static PaperSize DLEnvelope => dLEnvelope.Value;

        /// <summary>
        /// Gets the C5 envelope (162 mm by 229 mm).
        /// </summary>
        /// <value>
        /// 28
        /// </value>
        public static PaperSize C5Envelope => c5Envelope.Value;

        /// <summary>
        /// Gets the C3 envelope (324 mm by 458 mm).
        /// </summary>
        /// <value>
        /// 29
        /// </value>
        public static PaperSize C3Envelope => c3Envelope.Value;

        /// <summary>
        /// Gets the C4 envelope (229 mm by 324 mm).
        /// </summary>
        /// <value>
        /// 30
        /// </value>
        public static PaperSize C4Envelope => c4Envelope.Value;

        /// <summary>
        /// Gets the C6 envelope (114 mm by 162 mm).
        /// </summary>
        /// <value>
        /// 31
        /// </value>
        public static PaperSize C6Envelope => c6Envelope.Value;

        /// <summary>
        /// Gets the C65 envelope (114 mm by 229 mm).
        /// </summary>
        /// <value>
        /// 32
        /// </value>
        public static PaperSize C65Envelope => c65Envelope.Value;

        /// <summary>
        /// Gets the B4 envelope (250 mm by 353 mm).
        /// </summary>
        /// <value>
        /// 33.
        /// </value>
        public static PaperSize B4Envelope => b4Envelope.Value;

        /// <summary>
        /// Gets the B5 envelope (176 mm by 250 mm).
        /// </summary>
        /// <value>
        /// 34
        /// </value>
        public static PaperSize B5Envelope => b5Envelope.Value;

        /// <summary>
        /// Gets the B6 envelope (176 mm by 125 mm).
        /// </summary>
        /// <value>
        /// 35
        /// </value>
        public static PaperSize B6Envelope => b6Envelope.Value;

        /// <summary>
        /// Gets the italy envelope (110 mm by 230 mm).
        /// </summary>
        /// <value>
        /// 36
        /// </value>
        public static PaperSize ItalyEnvelope => italyEnvelope.Value;

        /// <summary>
        /// Gets the monarch envelope (3.875 in. by 7.5 in.).
        /// </summary>
        /// <value>
        /// 37
        /// </value>
        public static PaperSize MonarchEnvelope => monarchEnvelope.Value;

        /// <summary>
        /// Gets the personal (6 3/4) envelope (3.625 in. by 6.5 in.).
        /// </summary>
        /// <value>
        /// 38
        /// </value>
        public static PaperSize PersonalEnvelope => personalEnvelope.Value;

        /// <summary>
        /// Gets the us standard fanfold (14.875 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 39
        /// </value>
        public static PaperSize USStandardFanfold => uSStandardFanfold.Value;

        /// <summary>
        /// Gets the german standard fanfold (8.5 in. by 12 in.).
        /// </summary>
        /// <value>
        /// 40
        /// </value>
        public static PaperSize GermanStandardFanfold => germanStandardFanfold.Value;

        /// <summary>
        /// Gets the german legal fanfold (8.5 in. by 13 in.).
        /// </summary>
        /// <value>
        /// 41
        /// </value>
        public static PaperSize GermanLegalFanfold => germanLegalFanfold.Value;

        /// <summary>
        /// Gets the ISO B4 (250 mm by 353 mm).
        /// </summary>
        /// <value>
        /// 42
        /// </value>
        public static PaperSize IsoB4 => isoB4.Value;

        /// <summary>
        /// Gets the japanese postcard (200 mm by 148 mm).
        /// </summary>
        /// <value>
        /// 43
        /// </value>
        public static PaperSize JapanesePostcard => japanesePostcard.Value;

        /// <summary>
        /// Gets the standard paper (9 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 44
        /// </value>
        public static PaperSize Standard9x11 => standard9x11.Value;

        /// <summary>
        /// Gets the standard paper (10 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 45
        /// </value>
        public static PaperSize Standard10x11 => standard10x11.Value;

        /// <summary>
        /// Gets the standard paper (15 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 46
        /// </value>
        public static PaperSize Standard15x11 => standard15x11.Value;

        /// <summary>
        /// Gets the invite envelope (220 mm by 220 mm).
        /// </summary>
        /// <value>
        /// 47
        /// </value>
        public static PaperSize InviteEnvelope => inviteEnvelope.Value;

        /// <summary>
        /// Gets the letter extra paper (9.275 in. by 12 in.).
        /// </summary>
        /// <value>
        /// 50
        /// </value>
        public static PaperSize LetterExtra => letterExtra.Value;

        /// <summary>
        /// Gets the legal extra paper (9.275 in. by 15 in.).
        /// </summary>
        /// <value>
        /// 51
        /// </value>
        public static PaperSize LegalExtra => legalExtra.Value;

        /// <summary>
        /// Gets the a4 extra paper (236 mm by 322 mm).
        /// </summary>
        /// <value>
        /// 53
        /// </value>
        public static PaperSize A4Extra => a4Extra.Value;

        /// <summary>
        /// Gets the letter transverse (8.275 in. by 11 in.).
        /// </summary>
        /// <value>
        /// 54
        /// </value>
        public static PaperSize LetterTransverse => letterTransverse.Value;

        /// <summary>
        /// Gets the a4 transverse (210 mm by 297 mm).
        /// </summary>
        /// <value>
        /// 55
        /// </value>
        public static PaperSize A4Transverse => a4Transverse.Value;

        /// <summary>
        /// Gets the letter extra transverse (9.275 in. by 12 in.).
        /// </summary>
        /// <value>
        /// 56
        /// </value>
        public static PaperSize LetterExtraTransverse => letterExtraTransverse.Value;

        /// <summary>
        /// Gets a plus paper (227 mm by 356 mm).
        /// </summary>
        /// <value>
        /// 57
        /// </value>
        public static PaperSize APlus => aPlus.Value;

        /// <summary>
        /// Gets the b plus paper (305 mm by 487 mm).
        /// </summary>
        /// <value>
        /// 58
        /// </value>
        public static PaperSize BPlus => bPlus.Value;

        /// <summary>
        /// Gets the letter plus paper (8.5 in. by 12.69 in.).
        /// </summary>
        /// <value>
        /// 59
        /// </value>
        public static PaperSize LetterPlus => letterPlus.Value;

        /// <summary>
        /// Gets the A4 plus paper (210 mm by 330 mm).
        /// </summary>
        /// <value>
        /// 60
        /// </value>
        public static PaperSize A4Plus => a4Plus.Value;

        /// <summary>
        /// Gets the A5 transverse paper (148 mm by 210 mm).
        /// </summary>
        /// <value>
        /// 61
        /// </value>
        public static PaperSize A5Transverse => a5Transverse.Value;

        /// <summary>
        /// Gets the JIS B5 transverse paper (182 mm by 257 mm).
        /// </summary>
        /// <value>
        /// 62
        /// </value>
        public static PaperSize B5Transverse => b5Transverse.Value;

        /// <summary>
        /// Gets the A3 extra paper (322 mm by 445 mm).
        /// </summary>
        /// <value>
        /// 63
        /// </value>
        public static PaperSize A3Extra => a3Extra.Value;

        /// <summary>
        /// Gets the A5 extra paper (174 mm by 235 mm).
        /// </summary>
        /// <value>
        /// 64
        /// </value>
        public static PaperSize A5Extra => a5Extra.Value;

        /// <summary>
        /// Gets the ISO B5 extra paper (201 mm by 276 mm).
        /// </summary>
        /// <value>
        /// 65
        /// </value>
        public static PaperSize B5Extra => b5Extra.Value;

        /// <summary>
        /// Gets the A2 paper (420 mm by 594 mm).
        /// </summary>
        /// <value>
        /// 66
        /// </value>
        public static PaperSize A2 => a2.Value;

        /// <summary>
        /// Gets the A3 transverse paper (297 mm by 420 mm).
        /// </summary>
        /// <value>
        /// 67
        /// </value>
        public static PaperSize A3Transverse => a3Transverse.Value;

        /// <summary>
        /// Gets the A3 extra transverse paper (322 mm by 445 mm).
        /// </summary>
        /// <value>
        /// 68
        /// </value>
        public static PaperSize A3ExtraTransverse => a3ExtraTransverse.Value;

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public int Value { get; }

        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The display name.
        /// </value>
        public string Name { get; }

        /// <summary>
        /// Gets the width.
        /// </summary>
        /// <value>
        /// The width.
        /// </value>
        public int Width { get; }

        /// <summary>
        /// Gets the height.
        /// </summary>
        /// <value>
        /// The height.
        /// </value>
        public int Height { get; }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <returns>The all paper sizes.</returns>
        public static IEnumerable<PaperSize> GetAll() {
            yield return Letter;
            yield return LetterSmall;
            yield return Tabloid;
            yield return Ledger;
            yield return Legal;
            yield return Statement;
            yield return Executive;
            yield return A3;
            yield return A4;
            yield return A4Small;
            yield return A5;
            yield return B4;
            yield return B5;
            yield return Folio;
            yield return Quarto;
            yield return Standard10x14;
            yield return Standard11x17;
            yield return Note;
            yield return Number9Envelope;
            yield return Number10Envelope;
            yield return Number11Envelope;
            yield return Number12Envelope;
            yield return Number14Envelope;
            yield return CSheet;
            yield return DSheet;
            yield return ESheet;
            yield return DLEnvelope;
            yield return C5Envelope;
            yield return C3Envelope;
            yield return C4Envelope;
            yield return C6Envelope;
            yield return C65Envelope;
            yield return B4Envelope;
            yield return B5Envelope;
            yield return B6Envelope;
            yield return ItalyEnvelope;
            yield return MonarchEnvelope;
            yield return PersonalEnvelope;
            yield return USStandardFanfold;
            yield return GermanStandardFanfold;
            yield return GermanLegalFanfold;
            yield return IsoB4;
            yield return JapanesePostcard;
            yield return Standard9x11;
            yield return Standard10x11;
            yield return Standard15x11;
            yield return InviteEnvelope;
            yield return LetterExtra;
            yield return LegalExtra;
            yield return A4Extra;
            yield return LetterTransverse;
            yield return A4Transverse;
            yield return LetterExtraTransverse;
            yield return APlus;
            yield return BPlus;
            yield return LetterPlus;
            yield return A4Plus;
            yield return A5Transverse;
            yield return B5Transverse;
            yield return A3Extra;
            yield return A5Extra;
            yield return B5Extra;
            yield return A2;
            yield return A3Transverse;
            yield return A3ExtraTransverse;
        }

        /// <inheritdoc/>
        public override bool Equals(object obj) {
            if (obj is PaperSize paperSize) {
                return Equals(paperSize);
            } else if (obj is int i) {
                return Equals(i);
            }

            return base.Equals(obj);
        }

        /// <inheritdoc/>
        public bool Equals(PaperSize other) {
            return Value == other?.Value;
        }

        /// <inheritdoc/>
        public bool Equals(int other) {
            return Value == other;
        }

        /// <inheritdoc/>
        public int CompareTo(object obj) {
            return obj is int i ? Value.CompareTo(i) : Value.CompareTo(((PaperSize)obj).Value);
        }

        /// <inheritdoc/>
        public int CompareTo(PaperSize other) {
            return Value.CompareTo(other.Value);
        }

        /// <inheritdoc/>
        public int CompareTo(int other) {
            return Value.CompareTo(other);
        }

        /// <inheritdoc/>
        public override int GetHashCode() {
            return Value.GetHashCode();
        }

        /// <summary>
        /// Converts to string.
        /// </summary>
        /// <returns>The name.</returns>
        public override string ToString() {
            return Name;
        }
    }
}
