using System;
using System.Drawing;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Primitives;

namespace CloudyWing.SpreadsheetExporter.Config {
    public sealed class SpreadsheetConfiguration {
        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetConfiguration"/> class.
        /// <para> CellFont: 來自Config的預設字型樣式</para>
        /// <para> CellStyle: 來自Config的預設樣式</para>
        /// <para> ListHeaderFont: 預設格式、粗體</para>
        /// <para> ListHeaderStyle: 預設格式、粗體、置中</para>
        /// <para> ListTextStyle: 預設格式、置左</para>
        /// <para> ListNumberStyle: 預設格式、置右</para>
        /// <para> ListDateTimeStyle: 預設格式、置右</para>
        /// </summary>
        public SpreadsheetConfiguration(string basicPath, string jsonFile = "Spreadsheet.json") {
            IConfigurationRoot config = new ConfigurationBuilder()
                .SetBasePath(basicPath)
                .AddJsonFile(jsonFile, optional: false, reloadOnChange: true)
                .Build();

            ChangeToken.OnChange(() => config.GetReloadToken(), () => {
                Initialize();
            });

            Initialize();

            void Initialize() {
                ConfigSettings configSettings = config.GetSection("Spreadsheet").Get<ConfigSettings>();

                CellSettings cellSettings = configSettings.Cell;

                FontStyles style = FontStyles.None;
                if (cellSettings.Font.IsBold) {
                    style |= FontStyles.IsBold;
                }

                if (cellSettings.Font.IsItalic) {
                    style |= FontStyles.IsItalic;
                }

                if (cellSettings.Font.HasUnderline) {
                    style |= FontStyles.HasUnderline;
                }

                if (cellSettings.Font.IsStrikeout) {
                    style |= FontStyles.IsStrikeout;
                }

                CellStyle = new CellStyle(
                    cellSettings.HorizontalAlignment,
                    cellSettings.VerticalAlignment,
                    cellSettings.HasBorder,
                    cellSettings.WrapText,
                    Color.Empty,
                    new CellFont(
                        cellSettings.Font.FontName,
                        cellSettings.Font.FontSize,
                        Color.Black,
                        style
                    )
                );

                CellFont listHeaderFont = CellStyle.Font
                    .CloneAndSetStyle(CellStyle.Font.Style | FontStyles.IsBold);

                ListHeaderStyle = CellStyle
                    .CloneAndSetFont(listHeaderFont)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Center)
                    .CloneAndSetBorder(true)
                    .CloneAndSetFont(listHeaderFont);

                ListTextStyle = CellStyle
                    .CloneAndSetBorder(true)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);

                ListNumberStyle = CellStyle
                    .CloneAndSetBorder(true)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);

                ListDateTimeStyle = CellStyle
                    .CloneAndSetBorder(true)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);
            }
        }

        public SpreadsheetConfiguration(Action<SetupBuilder> setuper) {
            SetupBuilder setupBuilder = new SetupBuilder(this);
            setuper(setupBuilder);
        }

        public CellStyle CellStyle { get; internal set; }


        /// <summary>
        /// ListTemplate的預設標題樣式
        /// </summary>
        public CellStyle ListHeaderStyle { get; internal set; }

        /// <summary>
        /// 預設格式、置左
        /// </summary>
        public CellStyle ListTextStyle { get; internal set; }

        /// <summary>
        /// 預設格式、置右
        /// </summary>
        public CellStyle ListNumberStyle { get; internal set; }

        /// <summary>
        /// 預設格式、置右
        /// </summary>
        public CellStyle ListDateTimeStyle { get; internal set; }
    }
}
