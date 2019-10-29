using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Tyreshop.DbAccess;
using Tyreshop.Utils;
using System.Text.RegularExpressions;

namespace Tyreshop.Utils
{
    class XlsxRutine
    {
        //Метод убирает из строки запрещенные спец символы.
        //Если не использовать, то при наличии в строке таких символов, вылетит ошибка.
        public static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            if (txt != null)
                return Regex.Replace(txt, r, "", RegexOptions.Compiled);
            else
                return "---";
        }
        public static string GetCellReference(int colIndex, uint rowInd)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter + rowInd.ToString();
        }
        /*Метод добавления текста в sharedStringTable*/
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        /*
         Метод добавления ячейки в строку
         на входе: строка, столбец, значение, тип значения, необязательный параметр - стиль
         */
        public static int InsertCellToRow(Row row, int cellNum, string val, SharedStringTablePart sstp, uint style = 0U)
        {
            if (val != null)
            {
                string cellRef = GetCellReference(cellNum, row.RowIndex);
                Cell refCell = null;
                Cell newCell = new Cell() { CellReference = cellRef, StyleIndex = style };
                row.InsertBefore(newCell, refCell);
                decimal tmp = 0;
                int tmp2 = 0;
                if (decimal.TryParse(val.Replace(',', '.'), out tmp))
                {
                    newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    newCell.CellValue = new CellValue(val);
                }
                else if (int.TryParse(val.Replace(',', '.'), out tmp2))
                {
                    newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    newCell.CellValue = new CellValue(val);
                }
                else
                {
                    newCell.CellValue = new CellValue(InsertSharedStringItem(val, sstp).ToString());
                    newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }

                cellNum++;
            }
            return cellNum;
        }
        public static void FilePropertiesGeneration(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Листы";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Лист 1";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;

        }

        public static Stylesheet GenerateStyleSheetForExport2()
        {
            Stylesheet stylesheet1 = new Stylesheet();

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)8U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)5U, FormatCode = "#,##0\\ \"₽\";\\-#,##0\\ \"₽\"" };
            NumberingFormat numberingFormat2 = new NumberingFormat() { NumberFormatId = (UInt32Value)6U, FormatCode = "#,##0\\ \"₽\";[Red]\\-#,##0\\ \"₽\"" };
            NumberingFormat numberingFormat3 = new NumberingFormat() { NumberFormatId = (UInt32Value)7U, FormatCode = "#,##0.00\\ \"₽\";\\-#,##0.00\\ \"₽\"" };
            NumberingFormat numberingFormat4 = new NumberingFormat() { NumberFormatId = (UInt32Value)8U, FormatCode = "#,##0.00\\ \"₽\";[Red]\\-#,##0.00\\ \"₽\"" };
            NumberingFormat numberingFormat5 = new NumberingFormat() { NumberFormatId = (UInt32Value)42U, FormatCode = "_-* #,##0\\ \"₽\"_-;\\-* #,##0\\ \"₽\"_-;_-* \"-\"\\ \"₽\"_-;_-@_-" };
            NumberingFormat numberingFormat6 = new NumberingFormat() { NumberFormatId = (UInt32Value)41U, FormatCode = "_-* #,##0_-;\\-* #,##0_-;_-* \"-\"_-;_-@_-" };
            NumberingFormat numberingFormat7 = new NumberingFormat() { NumberFormatId = (UInt32Value)44U, FormatCode = "_-* #,##0.00\\ \"₽\"_-;\\-* #,##0.00\\ \"₽\"_-;_-* \"-\"??\\ \"₽\"_-;_-@_-" };
            NumberingFormat numberingFormat8 = new NumberingFormat() { NumberFormatId = (UInt32Value)43U, FormatCode = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-" };

            numberingFormats1.Append(numberingFormat1);
            numberingFormats1.Append(numberingFormat2);
            numberingFormats1.Append(numberingFormat3);
            numberingFormats1.Append(numberingFormat4);
            numberingFormats1.Append(numberingFormat5);
            numberingFormats1.Append(numberingFormat6);
            numberingFormats1.Append(numberingFormat7);
            numberingFormats1.Append(numberingFormat8);

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)35U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Indexed = (UInt32Value)8U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 204 };

            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 18D };
            Color color3 = new Color() { Indexed = (UInt32Value)54U };
            FontName fontName3 = new FontName() { Val = "Calibri Light" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 204 };

            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet3);

            Font font4 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 15D };
            Color color4 = new Color() { Indexed = (UInt32Value)54U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = 204 };

            font4.Append(bold1);
            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontCharSet4);

            Font font5 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = 13D };
            Color color5 = new Color() { Indexed = (UInt32Value)54U };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = 204 };

            font5.Append(bold2);
            font5.Append(fontSize5);
            font5.Append(color5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontCharSet5);

            Font font6 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize6 = new FontSize() { Val = 11D };
            Color color6 = new Color() { Indexed = (UInt32Value)54U };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = 204 };

            font6.Append(bold3);
            font6.Append(fontSize6);
            font6.Append(color6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontCharSet6);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 11D };
            Color color7 = new Color() { Indexed = (UInt32Value)17U };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = 204 };

            font7.Append(fontSize7);
            font7.Append(color7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontCharSet7);

            Font font8 = new Font();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            Color color8 = new Color() { Indexed = (UInt32Value)20U };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = 204 };

            font8.Append(fontSize8);
            font8.Append(color8);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);
            font8.Append(fontCharSet8);

            Font font9 = new Font();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color9 = new Color() { Indexed = (UInt32Value)60U };
            FontName fontName9 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet9 = new FontCharSet() { Val = 204 };

            font9.Append(fontSize9);
            font9.Append(color9);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontCharSet9);

            Font font10 = new Font();
            FontSize fontSize10 = new FontSize() { Val = 11D };
            Color color10 = new Color() { Indexed = (UInt32Value)62U };
            FontName fontName10 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet10 = new FontCharSet() { Val = 204 };

            font10.Append(fontSize10);
            font10.Append(color10);
            font10.Append(fontName10);
            font10.Append(fontFamilyNumbering10);
            font10.Append(fontCharSet10);

            Font font11 = new Font();
            Bold bold4 = new Bold();
            FontSize fontSize11 = new FontSize() { Val = 11D };
            Color color11 = new Color() { Indexed = (UInt32Value)63U };
            FontName fontName11 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet11 = new FontCharSet() { Val = 204 };

            font11.Append(bold4);
            font11.Append(fontSize11);
            font11.Append(color11);
            font11.Append(fontName11);
            font11.Append(fontFamilyNumbering11);
            font11.Append(fontCharSet11);

            Font font12 = new Font();
            Bold bold5 = new Bold();
            FontSize fontSize12 = new FontSize() { Val = 11D };
            Color color12 = new Color() { Indexed = (UInt32Value)52U };
            FontName fontName12 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet12 = new FontCharSet() { Val = 204 };

            font12.Append(bold5);
            font12.Append(fontSize12);
            font12.Append(color12);
            font12.Append(fontName12);
            font12.Append(fontFamilyNumbering12);
            font12.Append(fontCharSet12);

            Font font13 = new Font();
            FontSize fontSize13 = new FontSize() { Val = 11D };
            Color color13 = new Color() { Indexed = (UInt32Value)52U };
            FontName fontName13 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet13 = new FontCharSet() { Val = 204 };

            font13.Append(fontSize13);
            font13.Append(color13);
            font13.Append(fontName13);
            font13.Append(fontFamilyNumbering13);
            font13.Append(fontCharSet13);

            Font font14 = new Font();
            Bold bold6 = new Bold();
            FontSize fontSize14 = new FontSize() { Val = 11D };
            Color color14 = new Color() { Indexed = (UInt32Value)9U };
            FontName fontName14 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet14 = new FontCharSet() { Val = 204 };

            font14.Append(bold6);
            font14.Append(fontSize14);
            font14.Append(color14);
            font14.Append(fontName14);
            font14.Append(fontFamilyNumbering14);
            font14.Append(fontCharSet14);

            Font font15 = new Font();
            FontSize fontSize15 = new FontSize() { Val = 11D };
            Color color15 = new Color() { Indexed = (UInt32Value)10U };
            FontName fontName15 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering15 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet15 = new FontCharSet() { Val = 204 };

            font15.Append(fontSize15);
            font15.Append(color15);
            font15.Append(fontName15);
            font15.Append(fontFamilyNumbering15);
            font15.Append(fontCharSet15);

            Font font16 = new Font();
            Italic italic1 = new Italic();
            FontSize fontSize16 = new FontSize() { Val = 11D };
            Color color16 = new Color() { Indexed = (UInt32Value)23U };
            FontName fontName16 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering16 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet16 = new FontCharSet() { Val = 204 };

            font16.Append(italic1);
            font16.Append(fontSize16);
            font16.Append(color16);
            font16.Append(fontName16);
            font16.Append(fontFamilyNumbering16);
            font16.Append(fontCharSet16);

            Font font17 = new Font();
            Bold bold7 = new Bold();
            FontSize fontSize17 = new FontSize() { Val = 11D };
            Color color17 = new Color() { Indexed = (UInt32Value)8U };
            FontName fontName17 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering17 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet17 = new FontCharSet() { Val = 204 };

            font17.Append(bold7);
            font17.Append(fontSize17);
            font17.Append(color17);
            font17.Append(fontName17);
            font17.Append(fontFamilyNumbering17);
            font17.Append(fontCharSet17);

            Font font18 = new Font();
            FontSize fontSize18 = new FontSize() { Val = 11D };
            Color color18 = new Color() { Indexed = (UInt32Value)9U };
            FontName fontName18 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering18 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet18 = new FontCharSet() { Val = 204 };

            font18.Append(fontSize18);
            font18.Append(color18);
            font18.Append(fontName18);
            font18.Append(fontFamilyNumbering18);
            font18.Append(fontCharSet18);

            Font font19 = new Font();
            FontSize fontSize19 = new FontSize() { Val = 8D };
            FontName fontName19 = new FontName() { Val = "Segoe UI" };
            FontFamilyNumbering fontFamilyNumbering19 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet19 = new FontCharSet() { Val = 204 };

            font19.Append(fontSize19);
            font19.Append(fontName19);
            font19.Append(fontFamilyNumbering19);
            font19.Append(fontCharSet19);

            Font font20 = new Font();
            FontSize fontSize20 = new FontSize() { Val = 11D };
            Color color19 = new Color() { Rgb = "FF006100" };
            FontName fontName20 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering20 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet20 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font20.Append(fontSize20);
            font20.Append(color19);
            font20.Append(fontName20);
            font20.Append(fontFamilyNumbering20);
            font20.Append(fontCharSet20);
            font20.Append(fontScheme2);

            Font font21 = new Font();
            FontSize fontSize21 = new FontSize() { Val = 11D };
            Color color20 = new Color() { Rgb = "FFFF0000" };
            FontName fontName21 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering21 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet21 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font21.Append(fontSize21);
            font21.Append(color20);
            font21.Append(fontName21);
            font21.Append(fontFamilyNumbering21);
            font21.Append(fontCharSet21);
            font21.Append(fontScheme3);

            Font font22 = new Font();
            FontSize fontSize22 = new FontSize() { Val = 11D };
            Color color21 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName22 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering22 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet22 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font22.Append(fontSize22);
            font22.Append(color21);
            font22.Append(fontName22);
            font22.Append(fontFamilyNumbering22);
            font22.Append(fontCharSet22);
            font22.Append(fontScheme4);

            Font font23 = new Font();
            Italic italic2 = new Italic();
            FontSize fontSize23 = new FontSize() { Val = 11D };
            Color color22 = new Color() { Rgb = "FF7F7F7F" };
            FontName fontName23 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering23 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet23 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font23.Append(italic2);
            font23.Append(fontSize23);
            font23.Append(color22);
            font23.Append(fontName23);
            font23.Append(fontFamilyNumbering23);
            font23.Append(fontCharSet23);
            font23.Append(fontScheme5);

            Font font24 = new Font();
            FontSize fontSize24 = new FontSize() { Val = 11D };
            Color color23 = new Color() { Rgb = "FF9C0006" };
            FontName fontName24 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering24 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet24 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font24.Append(fontSize24);
            font24.Append(color23);
            font24.Append(fontName24);
            font24.Append(fontFamilyNumbering24);
            font24.Append(fontCharSet24);
            font24.Append(fontScheme6);

            Font font25 = new Font();
            FontSize fontSize25 = new FontSize() { Val = 11D };
            Color color24 = new Color() { Rgb = "FF9C6500" };
            FontName fontName25 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering25 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet25 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font25.Append(fontSize25);
            font25.Append(color24);
            font25.Append(fontName25);
            font25.Append(fontFamilyNumbering25);
            font25.Append(fontCharSet25);
            font25.Append(fontScheme7);

            Font font26 = new Font();
            FontSize fontSize26 = new FontSize() { Val = 18D };
            Color color25 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName26 = new FontName() { Val = "Calibri Light" };
            FontFamilyNumbering fontFamilyNumbering26 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet26 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Major };

            font26.Append(fontSize26);
            font26.Append(color25);
            font26.Append(fontName26);
            font26.Append(fontFamilyNumbering26);
            font26.Append(fontCharSet26);
            font26.Append(fontScheme8);

            Font font27 = new Font();
            Bold bold8 = new Bold();
            FontSize fontSize27 = new FontSize() { Val = 11D };
            Color color26 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName27 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering27 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet27 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font27.Append(bold8);
            font27.Append(fontSize27);
            font27.Append(color26);
            font27.Append(fontName27);
            font27.Append(fontFamilyNumbering27);
            font27.Append(fontCharSet27);
            font27.Append(fontScheme9);

            Font font28 = new Font();
            Bold bold9 = new Bold();
            FontSize fontSize28 = new FontSize() { Val = 11D };
            Color color27 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName28 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering28 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet28 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme10 = new FontScheme() { Val = FontSchemeValues.Minor };

            font28.Append(bold9);
            font28.Append(fontSize28);
            font28.Append(color27);
            font28.Append(fontName28);
            font28.Append(fontFamilyNumbering28);
            font28.Append(fontCharSet28);
            font28.Append(fontScheme10);

            Font font29 = new Font();
            Bold bold10 = new Bold();
            FontSize fontSize29 = new FontSize() { Val = 11D };
            Color color28 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName29 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering29 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet29 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme11 = new FontScheme() { Val = FontSchemeValues.Minor };

            font29.Append(bold10);
            font29.Append(fontSize29);
            font29.Append(color28);
            font29.Append(fontName29);
            font29.Append(fontFamilyNumbering29);
            font29.Append(fontCharSet29);
            font29.Append(fontScheme11);

            Font font30 = new Font();
            Bold bold11 = new Bold();
            FontSize fontSize30 = new FontSize() { Val = 13D };
            Color color29 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName30 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering30 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet30 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme12 = new FontScheme() { Val = FontSchemeValues.Minor };

            font30.Append(bold11);
            font30.Append(fontSize30);
            font30.Append(color29);
            font30.Append(fontName30);
            font30.Append(fontFamilyNumbering30);
            font30.Append(fontCharSet30);
            font30.Append(fontScheme12);

            Font font31 = new Font();
            Bold bold12 = new Bold();
            FontSize fontSize31 = new FontSize() { Val = 15D };
            Color color30 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName31 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering31 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet31 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme13 = new FontScheme() { Val = FontSchemeValues.Minor };

            font31.Append(bold12);
            font31.Append(fontSize31);
            font31.Append(color30);
            font31.Append(fontName31);
            font31.Append(fontFamilyNumbering31);
            font31.Append(fontCharSet31);
            font31.Append(fontScheme13);

            Font font32 = new Font();
            Bold bold13 = new Bold();
            FontSize fontSize32 = new FontSize() { Val = 11D };
            Color color31 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName32 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering32 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet32 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme14 = new FontScheme() { Val = FontSchemeValues.Minor };

            font32.Append(bold13);
            font32.Append(fontSize32);
            font32.Append(color31);
            font32.Append(fontName32);
            font32.Append(fontFamilyNumbering32);
            font32.Append(fontCharSet32);
            font32.Append(fontScheme14);

            Font font33 = new Font();
            Bold bold14 = new Bold();
            FontSize fontSize33 = new FontSize() { Val = 11D };
            Color color32 = new Color() { Rgb = "FF3F3F3F" };
            FontName fontName33 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering33 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet33 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme15 = new FontScheme() { Val = FontSchemeValues.Minor };

            font33.Append(bold14);
            font33.Append(fontSize33);
            font33.Append(color32);
            font33.Append(fontName33);
            font33.Append(fontFamilyNumbering33);
            font33.Append(fontCharSet33);
            font33.Append(fontScheme15);

            Font font34 = new Font();
            FontSize fontSize34 = new FontSize() { Val = 11D };
            Color color33 = new Color() { Rgb = "FF3F3F76" };
            FontName fontName34 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering34 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet34 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme16 = new FontScheme() { Val = FontSchemeValues.Minor };

            font34.Append(fontSize34);
            font34.Append(color33);
            font34.Append(fontName34);
            font34.Append(fontFamilyNumbering34);
            font34.Append(fontCharSet34);
            font34.Append(fontScheme16);

            Font font35 = new Font();
            FontSize fontSize35 = new FontSize() { Val = 11D };
            Color color34 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName35 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering35 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet35 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme17 = new FontScheme() { Val = FontSchemeValues.Minor };

            font35.Append(fontSize35);
            font35.Append(color34);
            font35.Append(fontName35);
            font35.Append(fontFamilyNumbering35);
            font35.Append(fontCharSet35);
            font35.Append(fontScheme17);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);
            fonts1.Append(font11);
            fonts1.Append(font12);
            fonts1.Append(font13);
            fonts1.Append(font14);
            fonts1.Append(font15);
            fonts1.Append(font16);
            fonts1.Append(font17);
            fonts1.Append(font18);
            fonts1.Append(font19);
            fonts1.Append(font20);
            fonts1.Append(font21);
            fonts1.Append(font22);
            fonts1.Append(font23);
            fonts1.Append(font24);
            fonts1.Append(font25);
            fonts1.Append(font26);
            fonts1.Append(font27);
            fonts1.Append(font28);
            fonts1.Append(font29);
            fonts1.Append(font30);
            fonts1.Append(font31);
            fonts1.Append(font32);
            fonts1.Append(font33);
            fonts1.Append(font34);
            fonts1.Append(font35);

            Fills fills1 = new Fills() { Count = (UInt32Value)33U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.79998D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.79998D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.79998D };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.79998D };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill6.Append(foregroundColor4);
            patternFill6.Append(backgroundColor4);

            fill6.Append(patternFill6);

            Fill fill7 = new Fill();

            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor5 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.79998D };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill7.Append(foregroundColor5);
            patternFill7.Append(backgroundColor5);

            fill7.Append(patternFill7);

            Fill fill8 = new Fill();

            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor6 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.79998D };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill8.Append(foregroundColor6);
            patternFill8.Append(backgroundColor6);

            fill8.Append(patternFill8);

            Fill fill9 = new Fill();

            PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor7 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.59999D };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill9.Append(foregroundColor7);
            patternFill9.Append(backgroundColor7);

            fill9.Append(patternFill9);

            Fill fill10 = new Fill();

            PatternFill patternFill10 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor8 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.59999D };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill10.Append(foregroundColor8);
            patternFill10.Append(backgroundColor8);

            fill10.Append(patternFill10);

            Fill fill11 = new Fill();

            PatternFill patternFill11 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor9 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.59999D };
            BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill11.Append(foregroundColor9);
            patternFill11.Append(backgroundColor9);

            fill11.Append(patternFill11);

            Fill fill12 = new Fill();

            PatternFill patternFill12 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor10 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.59999D };
            BackgroundColor backgroundColor10 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill12.Append(foregroundColor10);
            patternFill12.Append(backgroundColor10);

            fill12.Append(patternFill12);

            Fill fill13 = new Fill();

            PatternFill patternFill13 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor11 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.59999D };
            BackgroundColor backgroundColor11 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill13.Append(foregroundColor11);
            patternFill13.Append(backgroundColor11);

            fill13.Append(patternFill13);

            Fill fill14 = new Fill();

            PatternFill patternFill14 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor12 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.59999D };
            BackgroundColor backgroundColor12 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill14.Append(foregroundColor12);
            patternFill14.Append(backgroundColor12);

            fill14.Append(patternFill14);

            Fill fill15 = new Fill();

            PatternFill patternFill15 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor13 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.39998D };
            BackgroundColor backgroundColor13 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill15.Append(foregroundColor13);
            patternFill15.Append(backgroundColor13);

            fill15.Append(patternFill15);

            Fill fill16 = new Fill();

            PatternFill patternFill16 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor14 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.39998D };
            BackgroundColor backgroundColor14 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill16.Append(foregroundColor14);
            patternFill16.Append(backgroundColor14);

            fill16.Append(patternFill16);

            Fill fill17 = new Fill();

            PatternFill patternFill17 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor15 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.39998D };
            BackgroundColor backgroundColor15 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill17.Append(foregroundColor15);
            patternFill17.Append(backgroundColor15);

            fill17.Append(patternFill17);

            Fill fill18 = new Fill();

            PatternFill patternFill18 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor16 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.39998D };
            BackgroundColor backgroundColor16 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill18.Append(foregroundColor16);
            patternFill18.Append(backgroundColor16);

            fill18.Append(patternFill18);

            Fill fill19 = new Fill();

            PatternFill patternFill19 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor17 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.39998D };
            BackgroundColor backgroundColor17 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill19.Append(foregroundColor17);
            patternFill19.Append(backgroundColor17);

            fill19.Append(patternFill19);

            Fill fill20 = new Fill();

            PatternFill patternFill20 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor18 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.39998D };
            BackgroundColor backgroundColor18 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill20.Append(foregroundColor18);
            patternFill20.Append(backgroundColor18);

            fill20.Append(patternFill20);

            Fill fill21 = new Fill();

            PatternFill patternFill21 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor19 = new ForegroundColor() { Theme = (UInt32Value)4U };
            BackgroundColor backgroundColor19 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill21.Append(foregroundColor19);
            patternFill21.Append(backgroundColor19);

            fill21.Append(patternFill21);

            Fill fill22 = new Fill();

            PatternFill patternFill22 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor20 = new ForegroundColor() { Theme = (UInt32Value)5U };
            BackgroundColor backgroundColor20 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill22.Append(foregroundColor20);
            patternFill22.Append(backgroundColor20);

            fill22.Append(patternFill22);

            Fill fill23 = new Fill();

            PatternFill patternFill23 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor21 = new ForegroundColor() { Theme = (UInt32Value)6U };
            BackgroundColor backgroundColor21 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill23.Append(foregroundColor21);
            patternFill23.Append(backgroundColor21);

            fill23.Append(patternFill23);

            Fill fill24 = new Fill();

            PatternFill patternFill24 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor22 = new ForegroundColor() { Theme = (UInt32Value)7U };
            BackgroundColor backgroundColor22 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill24.Append(foregroundColor22);
            patternFill24.Append(backgroundColor22);

            fill24.Append(patternFill24);

            Fill fill25 = new Fill();

            PatternFill patternFill25 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor23 = new ForegroundColor() { Theme = (UInt32Value)8U };
            BackgroundColor backgroundColor23 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill25.Append(foregroundColor23);
            patternFill25.Append(backgroundColor23);

            fill25.Append(patternFill25);

            Fill fill26 = new Fill();

            PatternFill patternFill26 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor24 = new ForegroundColor() { Theme = (UInt32Value)9U };
            BackgroundColor backgroundColor24 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill26.Append(foregroundColor24);
            patternFill26.Append(backgroundColor24);

            fill26.Append(patternFill26);

            Fill fill27 = new Fill();

            PatternFill patternFill27 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor25 = new ForegroundColor() { Rgb = "FFFFCC99" };
            BackgroundColor backgroundColor25 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill27.Append(foregroundColor25);
            patternFill27.Append(backgroundColor25);

            fill27.Append(patternFill27);

            Fill fill28 = new Fill();

            PatternFill patternFill28 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor26 = new ForegroundColor() { Rgb = "FFF2F2F2" };
            BackgroundColor backgroundColor26 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill28.Append(foregroundColor26);
            patternFill28.Append(backgroundColor26);

            fill28.Append(patternFill28);

            Fill fill29 = new Fill();

            PatternFill patternFill29 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor27 = new ForegroundColor() { Rgb = "FFA5A5A5" };
            BackgroundColor backgroundColor27 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill29.Append(foregroundColor27);
            patternFill29.Append(backgroundColor27);

            fill29.Append(patternFill29);

            Fill fill30 = new Fill();

            PatternFill patternFill30 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor28 = new ForegroundColor() { Rgb = "FFFFEB9C" };
            BackgroundColor backgroundColor28 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill30.Append(foregroundColor28);
            patternFill30.Append(backgroundColor28);

            fill30.Append(patternFill30);

            Fill fill31 = new Fill();

            PatternFill patternFill31 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor29 = new ForegroundColor() { Rgb = "FFFFC7CE" };
            BackgroundColor backgroundColor29 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill31.Append(foregroundColor29);
            patternFill31.Append(backgroundColor29);

            fill31.Append(patternFill31);

            Fill fill32 = new Fill();

            PatternFill patternFill32 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor30 = new ForegroundColor() { Rgb = "FFFFFFCC" };
            BackgroundColor backgroundColor30 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill32.Append(foregroundColor30);
            patternFill32.Append(backgroundColor30);

            fill32.Append(patternFill32);

            Fill fill33 = new Fill();

            PatternFill patternFill33 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor31 = new ForegroundColor() { Rgb = "FFC6EFCE" };
            BackgroundColor backgroundColor31 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill33.Append(foregroundColor31);
            patternFill33.Append(backgroundColor31);

            fill33.Append(patternFill33);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);
            fills1.Append(fill7);
            fills1.Append(fill8);
            fills1.Append(fill9);
            fills1.Append(fill10);
            fills1.Append(fill11);
            fills1.Append(fill12);
            fills1.Append(fill13);
            fills1.Append(fill14);
            fills1.Append(fill15);
            fills1.Append(fill16);
            fills1.Append(fill17);
            fills1.Append(fill18);
            fills1.Append(fill19);
            fills1.Append(fill20);
            fills1.Append(fill21);
            fills1.Append(fill22);
            fills1.Append(fill23);
            fills1.Append(fill24);
            fills1.Append(fill25);
            fills1.Append(fill26);
            fills1.Append(fill27);
            fills1.Append(fill28);
            fills1.Append(fill29);
            fills1.Append(fill30);
            fills1.Append(fill31);
            fills1.Append(fill32);
            fills1.Append(fill33);

            Borders borders1 = new Borders() { Count = (UInt32Value)13U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color35 = new Color() { Rgb = "FF7F7F7F" };

            leftBorder2.Append(color35);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color36 = new Color() { Rgb = "FF7F7F7F" };

            rightBorder2.Append(color36);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color37 = new Color() { Rgb = "FF7F7F7F" };

            topBorder2.Append(color37);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color38 = new Color() { Rgb = "FF7F7F7F" };

            bottomBorder2.Append(color38);

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);

            Border border3 = new Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color39 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder3.Append(color39);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color40 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder3.Append(color40);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color41 = new Color() { Rgb = "FF3F3F3F" };

            topBorder3.Append(color41);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color42 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder3.Append(color42);

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);

            Border border4 = new Border();

            LeftBorder leftBorder4 = new LeftBorder();
            Color color43 = new Color() { Indexed = (UInt32Value)0U };

            leftBorder4.Append(color43);

            RightBorder rightBorder4 = new RightBorder();
            Color color44 = new Color() { Indexed = (UInt32Value)0U };

            rightBorder4.Append(color44);

            TopBorder topBorder4 = new TopBorder();
            Color color45 = new Color() { Indexed = (UInt32Value)0U };

            topBorder4.Append(color45);

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color46 = new Color() { Theme = (UInt32Value)4U };

            bottomBorder4.Append(color46);

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);

            Border border5 = new Border();

            LeftBorder leftBorder5 = new LeftBorder();
            Color color47 = new Color() { Indexed = (UInt32Value)0U };

            leftBorder5.Append(color47);

            RightBorder rightBorder5 = new RightBorder();
            Color color48 = new Color() { Indexed = (UInt32Value)0U };

            rightBorder5.Append(color48);

            TopBorder topBorder5 = new TopBorder();
            Color color49 = new Color() { Indexed = (UInt32Value)0U };

            topBorder5.Append(color49);

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color50 = new Color() { Theme = (UInt32Value)4U, Tint = 0.49998D };

            bottomBorder5.Append(color50);

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);

            Border border6 = new Border();

            LeftBorder leftBorder6 = new LeftBorder();
            Color color51 = new Color() { Indexed = (UInt32Value)0U };

            leftBorder6.Append(color51);

            RightBorder rightBorder6 = new RightBorder();
            Color color52 = new Color() { Indexed = (UInt32Value)0U };

            rightBorder6.Append(color52);

            TopBorder topBorder6 = new TopBorder();
            Color color53 = new Color() { Indexed = (UInt32Value)0U };

            topBorder6.Append(color53);

            BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color54 = new Color() { Theme = (UInt32Value)4U, Tint = 0.39998D };

            bottomBorder6.Append(color54);

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);

            Border border7 = new Border();

            LeftBorder leftBorder7 = new LeftBorder();
            Color color55 = new Color() { Indexed = (UInt32Value)0U };

            leftBorder7.Append(color55);

            RightBorder rightBorder7 = new RightBorder();
            Color color56 = new Color() { Indexed = (UInt32Value)0U };

            rightBorder7.Append(color56);

            TopBorder topBorder7 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color57 = new Color() { Theme = (UInt32Value)4U };

            topBorder7.Append(color57);

            BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color58 = new Color() { Theme = (UInt32Value)4U };

            bottomBorder7.Append(color58);

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);

            Border border8 = new Border();

            LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Double };
            Color color59 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder8.Append(color59);

            RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Double };
            Color color60 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder8.Append(color60);

            TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Double };
            Color color61 = new Color() { Rgb = "FF3F3F3F" };

            topBorder8.Append(color61);

            BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color62 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder8.Append(color62);

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);

            Border border9 = new Border();

            LeftBorder leftBorder9 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color63 = new Color() { Rgb = "FFB2B2B2" };

            leftBorder9.Append(color63);

            RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color64 = new Color() { Rgb = "FFB2B2B2" };

            rightBorder9.Append(color64);

            TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color65 = new Color() { Rgb = "FFB2B2B2" };

            topBorder9.Append(color65);

            BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color66 = new Color() { Rgb = "FFB2B2B2" };

            bottomBorder9.Append(color66);

            border9.Append(leftBorder9);
            border9.Append(rightBorder9);
            border9.Append(topBorder9);
            border9.Append(bottomBorder9);

            Border border10 = new Border();

            LeftBorder leftBorder10 = new LeftBorder();
            Color color67 = new Color() { Indexed = (UInt32Value)0U };

            leftBorder10.Append(color67);

            RightBorder rightBorder10 = new RightBorder();
            Color color68 = new Color() { Indexed = (UInt32Value)0U };

            rightBorder10.Append(color68);

            TopBorder topBorder10 = new TopBorder();
            Color color69 = new Color() { Indexed = (UInt32Value)0U };

            topBorder10.Append(color69);

            BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color70 = new Color() { Rgb = "FFFF8001" };

            bottomBorder10.Append(color70);

            border10.Append(leftBorder10);
            border10.Append(rightBorder10);
            border10.Append(topBorder10);
            border10.Append(bottomBorder10);

            Border border11 = new Border();

            LeftBorder leftBorder11 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color71 = new Color() { Auto = true };

            leftBorder11.Append(color71);

            RightBorder rightBorder11 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color72 = new Color() { Auto = true };

            rightBorder11.Append(color72);

            TopBorder topBorder11 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color73 = new Color() { Auto = true };

            topBorder11.Append(color73);

            BottomBorder bottomBorder11 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color74 = new Color() { Auto = true };

            bottomBorder11.Append(color74);

            border11.Append(leftBorder11);
            border11.Append(rightBorder11);
            border11.Append(topBorder11);
            border11.Append(bottomBorder11);

            Border border12 = new Border();

            LeftBorder leftBorder12 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color75 = new Color() { Auto = true };

            leftBorder12.Append(color75);

            RightBorder rightBorder12 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color76 = new Color() { Auto = true };

            rightBorder12.Append(color76);

            TopBorder topBorder12 = new TopBorder();
            Color color77 = new Color() { Indexed = (UInt32Value)0U };

            topBorder12.Append(color77);

            BottomBorder bottomBorder12 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color78 = new Color() { Auto = true };

            bottomBorder12.Append(color78);

            border12.Append(leftBorder12);
            border12.Append(rightBorder12);
            border12.Append(topBorder12);
            border12.Append(bottomBorder12);

            Border border13 = new Border();

            LeftBorder leftBorder13 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color79 = new Color() { Auto = true };

            leftBorder13.Append(color79);

            RightBorder rightBorder13 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color80 = new Color() { Auto = true };

            rightBorder13.Append(color80);

            TopBorder topBorder13 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color81 = new Color() { Auto = true };

            topBorder13.Append(color81);

            BottomBorder bottomBorder13 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color82 = new Color() { Auto = true };

            bottomBorder13.Append(color82);

            border13.Append(leftBorder13);
            border13.Append(rightBorder13);
            border13.Append(topBorder13);
            border13.Append(bottomBorder13);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);
            borders1.Append(border6);
            borders1.Append(border7);
            borders1.Append(border8);
            borders1.Append(border9);
            borders1.Append(border10);
            borders1.Append(border11);
            borders1.Append(border12);
            borders1.Append(border13);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)61U };

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment1 = new Alignment();
            Protection protection1 = new Protection();

            cellFormat1.Append(alignment1);
            cellFormat1.Append(protection1);

            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment2 = new Alignment();
            Protection protection2 = new Protection();

            cellFormat2.Append(alignment2);
            cellFormat2.Append(protection2);

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment3 = new Alignment();
            Protection protection3 = new Protection();

            cellFormat3.Append(alignment3);
            cellFormat3.Append(protection3);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment4 = new Alignment();
            Protection protection4 = new Protection();

            cellFormat4.Append(alignment4);
            cellFormat4.Append(protection4);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment5 = new Alignment();
            Protection protection5 = new Protection();

            cellFormat5.Append(alignment5);
            cellFormat5.Append(protection5);

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment6 = new Alignment();
            Protection protection6 = new Protection();

            cellFormat6.Append(alignment6);
            cellFormat6.Append(protection6);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment7 = new Alignment();
            Protection protection7 = new Protection();

            cellFormat7.Append(alignment7);
            cellFormat7.Append(protection7);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment8 = new Alignment();
            Protection protection8 = new Protection();

            cellFormat8.Append(alignment8);
            cellFormat8.Append(protection8);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment9 = new Alignment();
            Protection protection9 = new Protection();

            cellFormat9.Append(alignment9);
            cellFormat9.Append(protection9);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment10 = new Alignment();
            Protection protection10 = new Protection();

            cellFormat10.Append(alignment10);
            cellFormat10.Append(protection10);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment11 = new Alignment();
            Protection protection11 = new Protection();

            cellFormat11.Append(alignment11);
            cellFormat11.Append(protection11);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment12 = new Alignment();
            Protection protection12 = new Protection();

            cellFormat12.Append(alignment12);
            cellFormat12.Append(protection12);

            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment13 = new Alignment();
            Protection protection13 = new Protection();

            cellFormat13.Append(alignment13);
            cellFormat13.Append(protection13);

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment14 = new Alignment();
            Protection protection14 = new Protection();

            cellFormat14.Append(alignment14);
            cellFormat14.Append(protection14);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment15 = new Alignment();
            Protection protection15 = new Protection();

            cellFormat15.Append(alignment15);
            cellFormat15.Append(protection15);
            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)8U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)9U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)10U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)11U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)12U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)13U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)14U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)15U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)16U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)17U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)18U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)19U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)20U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)21U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)22U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)23U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)24U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)34U, FillId = (UInt32Value)25U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)33U, FillId = (UInt32Value)26U, BorderId = (UInt32Value)1U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)32U, FillId = (UInt32Value)27U, BorderId = (UInt32Value)2U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)31U, FillId = (UInt32Value)27U, BorderId = (UInt32Value)1U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)44U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = (UInt32Value)42U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat45 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)30U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat46 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)29U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat47 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)28U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)5U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat48 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)28U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat49 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)27U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat50 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)26U, FillId = (UInt32Value)28U, BorderId = (UInt32Value)7U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat51 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)25U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat52 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)24U, FillId = (UInt32Value)29U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat53 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)23U, FillId = (UInt32Value)30U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat54 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)22U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat55 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)31U, BorderId = (UInt32Value)8U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat56 = new CellFormat() { NumberFormatId = (UInt32Value)9U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat57 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)21U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)9U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat58 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)20U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat59 = new CellFormat() { NumberFormatId = (UInt32Value)43U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat60 = new CellFormat() { NumberFormatId = (UInt32Value)41U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat61 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)19U, FillId = (UInt32Value)32U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);
            cellStyleFormats1.Append(cellFormat3);
            cellStyleFormats1.Append(cellFormat4);
            cellStyleFormats1.Append(cellFormat5);
            cellStyleFormats1.Append(cellFormat6);
            cellStyleFormats1.Append(cellFormat7);
            cellStyleFormats1.Append(cellFormat8);
            cellStyleFormats1.Append(cellFormat9);
            cellStyleFormats1.Append(cellFormat10);
            cellStyleFormats1.Append(cellFormat11);
            cellStyleFormats1.Append(cellFormat12);
            cellStyleFormats1.Append(cellFormat13);
            cellStyleFormats1.Append(cellFormat14);
            cellStyleFormats1.Append(cellFormat15);
            cellStyleFormats1.Append(cellFormat16);
            cellStyleFormats1.Append(cellFormat17);
            cellStyleFormats1.Append(cellFormat18);
            cellStyleFormats1.Append(cellFormat19);
            cellStyleFormats1.Append(cellFormat20);
            cellStyleFormats1.Append(cellFormat21);
            cellStyleFormats1.Append(cellFormat22);
            cellStyleFormats1.Append(cellFormat23);
            cellStyleFormats1.Append(cellFormat24);
            cellStyleFormats1.Append(cellFormat25);
            cellStyleFormats1.Append(cellFormat26);
            cellStyleFormats1.Append(cellFormat27);
            cellStyleFormats1.Append(cellFormat28);
            cellStyleFormats1.Append(cellFormat29);
            cellStyleFormats1.Append(cellFormat30);
            cellStyleFormats1.Append(cellFormat31);
            cellStyleFormats1.Append(cellFormat32);
            cellStyleFormats1.Append(cellFormat33);
            cellStyleFormats1.Append(cellFormat34);
            cellStyleFormats1.Append(cellFormat35);
            cellStyleFormats1.Append(cellFormat36);
            cellStyleFormats1.Append(cellFormat37);
            cellStyleFormats1.Append(cellFormat38);
            cellStyleFormats1.Append(cellFormat39);
            cellStyleFormats1.Append(cellFormat40);
            cellStyleFormats1.Append(cellFormat41);
            cellStyleFormats1.Append(cellFormat42);
            cellStyleFormats1.Append(cellFormat43);
            cellStyleFormats1.Append(cellFormat44);
            cellStyleFormats1.Append(cellFormat45);
            cellStyleFormats1.Append(cellFormat46);
            cellStyleFormats1.Append(cellFormat47);
            cellStyleFormats1.Append(cellFormat48);
            cellStyleFormats1.Append(cellFormat49);
            cellStyleFormats1.Append(cellFormat50);
            cellStyleFormats1.Append(cellFormat51);
            cellStyleFormats1.Append(cellFormat52);
            cellStyleFormats1.Append(cellFormat53);
            cellStyleFormats1.Append(cellFormat54);
            cellStyleFormats1.Append(cellFormat55);
            cellStyleFormats1.Append(cellFormat56);
            cellStyleFormats1.Append(cellFormat57);
            cellStyleFormats1.Append(cellFormat58);
            cellStyleFormats1.Append(cellFormat59);
            cellStyleFormats1.Append(cellFormat60);
            cellStyleFormats1.Append(cellFormat61);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)16U };

            CellFormat cellFormat62 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment16 = new Alignment();

            cellFormat62.Append(alignment16);

            CellFormat cellFormat63 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment17 = new Alignment();

            cellFormat63.Append(alignment17);

            CellFormat cellFormat64 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment18 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat64.Append(alignment18);

            CellFormat cellFormat65 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment19 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat65.Append(alignment19);

            CellFormat cellFormat66 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment20 = new Alignment();

            cellFormat66.Append(alignment20);

            CellFormat cellFormat67 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment21 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat67.Append(alignment21);

            CellFormat cellFormat68 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment22 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat68.Append(alignment22);

            CellFormat cellFormat69 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment23 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat69.Append(alignment23);

            CellFormat cellFormat70 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)58U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment24 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat70.Append(alignment24);

            CellFormat cellFormat71 = new CellFormat() { NumberFormatId = (UInt32Value)1U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)58U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment25 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat71.Append(alignment25);

            CellFormat cellFormat72 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)58U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment26 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat72.Append(alignment26);

            CellFormat cellFormat73 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment27 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat73.Append(alignment27);

            CellFormat cellFormat74 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment28 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat74.Append(alignment28);

            CellFormat cellFormat75 = new CellFormat() { NumberFormatId = (UInt32Value)1U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)58U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment29 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat75.Append(alignment29);

            CellFormat cellFormat76 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment30 = new Alignment() { Vertical = VerticalAlignmentValues.Top, TextRotation = (UInt32Value)90U };

            cellFormat76.Append(alignment30);

            CellFormat cellFormat77 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment31 = new Alignment() { Vertical = VerticalAlignmentValues.Top, TextRotation = (UInt32Value)90U, WrapText = true };

            cellFormat77.Append(alignment31);

            cellFormats1.Append(cellFormat62);
            cellFormats1.Append(cellFormat63);
            cellFormats1.Append(cellFormat64);
            cellFormats1.Append(cellFormat65);
            cellFormats1.Append(cellFormat66);
            cellFormats1.Append(cellFormat67);
            cellFormats1.Append(cellFormat68);
            cellFormats1.Append(cellFormat69);
            cellFormats1.Append(cellFormat70);
            cellFormats1.Append(cellFormat71);
            cellFormats1.Append(cellFormat72);
            cellFormats1.Append(cellFormat73);
            cellFormats1.Append(cellFormat74);
            cellFormats1.Append(cellFormat75);
            cellFormats1.Append(cellFormat76);
            cellFormats1.Append(cellFormat77);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)47U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
            CellStyle cellStyle2 = new CellStyle() { Name = "20% — акцент1", FormatId = (UInt32Value)15U };
            CellStyle cellStyle3 = new CellStyle() { Name = "20% — акцент2", FormatId = (UInt32Value)16U };
            CellStyle cellStyle4 = new CellStyle() { Name = "20% — акцент3", FormatId = (UInt32Value)17U };
            CellStyle cellStyle5 = new CellStyle() { Name = "20% — акцент4", FormatId = (UInt32Value)18U };
            CellStyle cellStyle6 = new CellStyle() { Name = "20% — акцент5", FormatId = (UInt32Value)19U };
            CellStyle cellStyle7 = new CellStyle() { Name = "20% — акцент6", FormatId = (UInt32Value)20U };
            CellStyle cellStyle8 = new CellStyle() { Name = "40% — акцент1", FormatId = (UInt32Value)21U };
            CellStyle cellStyle9 = new CellStyle() { Name = "40% — акцент2", FormatId = (UInt32Value)22U };
            CellStyle cellStyle10 = new CellStyle() { Name = "40% — акцент3", FormatId = (UInt32Value)23U };
            CellStyle cellStyle11 = new CellStyle() { Name = "40% — акцент4", FormatId = (UInt32Value)24U };
            CellStyle cellStyle12 = new CellStyle() { Name = "40% — акцент5", FormatId = (UInt32Value)25U };
            CellStyle cellStyle13 = new CellStyle() { Name = "40% — акцент6", FormatId = (UInt32Value)26U };
            CellStyle cellStyle14 = new CellStyle() { Name = "60% — акцент1", FormatId = (UInt32Value)27U };
            CellStyle cellStyle15 = new CellStyle() { Name = "60% — акцент2", FormatId = (UInt32Value)28U };
            CellStyle cellStyle16 = new CellStyle() { Name = "60% — акцент3", FormatId = (UInt32Value)29U };
            CellStyle cellStyle17 = new CellStyle() { Name = "60% — акцент4", FormatId = (UInt32Value)30U };
            CellStyle cellStyle18 = new CellStyle() { Name = "60% — акцент5", FormatId = (UInt32Value)31U };
            CellStyle cellStyle19 = new CellStyle() { Name = "60% — акцент6", FormatId = (UInt32Value)32U };
            CellStyle cellStyle20 = new CellStyle() { Name = "Акцент1", FormatId = (UInt32Value)33U };
            CellStyle cellStyle21 = new CellStyle() { Name = "Акцент2", FormatId = (UInt32Value)34U };
            CellStyle cellStyle22 = new CellStyle() { Name = "Акцент3", FormatId = (UInt32Value)35U };
            CellStyle cellStyle23 = new CellStyle() { Name = "Акцент4", FormatId = (UInt32Value)36U };
            CellStyle cellStyle24 = new CellStyle() { Name = "Акцент5", FormatId = (UInt32Value)37U };
            CellStyle cellStyle25 = new CellStyle() { Name = "Акцент6", FormatId = (UInt32Value)38U };
            CellStyle cellStyle26 = new CellStyle() { Name = "Ввод ", FormatId = (UInt32Value)39U };
            CellStyle cellStyle27 = new CellStyle() { Name = "Вывод", FormatId = (UInt32Value)40U };
            CellStyle cellStyle28 = new CellStyle() { Name = "Вычисление", FormatId = (UInt32Value)41U };
            CellStyle cellStyle29 = new CellStyle() { Name = "Currency", FormatId = (UInt32Value)42U, BuiltinId = (UInt32Value)4U };
            CellStyle cellStyle30 = new CellStyle() { Name = "Currency [0]", FormatId = (UInt32Value)43U, BuiltinId = (UInt32Value)7U };
            CellStyle cellStyle31 = new CellStyle() { Name = "Заголовок 1", FormatId = (UInt32Value)44U };
            CellStyle cellStyle32 = new CellStyle() { Name = "Заголовок 2", FormatId = (UInt32Value)45U };
            CellStyle cellStyle33 = new CellStyle() { Name = "Заголовок 3", FormatId = (UInt32Value)46U };
            CellStyle cellStyle34 = new CellStyle() { Name = "Заголовок 4", FormatId = (UInt32Value)47U };
            CellStyle cellStyle35 = new CellStyle() { Name = "Итог", FormatId = (UInt32Value)48U };
            CellStyle cellStyle36 = new CellStyle() { Name = "Контрольная ячейка", FormatId = (UInt32Value)49U };
            CellStyle cellStyle37 = new CellStyle() { Name = "Название", FormatId = (UInt32Value)50U };
            CellStyle cellStyle38 = new CellStyle() { Name = "Нейтральный", FormatId = (UInt32Value)51U };
            CellStyle cellStyle39 = new CellStyle() { Name = "Плохой", FormatId = (UInt32Value)52U };
            CellStyle cellStyle40 = new CellStyle() { Name = "Пояснение", FormatId = (UInt32Value)53U };
            CellStyle cellStyle41 = new CellStyle() { Name = "Примечание", FormatId = (UInt32Value)54U };
            CellStyle cellStyle42 = new CellStyle() { Name = "Percent", FormatId = (UInt32Value)55U, BuiltinId = (UInt32Value)5U };
            CellStyle cellStyle43 = new CellStyle() { Name = "Связанная ячейка", FormatId = (UInt32Value)56U };
            CellStyle cellStyle44 = new CellStyle() { Name = "Текст предупреждения", FormatId = (UInt32Value)57U };
            CellStyle cellStyle45 = new CellStyle() { Name = "Comma", FormatId = (UInt32Value)58U, BuiltinId = (UInt32Value)3U };
            CellStyle cellStyle46 = new CellStyle() { Name = "Comma [0]", FormatId = (UInt32Value)59U, BuiltinId = (UInt32Value)6U };
            CellStyle cellStyle47 = new CellStyle() { Name = "Хороший", FormatId = (UInt32Value)60U };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            cellStyles1.Append(cellStyle3);
            cellStyles1.Append(cellStyle4);
            cellStyles1.Append(cellStyle5);
            cellStyles1.Append(cellStyle6);
            cellStyles1.Append(cellStyle7);
            cellStyles1.Append(cellStyle8);
            cellStyles1.Append(cellStyle9);
            cellStyles1.Append(cellStyle10);
            cellStyles1.Append(cellStyle11);
            cellStyles1.Append(cellStyle12);
            cellStyles1.Append(cellStyle13);
            cellStyles1.Append(cellStyle14);
            cellStyles1.Append(cellStyle15);
            cellStyles1.Append(cellStyle16);
            cellStyles1.Append(cellStyle17);
            cellStyles1.Append(cellStyle18);
            cellStyles1.Append(cellStyle19);
            cellStyles1.Append(cellStyle20);
            cellStyles1.Append(cellStyle21);
            cellStyles1.Append(cellStyle22);
            cellStyles1.Append(cellStyle23);
            cellStyles1.Append(cellStyle24);
            cellStyles1.Append(cellStyle25);
            cellStyles1.Append(cellStyle26);
            cellStyles1.Append(cellStyle27);
            cellStyles1.Append(cellStyle28);
            cellStyles1.Append(cellStyle29);
            cellStyles1.Append(cellStyle30);
            cellStyles1.Append(cellStyle31);
            cellStyles1.Append(cellStyle32);
            cellStyles1.Append(cellStyle33);
            cellStyles1.Append(cellStyle34);
            cellStyles1.Append(cellStyle35);
            cellStyles1.Append(cellStyle36);
            cellStyles1.Append(cellStyle37);
            cellStyles1.Append(cellStyle38);
            cellStyles1.Append(cellStyle39);
            cellStyles1.Append(cellStyle40);
            cellStyles1.Append(cellStyle41);
            cellStyles1.Append(cellStyle42);
            cellStyles1.Append(cellStyle43);
            cellStyles1.Append(cellStyle44);
            cellStyles1.Append(cellStyle45);
            cellStyles1.Append(cellStyle46);
            cellStyles1.Append(cellStyle47);
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            Colors colors1 = new Colors();

            IndexedColors indexedColors1 = new IndexedColors();
            RgbColor rgbColor1 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor2 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor3 = new RgbColor() { Rgb = "00FF0000" };
            RgbColor rgbColor4 = new RgbColor() { Rgb = "0000FF00" };
            RgbColor rgbColor5 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor6 = new RgbColor() { Rgb = "00FFFF00" };
            RgbColor rgbColor7 = new RgbColor() { Rgb = "00FF00FF" };
            RgbColor rgbColor8 = new RgbColor() { Rgb = "0000FFFF" };
            RgbColor rgbColor9 = new RgbColor() { Rgb = "00000000" };
            RgbColor rgbColor10 = new RgbColor() { Rgb = "00FFFFFF" };
            RgbColor rgbColor11 = new RgbColor() { Rgb = "00FF0000" };
            RgbColor rgbColor12 = new RgbColor() { Rgb = "0000FF00" };
            RgbColor rgbColor13 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor14 = new RgbColor() { Rgb = "00FFFF00" };
            RgbColor rgbColor15 = new RgbColor() { Rgb = "00FF00FF" };
            RgbColor rgbColor16 = new RgbColor() { Rgb = "0000FFFF" };
            RgbColor rgbColor17 = new RgbColor() { Rgb = "00800000" };
            RgbColor rgbColor18 = new RgbColor() { Rgb = "00008000" };
            RgbColor rgbColor19 = new RgbColor() { Rgb = "00000080" };
            RgbColor rgbColor20 = new RgbColor() { Rgb = "00808000" };
            RgbColor rgbColor21 = new RgbColor() { Rgb = "00800080" };
            RgbColor rgbColor22 = new RgbColor() { Rgb = "00008080" };
            RgbColor rgbColor23 = new RgbColor() { Rgb = "00C0C0C0" };
            RgbColor rgbColor24 = new RgbColor() { Rgb = "00808080" };
            RgbColor rgbColor25 = new RgbColor() { Rgb = "009999FF" };
            RgbColor rgbColor26 = new RgbColor() { Rgb = "00993366" };
            RgbColor rgbColor27 = new RgbColor() { Rgb = "00FFFFCC" };
            RgbColor rgbColor28 = new RgbColor() { Rgb = "00CCFFFF" };
            RgbColor rgbColor29 = new RgbColor() { Rgb = "00660066" };
            RgbColor rgbColor30 = new RgbColor() { Rgb = "00FF8080" };
            RgbColor rgbColor31 = new RgbColor() { Rgb = "000066CC" };
            RgbColor rgbColor32 = new RgbColor() { Rgb = "00CCCCFF" };
            RgbColor rgbColor33 = new RgbColor() { Rgb = "00000080" };
            RgbColor rgbColor34 = new RgbColor() { Rgb = "00FF00FF" };
            RgbColor rgbColor35 = new RgbColor() { Rgb = "00FFFF00" };
            RgbColor rgbColor36 = new RgbColor() { Rgb = "0000FFFF" };
            RgbColor rgbColor37 = new RgbColor() { Rgb = "00800080" };
            RgbColor rgbColor38 = new RgbColor() { Rgb = "00800000" };
            RgbColor rgbColor39 = new RgbColor() { Rgb = "00008080" };
            RgbColor rgbColor40 = new RgbColor() { Rgb = "000000FF" };
            RgbColor rgbColor41 = new RgbColor() { Rgb = "0000CCFF" };
            RgbColor rgbColor42 = new RgbColor() { Rgb = "00CCFFFF" };
            RgbColor rgbColor43 = new RgbColor() { Rgb = "00CCFFCC" };
            RgbColor rgbColor44 = new RgbColor() { Rgb = "00FFFF99" };
            RgbColor rgbColor45 = new RgbColor() { Rgb = "0099CCFF" };
            RgbColor rgbColor46 = new RgbColor() { Rgb = "00FF99CC" };
            RgbColor rgbColor47 = new RgbColor() { Rgb = "00CC99FF" };
            RgbColor rgbColor48 = new RgbColor() { Rgb = "00FFCC99" };
            RgbColor rgbColor49 = new RgbColor() { Rgb = "003366FF" };
            RgbColor rgbColor50 = new RgbColor() { Rgb = "0033CCCC" };
            RgbColor rgbColor51 = new RgbColor() { Rgb = "0099CC00" };
            RgbColor rgbColor52 = new RgbColor() { Rgb = "00FFCC00" };
            RgbColor rgbColor53 = new RgbColor() { Rgb = "00FF9900" };
            RgbColor rgbColor54 = new RgbColor() { Rgb = "00FF6600" };
            RgbColor rgbColor55 = new RgbColor() { Rgb = "00666699" };
            RgbColor rgbColor56 = new RgbColor() { Rgb = "00969696" };
            RgbColor rgbColor57 = new RgbColor() { Rgb = "00003366" };
            RgbColor rgbColor58 = new RgbColor() { Rgb = "00339966" };
            RgbColor rgbColor59 = new RgbColor() { Rgb = "00003300" };
            RgbColor rgbColor60 = new RgbColor() { Rgb = "00333300" };
            RgbColor rgbColor61 = new RgbColor() { Rgb = "00993300" };
            RgbColor rgbColor62 = new RgbColor() { Rgb = "00993366" };
            RgbColor rgbColor63 = new RgbColor() { Rgb = "00333399" };
            RgbColor rgbColor64 = new RgbColor() { Rgb = "00333333" };

            indexedColors1.Append(rgbColor1);
            indexedColors1.Append(rgbColor2);
            indexedColors1.Append(rgbColor3);
            indexedColors1.Append(rgbColor4);
            indexedColors1.Append(rgbColor5);
            indexedColors1.Append(rgbColor6);
            indexedColors1.Append(rgbColor7);
            indexedColors1.Append(rgbColor8);
            indexedColors1.Append(rgbColor9);
            indexedColors1.Append(rgbColor10);
            indexedColors1.Append(rgbColor11);
            indexedColors1.Append(rgbColor12);
            indexedColors1.Append(rgbColor13);
            indexedColors1.Append(rgbColor14);
            indexedColors1.Append(rgbColor15);
            indexedColors1.Append(rgbColor16);
            indexedColors1.Append(rgbColor17);
            indexedColors1.Append(rgbColor18);
            indexedColors1.Append(rgbColor19);
            indexedColors1.Append(rgbColor20);
            indexedColors1.Append(rgbColor21);
            indexedColors1.Append(rgbColor22);
            indexedColors1.Append(rgbColor23);
            indexedColors1.Append(rgbColor24);
            indexedColors1.Append(rgbColor25);
            indexedColors1.Append(rgbColor26);
            indexedColors1.Append(rgbColor27);
            indexedColors1.Append(rgbColor28);
            indexedColors1.Append(rgbColor29);
            indexedColors1.Append(rgbColor30);
            indexedColors1.Append(rgbColor31);
            indexedColors1.Append(rgbColor32);
            indexedColors1.Append(rgbColor33);
            indexedColors1.Append(rgbColor34);
            indexedColors1.Append(rgbColor35);
            indexedColors1.Append(rgbColor36);
            indexedColors1.Append(rgbColor37);
            indexedColors1.Append(rgbColor38);
            indexedColors1.Append(rgbColor39);
            indexedColors1.Append(rgbColor40);
            indexedColors1.Append(rgbColor41);
            indexedColors1.Append(rgbColor42);
            indexedColors1.Append(rgbColor43);
            indexedColors1.Append(rgbColor44);
            indexedColors1.Append(rgbColor45);
            indexedColors1.Append(rgbColor46);
            indexedColors1.Append(rgbColor47);
            indexedColors1.Append(rgbColor48);
            indexedColors1.Append(rgbColor49);
            indexedColors1.Append(rgbColor50);
            indexedColors1.Append(rgbColor51);
            indexedColors1.Append(rgbColor52);
            indexedColors1.Append(rgbColor53);
            indexedColors1.Append(rgbColor54);
            indexedColors1.Append(rgbColor55);
            indexedColors1.Append(rgbColor56);
            indexedColors1.Append(rgbColor57);
            indexedColors1.Append(rgbColor58);
            indexedColors1.Append(rgbColor59);
            indexedColors1.Append(rgbColor60);
            indexedColors1.Append(rgbColor61);
            indexedColors1.Append(rgbColor62);
            indexedColors1.Append(rgbColor63);
            indexedColors1.Append(rgbColor64);

            colors1.Append(indexedColors1);

            stylesheet1.Append(numberingFormats1);
            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(colors1);
            return stylesheet1;
        }

        public static Stylesheet GenerateStyleSheetForExportToSite()
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)6U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)0U, Tint = -0.249977111117893D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF00FFFF" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFF66CC" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill6.Append(foregroundColor4);
            patternFill6.Append(backgroundColor4);

            fill6.Append(patternFill6);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)6U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            Colors colors1 = new Colors();

            MruColors mruColors1 = new MruColors();
            Color color2 = new Color() { Rgb = "FFFF66CC" };
            Color color3 = new Color() { Rgb = "FF00FFFF" };
            Color color4 = new Color() { Rgb = "FF33CCFF" };
            Color color5 = new Color() { Rgb = "FFFF9933" };
            Color color6 = new Color() { Rgb = "FFFF6600" };

            mruColors1.Append(color2);
            mruColors1.Append(color3);
            mruColors1.Append(color4);
            mruColors1.Append(color5);
            mruColors1.Append(color6);

            colors1.Append(mruColors1);

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(colors1);
            stylesheet1.Append(stylesheetExtensionList1);


            return stylesheet1;
        }
    }
}
