using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using Tyreshop.Utils;
using Tyreshop.DbAccess;

namespace Tyreshop.Utils
{
    class XlsxExport
    {
        private static List<DGSaleItems> lst2;
        public static void ExportToSite(List<BdProducts> list, SaveFileDialog sfd) {
            string fileName = sfd.FileName;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbpart = doc.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
                ExtendedFilePropertiesPart propertiesPart = doc.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                XlsxRutine.FilePropertiesGeneration(propertiesPart);
                SharedStringTablePart shareStringPart;
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                WorksheetPart wspart = wbpart.AddNewPart<WorksheetPart>();
                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";
                wspart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
                wspart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                wspart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                wspart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                wspart.Worksheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
                wspart.Worksheet.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
                wspart.Worksheet.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
                wspart.Worksheet.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0000-000000000000}"));
                SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:AMK6" };
                //добавление стилей
                WorkbookStylesPart wbStyles = wbpart.AddNewPart<WorkbookStylesPart>();
                wbStyles.Stylesheet = XlsxRutine.GenerateStyleSheetForExport2();// GenerateStylesShort();
                wbStyles.Stylesheet.Save();
                //wbpart.AddNewPart<ThemePart>("rId4");
                //GenerateThemePart1Content(wbpart.ThemePart);
                //добавляем лист в документ
                Sheets sheets = wbpart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = wbpart.GetIdOfPart(wspart), SheetId = 1U, Name = "Лист 1" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
                sheets.Append(sheet);
                SheetViews sheetViews1 = new SheetViews();

                SheetView sheetView1 = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
                sheetViews1.Append(sheetView1);
                SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 12.8D, DyDescent = 0.25D };
                Columns columns1 = new Columns();
                columns1.Append(new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 16D, Style = (UInt32Value)6U, CustomWidth = true });//a
                columns1.Append(new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 19.4285714285714D, Style = (UInt32Value)2U, CustomWidth = true });//b
                columns1.Append(new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 13.4285714285714D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//c
                columns1.Append(new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 15.2857142857143D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true });//d
                columns1.Append(new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 28.7142857142857D, Style = (UInt32Value)1U, BestFit = true, CustomWidth = true });//e
                columns1.Append(new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 5.71428571428571D, Style = (UInt32Value)2U, CustomWidth = true });//f
                columns1.Append(new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 4.42857142857143D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//g
                columns1.Append(new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 5.28571428571429D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//h
                columns1.Append(new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 7.85714285714286D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//i
                columns1.Append(new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 3.57142857142857D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//j
                columns1.Append(new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 7.71428571428571D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//k
                columns1.Append(new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 5.57142857142857D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//l
                columns1.Append(new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)13U, Width = 4.85714285714286D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//m
                columns1.Append(new Column() { Min = (UInt32Value)14U, Max = (UInt32Value)14U, Width = 12.4285714285714D, Style = (UInt32Value)2U, CustomWidth = true });//n
                columns1.Append(new Column() { Min = (UInt32Value)15U, Max = (UInt32Value)15U, Width = 14.8571428571429D, Style = (UInt32Value)10U, CustomWidth = true });//o
                columns1.Append(new Column() { Min = (UInt32Value)16U, Max = (UInt32Value)16U, Width = 16D, Style = (UInt32Value)10U, CustomWidth = true });//p
                columns1.Append(new Column() { Min = (UInt32Value)17U, Max = (UInt32Value)31U, Width = 3.71428571428571D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//q
                columns1.Append(new Column() { Min = (UInt32Value)32U, Max = (UInt32Value)32U, Width = 7.28571428571429D, Style = (UInt32Value)12U, CustomWidth = true });//r
                columns1.Append(new Column() { Min = (UInt32Value)33U, Max = (UInt32Value)38U, Width = 3.71428571428571D, Style = (UInt32Value)2U, BestFit = true, CustomWidth = true });//s

                SheetData sheetData = new SheetData();
                List<string> headerRow = new List<string>() {"Сезон", "Артикул", "Размеры", "Производитель", "Модель", "Шир", "Выс", "Рад", "ИН", "ИС", "RFT", "ШИП", "Груз", "Кол", "Цена",
                    "Оптовая цена", "Контейнер 7", "Контейнер 6", "Контейнер 5", "Контейнер 4","Контейнер 3","Контейнер 2","Контейнер 1","Задний Бокс","Витрины","Улица","Петроспирт контейнер 1",
                                        //29        28              27              26              25          24              23          22              21      30          31
                    "Петроспирт контейнер 2","Петроспирт контейнер 3", "Петроспирт Ангар Левый контик","Петроспирт Ангар Левый Бокс","Петроспирт Ангар Центральный контейнер",
                        //32                    33                          34                                  35                              36
                    "Петроспирт Ангар Правый бокс","Петроспирт Ангар Правый контейнер","Гараж таллинское шоссе", "Фрунзенская контейнер","Петроспирт Контейнер 4","Будка" };
                            //37                        38                                  39                      40                          41                   42
                uint rowInd = 1U;
                int cellNum = 1;
                Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D, Height=219D };
                for (int i=0;i<headerRow.Count; i++) {
                    string s = headerRow[i];
                    if (i <= 15)
                    {
                        if(i==14 || i == 15)
                            cellNum = XlsxRutine.InsertCellToRow(rowHead, cellNum, s, shareStringPart, 8U);
                        else
                            cellNum = XlsxRutine.InsertCellToRow(rowHead, cellNum, s, shareStringPart, 7U);
                    }
                    else
                        cellNum = XlsxRutine.InsertCellToRow(rowHead, cellNum, s, shareStringPart, 14U);
                }
                sheetData.AppendChild(rowHead);
                rowInd++;
                cellNum = 1;
                foreach (var prod in list) {
                    string sizes = prod.Width + ", " + prod.Height + ", " + prod.Radius;
                    Row prodRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Season, shareStringPart, 5U);//a
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Articul, shareStringPart, 3U);//b
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, sizes, shareStringPart,3U);//c
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Manufacturer, shareStringPart, 4U);//d
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Model, shareStringPart, 4U);//e
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Width.ToString(), shareStringPart, 3U);//f
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Height.ToString(), shareStringPart, 3U);//g
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Radius, shareStringPart, 3U);//h
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.InCol, shareStringPart, 3U);//i
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.IsCol, shareStringPart, 3U);//j
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.RFT, shareStringPart, 3U);//k
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Spikes, shareStringPart, 3U);//l
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Gruz, shareStringPart, 3U);//m
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.TotalQuantity.ToString(), shareStringPart, 3U);//n
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Price.ToString(), shareStringPart, 9U);//o
                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.OptPrice.ToString(), shareStringPart, 9U);//p
                    var store = prod.Storehouse;
                    if(store.Exists(w => w.StorehouseId == 29 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w=>w.StorehouseId==29 && w.ProductId==prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 28 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 28 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 27 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 27 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 26 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 26 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 25 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 25 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 24 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 24 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 23 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 23 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q
                    if (store.Exists(w => w.StorehouseId == 22 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 22 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);//q if(prod.Storehouse.Single(w => w.StorehouseId == 29 && w.ProductId == prod.ProductId).Quantity.ToString()!=string.Empty)
                    if (store.Exists(w => w.StorehouseId == 21 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 21 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 30 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 30 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 31 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 31 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 32 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 32 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 33 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 33 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 34 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 34 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 35 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 35 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 36 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 36 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 37 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 37 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 38 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 38 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 39 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 39 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 40 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 40 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 41 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 41 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    if (store.Exists(w => w.StorehouseId == 42 && w.ProductId == prod.ProductId))
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 42 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//q
                    else
                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 28 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//r
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 27 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//s
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 26 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//t
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 25 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//u
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 24 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//v
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 23 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//w
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 22 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//x
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 21 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//y
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 30 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//z
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 31 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//aa
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 32 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ab
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 33 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ac
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 34 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ad
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 35 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ae
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 36 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//af
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 37 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ag
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 38 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ah
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 39 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ai
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 40 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//aj
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 41 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//ak
                    //cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, prod.Storehouse.Single(w => w.StorehouseId == 42 && w.ProductId == prod.ProductId).Quantity.ToString(), shareStringPart, 3U);//al
                    //foreach (var item in prod.Storehouse)
                    //{

                    //}
                    sheetData.AppendChild(prodRow);
                    rowInd++;
                    cellNum = 1;
                }
                PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
                PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
                wspart.Worksheet.Append(sheetDimension1);
                wspart.Worksheet.Append(sheetViews1);
                wspart.Worksheet.Append(sheetFormatProperties1);
                wspart.Worksheet.Append(columns1);
                wspart.Worksheet.Append(sheetData);
                var autoFilter = new AutoFilter() { Reference = "A:AL" };
                wspart.Worksheet.Append(autoFilter);
                wspart.Worksheet.Append(pageMargins1);
                wspart.Worksheet.Append(pageSetup1);
                wbpart.Workbook.Save();
                doc.Close();
            }
            int counter = 0;
            do
            {
                System.Threading.Thread.Sleep(500);
                counter++;
            } while (!File.Exists(fileName) && counter < 10);
            if (counter < 10)
                System.Diagnostics.Process.Start(fileName);
        }

        public static void ExportDailyReport(SaveFileDialog sfd, DateTime? date, DateTime? date2 = null) {
            string fileName = sfd.FileName;

            List<operation> list2 = new List<operation>();
            List<user> users = new List<user>();
            using (u0324292_tyreshopEntities db = new u0324292_tyreshopEntities()) {
                var list = db.products.Select(s => new PComboBox()
                {
                    ProductId = s.ProductId,
                    ProductName = @"" + db.manufacturers.Where(w => w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ManufacturerName).FirstOrDefault() + " " +
                        db.models.Where(w => w.ModelId == s.ModelId && w.ManufacturerId == s.ManufacturerId).Select(sel => sel.ModelName).FirstOrDefault() + " " + s.Width + " / " + s.Height + " / " + s.Radius
                }).ToList();
                users = db.users.Where(w => w.Role == "manager").ToList();
                if (date2 == null)
                {
                    list2 = db.operations.Where(w => w.OperationDate == date).ToList();
                }
                else {
                    list2 = db.operations.Where(w => w.OperationDate >= date && w.OperationDate <= date2).ToList();
                }
                lst2 = new List<DGSaleItems>();
                foreach (var oper in list2)
                {
                    string uName = "";
                    if (users.Exists(e => e.UserId == oper.UserId))
                        uName = users.Single(s => s.UserId == oper.UserId).UserName;
                    else
                        uName = users.Single(s => s.UserId == 14).UserName;
                    DGSaleItems dg = new DGSaleItems();
                    dg.Comment = oper.Comment;
                    dg.Date = oper.OperationDate.ToString();
                    dg.OperationType = oper.OperationType;
                    dg.PayType = oper.PayType;
                    dg.Price = oper.Price;
                    if (oper.ProductId != null && oper.ProductId != 0)
                    {
                        if (list.Exists(e => e.ProductId == oper.ProductId))
                            dg.ProdName = list.Where(w => w.ProductId == oper.ProductId).Select(s => s.ProductName).First();
                        else
                            continue;
                    }
                    else if (oper.ServiceId != null && oper.ServiceId != 0)
                    {
                        if (db.services.ToList().Exists(e => e.ServiceId == oper.ServiceId))
                            dg.ProdName = db.services.Where(w => w.ServiceId == oper.ServiceId).Select(s => s.ServiceName).First();
                        else
                            continue;
                    }
                    dg.Time = oper.OperationTime.ToString();
                    dg.Price = oper.Price;
                    dg.Quantity = oper.Count;
                    dg.StoreHouse = oper.Storehouse;
                    dg.SaleNumber = (long)oper.SaleNumber;
                    dg.CardPayed = oper.CardPay;
                    dg.CardToTotalSum = oper.CardToTotalSum;
                    dg.UserName = uName;
                    dg.UserId = oper.UserId;
                    lst2.Add(dg);
                }
                //lst2.Add(new DGSaleItems() { ProdName = "Сдача наличных", Price = totalSum });
            }
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbpart = doc.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
                ExtendedFilePropertiesPart propertiesPart = doc.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                XlsxRutine.FilePropertiesGeneration(propertiesPart);
                SharedStringTablePart shareStringPart;
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                WorksheetPart wspart = wbpart.AddNewPart<WorksheetPart>();
                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";
                wspart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac xr xr2 xr3" } };
                wspart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                wspart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                wspart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                wspart.Worksheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
                wspart.Worksheet.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
                wspart.Worksheet.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
                wspart.Worksheet.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0000-000000000000}"));
                SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:AMK6" };
                //добавление стилей
                WorkbookStylesPart wbStyles = wbpart.AddNewPart<WorkbookStylesPart>();
                wbStyles.Stylesheet = XlsxRutine.GenerateStyleSheetForExportToSite();// GenerateStylesShort();
                wbStyles.Stylesheet.Save();
                //wbpart.AddNewPart<ThemePart>("rId4");
                //GenerateThemePart1Content(wbpart.ThemePart);
                //добавляем лист в документ
                Sheets sheets = wbpart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = wbpart.GetIdOfPart(wspart), SheetId = 1U, Name = "Лист 1" };//имя можно задать любое, единственное ограничение - количество символов не должно превышать 40, хотя в документации указан предел в 255 символов, мистика одним словом
                sheets.Append(sheet);
                SheetViews sheetViews1 = new SheetViews();

                SheetView sheetView1 = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
                sheetViews1.Append(sheetView1);
                SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 12.8D, DyDescent = 0.25D };
                Columns columns1 = new Columns();
                columns1.Append(new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 30D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 25D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 20D, BestFit = true, CustomWidth = true });
                columns1.Append(new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 20D, BestFit = true, CustomWidth = true });
                //columns1.Append(new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 5.5703125D, BestFit = true, CustomWidth = true });
                //columns1.Append(new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)13U, Width = 4.85546875D, BestFit = true, CustomWidth = true });
                //columns1.Append(new Column() { Min = (UInt32Value)14U, Max = (UInt32Value)14U, Width = 4.42578125D, BestFit = true, CustomWidth = true });
                //columns1.Append(new Column() { Min = (UInt32Value)15U, Max = (UInt32Value)15U, Width = 6D, BestFit = true, CustomWidth = true });
                //columns1.Append(new Column() { Min = (UInt32Value)16U, Max = (UInt32Value)16U, Width = 13.5703125D, BestFit = true, CustomWidth = true });

                SheetData sheetData = new SheetData();
                string[] headerRowSale = new string[] { "Продажи", "", "", "", "", "", "", "", "", "" };
                string[] headerRowServ = new string[] { "Услуги", "", "", "", "", "", "", "", "", "" };
                string[] headerRow = new string[] { "Номер операции", "Время", "Тип", "Товар/Услуга", "Количество", "Склад", "Стоимость", "Тип платежа", "Платеж на карту", "Комментарий", "Менеджер" };
                uint rowInd = 1U;
                int cellNum = 1;
                Row saleRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRowSale)
                {
                    cellNum = XlsxRutine.InsertCellToRow(saleRow, cellNum, s, shareStringPart);
                }
                sheetData.AppendChild(saleRow);
                rowInd++;
                cellNum = 1;
                Row rowHead = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRow)
                {
                    cellNum = XlsxRutine.InsertCellToRow(rowHead, cellNum, s, shareStringPart, 2U);
                }
                sheetData.AppendChild(rowHead);
                rowInd++;
                cellNum = 1;
                var lst3 = lst2.Where(w => w.OperationType != "Услуга").ToList();
                for (int i=0;i<lst3.Count;i++)//foreach (var oper in lst2)
                {
                    var oper = lst3[i];
                    Row prodRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    switch (oper.OperationType)
                    {
                        case "Продажа":
                            if (oper.PayType == "Безналичный расчет")
                            {
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Time, shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.OperationType, shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.ProdName, shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Quantity.ToString(), shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.StoreHouse, shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Price.ToString(), shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.PayType, shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.CardPayed, shareStringPart, 3U);
                                if (oper.Comment != string.Empty && oper.Comment!=null)
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Comment, shareStringPart, 3U);
                                else
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 3U);
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.UserName, shareStringPart, 3U);
                            }
                            else
                            {
                                if (oper.CardToTotalSum == 0)
                                {
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Time, shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.OperationType, shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.ProdName, shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Quantity.ToString(), shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.StoreHouse, shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Price.ToString(), shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.PayType, shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.CardPayed, shareStringPart, 5U);
                                    if (oper.Comment != string.Empty && oper.Comment != null)
                                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Comment, shareStringPart, 5U);
                                    else
                                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 5U);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.UserName, shareStringPart, 5U);
                                }
                                else
                                {
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.SaleNumber.ToString(), shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Time, shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.OperationType, shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.ProdName, shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Quantity.ToString(), shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.StoreHouse, shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Price.ToString(), shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.PayType, shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.CardPayed, shareStringPart);
                                    if (oper.Comment != string.Empty && oper.Comment != null)
                                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Comment, shareStringPart);
                                    else
                                        cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart);
                                    cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.UserName, shareStringPart);
                                }
                            }
                            break;
                        case "Списание товара":
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Time, shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.OperationType, shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.ProdName, shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Quantity.ToString(), shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.StoreHouse, shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Price.ToString(), shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.PayType, shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.CardPayed, shareStringPart, 1U);
                            if (oper.Comment != string.Empty && oper.Comment != null)
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Comment, shareStringPart, 1U);
                            else
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 1U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.UserName, shareStringPart, 1U);
                            break;
                        case "Списание наличных":
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Time, shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.OperationType, shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Price.ToString(), shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.PayType, shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.CardPayed, shareStringPart, 2U);
                            if (oper.Comment != string.Empty && oper.Comment != null)
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.Comment, shareStringPart, 2U);
                            else
                                cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, "", shareStringPart, 2U);
                            cellNum = XlsxRutine.InsertCellToRow(prodRow, cellNum, oper.UserName, shareStringPart, 2U);
                            break;
                    }
                    sheetData.AppendChild(prodRow);
                    rowInd++;
                    cellNum = 1;

                }
                var listLast = new List<DGSaleItems>();
                decimal totalSum = 0;
                var listius = lst2.Where(w => w.OperationType != "Услуга").ToList();
                foreach (var item in listius)
                {
                    if (item.PayType == "Наличный расчет" && item.OperationType != "Списание наличных")
                    {
                        if (item.CardPayed == "Да")
                        {
                            if (item.CardToTotalSum != 0)
                                totalSum += item.Price;
                        }
                        else
                            totalSum += item.Price;
                    }
                    if (item.OperationType == "Списание наличных")
                        totalSum -= item.Price;
                }
                listLast.Add(new DGSaleItems() { ProdName = "Сдача налички", Price = totalSum });
                foreach (var operLast in listLast)
                {
                    //var operLast = lst2[lst2.Count - 1];


                    Row lastRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, operLast.ProdName, shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, operLast.Price.ToString(), shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    sheetData.AppendChild(lastRow);
                    rowInd++;
                    cellNum = 1;
                }
                Row emptyRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRow)
                {
                    cellNum = XlsxRutine.InsertCellToRow(emptyRow, cellNum, "", shareStringPart);
                }
                sheetData.AppendChild(emptyRow);
                rowInd++;
                cellNum = 1;
                uint saleEnd = rowInd;
                Row servRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRowServ)
                {
                    cellNum = XlsxRutine.InsertCellToRow(servRow, cellNum, s, shareStringPart);
                }
                sheetData.AppendChild(servRow);
                rowInd++;
                cellNum = 1;
                Row rowHead2 = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRow)
                {
                    cellNum = XlsxRutine.InsertCellToRow(rowHead2, cellNum, s, shareStringPart, 2U);
                }
                sheetData.AppendChild(rowHead2);
                rowInd++;
                cellNum = 1;
                lst3 = lst2.Where(w => w.OperationType == "Услуга").ToList();
                for (int i = 0; i < lst3.Count; i++)//foreach (var oper in lst2)
                {
                    var oper = lst3[i];
                    Row serviceRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    if (oper.PayType == "Безналичный расчет")
                    {
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Time, shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.OperationType, shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.ProdName, shareStringPart,2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Quantity.ToString(), shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.StoreHouse, shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Price.ToString(), shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.PayType, shareStringPart, 2U);
                        cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.CardPayed, shareStringPart, 2U);
                        if (oper.Comment != string.Empty && oper.Comment != null)
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Comment, shareStringPart, 2U);
                        else
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, "", shareStringPart, 2U);
                    }
                    else
                    {
                        if (oper.CardToTotalSum == 0)
                        {
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.SaleNumber.ToString(), shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Time, shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.OperationType, shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.ProdName, shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Quantity.ToString(), shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.StoreHouse, shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Price.ToString(), shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.PayType, shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.CardPayed, shareStringPart, 5U);
                            if (oper.Comment != string.Empty && oper.Comment != null)
                                cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Comment, shareStringPart, 5U);
                            else
                                cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, "", shareStringPart, 5U);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.UserName, shareStringPart, 5U);
                        }
                        else
                        {
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.SaleNumber.ToString(), shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Time, shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.OperationType, shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.ProdName, shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Quantity.ToString(), shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.StoreHouse, shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Price.ToString(), shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.PayType, shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.CardPayed, shareStringPart);
                            if (oper.Comment != string.Empty && oper.Comment != null)
                                cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.Comment, shareStringPart);
                            else
                                cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, "", shareStringPart);
                            cellNum = XlsxRutine.InsertCellToRow(serviceRow, cellNum, oper.UserName, shareStringPart);
                        }
                    }
                    
                    sheetData.AppendChild(serviceRow);
                    rowInd++;
                    cellNum = 1;
                }
                Row emptyRow2 = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRow)
                {
                    cellNum = XlsxRutine.InsertCellToRow(emptyRow2, cellNum, "", shareStringPart);
                }
                sheetData.AppendChild(emptyRow2);
                rowInd++;
                cellNum = 1;
                listius = lst2.Where(w => w.OperationType == "Услуга").ToList();
                totalSum = 0;
                listLast.Clear();
                foreach (var item in listius)
                {
                    if (item.PayType == "Наличный расчет" && item.OperationType != "Списание наличных")
                    {
                        if (item.CardPayed == "Да")
                        {
                            if (item.CardToTotalSum != 0)
                                totalSum += item.Price;
                        }
                        else
                            totalSum += item.Price;
                    }

                }
                listLast.Add(new DGSaleItems() { ProdName = "Сумма по услугам (Наличка)", Price = totalSum });
                totalSum = 0;
                foreach (var item in listius) {
                    totalSum += item.Price;
                }
                listLast.Add(new DGSaleItems() { ProdName = "Сумма по услугам (Все)", Price = totalSum });
                //totalSum = 0;
                
                foreach (var operLast in listLast)
                {
                    //var operLast = lst2[lst2.Count - 1];


                    Row lastRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, operLast.ProdName, shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, operLast.Price.ToString(), shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(lastRow, cellNum, "", shareStringPart);
                    sheetData.AppendChild(lastRow);
                    rowInd++;
                    cellNum = 1;
                }
                Row emptyRow3 = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                foreach (string s in headerRow)
                {
                    cellNum = XlsxRutine.InsertCellToRow(emptyRow3, cellNum, "", shareStringPart);
                }
                sheetData.AppendChild(emptyRow3);
                rowInd++;
                cellNum = 1;
                var lst4 = lst2.Where(w => w.OperationType != "Услуга").ToList();
                
                foreach (var user in users) {
                    int totalQuant = 0;
                    var id = user.UserId;
                    var tmp = lst4.Where(w => w.UserId == id).Select(s => s.Quantity).ToList();
                    foreach (var t in tmp) {
                        totalQuant += t;
                    }
                    Row userRow = new Row() { RowIndex = rowInd, CustomHeight = true, DyDescent = 0.25D };
                    cellNum = XlsxRutine.InsertCellToRow(userRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(userRow, cellNum, "", shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(userRow, cellNum, user.UserName, shareStringPart);
                    cellNum = XlsxRutine.InsertCellToRow(userRow, cellNum, totalQuant.ToString(), shareStringPart);
                    sheetData.AppendChild(userRow);
                    rowInd++;
                    cellNum = 1;
                }
                PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
                PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, FirstPageNumber = (UInt32Value)0U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U };
                wspart.Worksheet.Append(sheetDimension1);
                wspart.Worksheet.Append(sheetViews1);
                wspart.Worksheet.Append(sheetFormatProperties1);
                wspart.Worksheet.Append(columns1);
                wspart.Worksheet.Append(sheetData);
                MergeCells mergeCells1 = new MergeCells();
                mergeCells1.Append(new MergeCell() { Reference = "A1:J1" });
                mergeCells1.Append(new MergeCell() { Reference = "A"+saleEnd+":J"+saleEnd });
                wspart.Worksheet.Append(mergeCells1);
                wspart.Worksheet.Append(pageMargins1);
                wspart.Worksheet.Append(pageSetup1);
                wbpart.Workbook.Save();
                doc.Close();
            }
            int counter = 0;
            do
            {
                System.Threading.Thread.Sleep(500);
                counter++;
            } while (!File.Exists(fileName) && counter < 10);
            if (counter < 10)
                System.Diagnostics.Process.Start(fileName);
        }
    }
}
