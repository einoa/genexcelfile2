using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging; // add reference WindowsBase
using System.Xml;


namespace GenExcelFile2.Models
{
    public class Excel
    {

        public static System.IO.MemoryStream TestFile()
        {
            /* ************************************************************************************************************************
             *  We will generate 4 XML documents.
             *   1) /xl/styles.xml
             *   2) /xl/workbook.xml
             *   3) /xl/worksheets/sheet1.xml
             *   4) /xl/sharedStrings.xml
             * ************************************************************************************************************************/

            System.IO.MemoryStream ms = new MemoryStream();
            Package pkg = Package.Open(ms, FileMode.Create);

            System.Collections.Generic.List<string> sharedStrings = new List<string>(); // *** sharedStrings


            const string nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string nsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            const string contentTypeMain = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
            const string contentTypeWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
            const string contentTypeStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
            const string contentTypeSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";


            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // styles.xml
            // ************************************************************************************************************************
            // ************************************************************************************************************************

            // Stylesheet document
            XmlDocument stylesheetXMLDoc = new XmlDocument();

            // ************************************************************************************************************************
            // Stylesheet root
            // ************************************************************************************************************************
            /*
             *   <stylesheet>
             */
            XmlElement xmlStylesheet = stylesheetXMLDoc.CreateElement("styleSheet", nsSpreadsheetML);
            stylesheetXMLDoc.AppendChild(xmlStylesheet);

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Fonts
            // ------------------------------------------------------------------------------------------------------------------------
            /*
             *   <stylesheet>
             *     <fonts>
             */
            XmlElement xmlFonts = stylesheetXMLDoc.CreateElement("fonts", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xmlFonts);

            /*
             *   < stylesheet>
             *     <fonts>
             *       <font>
             */
            XmlElement xFont = stylesheetXMLDoc.CreateElement("font", nsSpreadsheetML);
            xmlFonts.AppendChild(xFont);

            /*
             *   < stylesheet>
             *     <fonts>
             *       <font>
             *         <sz val="11">
             */
            XmlElement xSz = stylesheetXMLDoc.CreateElement("sz", nsSpreadsheetML);
            xFont.AppendChild(xSz);
            xSz.SetAttribute("val", "11");

            // Stylesheet - Fonts - Font - Val
            /*
             *   < stylesheet>
             *     <fonts>
             *       <font>
             *         <sz val="11">
             *         <name val="Calibri">
             */
            XmlElement xName = stylesheetXMLDoc.CreateElement("name", nsSpreadsheetML);
            xFont.AppendChild(xName);
            xName.SetAttribute("val", "Calibri");

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Fills
            // ------------------------------------------------------------------------------------------------------------------------
            /*
             *   <stylesheet>
             *     <fills>
             */
            XmlElement xFills = stylesheetXMLDoc.CreateElement("fills", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xFills);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills>
             *       <fill>
             */
            XmlElement xFill = stylesheetXMLDoc.CreateElement("fill", nsSpreadsheetML);
            xFills.AppendChild(xFill);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills>
             *       <fill>
             *         <patternFill patternType="none">
             */
            XmlElement xPatternFill = stylesheetXMLDoc.CreateElement("patternFill", nsSpreadsheetML);
            xFill.AppendChild(xPatternFill);
            xPatternFill.SetAttribute("patternType", "none");

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills>
             *       <fill>
             *         <patternFill patternType="none">
             *       </fill>
             *       <fill>
             *         <patternFill patternType="gray125">
             */
            xFill = stylesheetXMLDoc.CreateElement("fill", nsSpreadsheetML);
            xFills.AppendChild(xFill);

            xPatternFill = stylesheetXMLDoc.CreateElement("patternFill", nsSpreadsheetML);
            xFill.AppendChild(xPatternFill);
            xPatternFill.SetAttribute("patternType", "gray125");

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - Borders
            // ------------------------------------------------------------------------------------------------------------------------
            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders>
             */
            XmlElement xBorders = stylesheetXMLDoc.CreateElement("borders", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xBorders);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders>
             *       <border>
             */
            XmlElement xBorder = stylesheetXMLDoc.CreateElement("border", nsSpreadsheetML);
            xBorders.AppendChild(xBorder);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders>
             *       <border>
             *         <left>
             *         <right>
             *         <top>
             *         <bottom>
             *         <diagonal>
             */
            xBorder.AppendChild(stylesheetXMLDoc.CreateElement("left", nsSpreadsheetML));
            xBorder.AppendChild(stylesheetXMLDoc.CreateElement("right", nsSpreadsheetML));
            xBorder.AppendChild(stylesheetXMLDoc.CreateElement("top", nsSpreadsheetML));
            xBorder.AppendChild(stylesheetXMLDoc.CreateElement("bottom", nsSpreadsheetML));
            xBorder.AppendChild(stylesheetXMLDoc.CreateElement("diagonal", nsSpreadsheetML));

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - CellStyleXfs
            // ------------------------------------------------------------------------------------------------------------------------
            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders> ... </border>
             *     <cellStyleXfs>
             */
            XmlElement xCellStyleXfs = stylesheetXMLDoc.CreateElement("cellStyleXfs", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xCellStyleXfs);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders> ... </border>
             *     <cellStyleXfs>
             *       <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
             */
            XmlElement xXf = stylesheetXMLDoc.CreateElement("xf", nsSpreadsheetML);
            xCellStyleXfs.AppendChild(xXf);
            xXf.SetAttribute("numFmtId", "0");
            xXf.SetAttribute("fontId", "0");
            xXf.SetAttribute("fillId", "0");
            xXf.SetAttribute("borderId", "0");

            // ------------------------------------------------------------------------------------------------------------------------
            // Stylesheet - CellXfs
            // ------------------------------------------------------------------------------------------------------------------------
            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders> ... </border>
             *     <cellStyleXfs> ... </cellStyleXfs>
             *     <cellXfs>
             */
            XmlElement xCellXfs = stylesheetXMLDoc.CreateElement("cellXfs", nsSpreadsheetML);
            xmlStylesheet.AppendChild(xCellXfs);

            /*
             *   <stylesheet>
             *     <fonts> ... </fonts>
             *     <fills> ... </fills>
             *     <borders> ... </border>
             *     <cellStyleXfs> ... </cellStyleXfs>
             *     <cellXfs>
             *       <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
             */
            xXf = stylesheetXMLDoc.CreateElement("xf", nsSpreadsheetML);
            xCellXfs.AppendChild(xXf);
            xXf.SetAttribute("numFmtId", "0");
            xXf.SetAttribute("fontId", "0");
            xXf.SetAttribute("fillId", "0");
            xXf.SetAttribute("borderId", "0");
            xXf.SetAttribute("xfId", "0");

            // ************************************************************************************************************************
            // WRITE styles.xml
            // ************************************************************************************************************************

            Uri uriStylesheet = new Uri("/xl/styles.xml", UriKind.Relative);
            PackagePart ppStylesheet = pkg.CreatePart(uriStylesheet, contentTypeStyles);
            StreamWriter swStylesheet = new StreamWriter(ppStylesheet.GetStream(FileMode.Create, FileAccess.Write));
            stylesheetXMLDoc.Save(swStylesheet);
            swStylesheet.Close();
            pkg.Flush();



            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // Worksheet document
            // ************************************************************************************************************************
            // ************************************************************************************************************************
            XmlDocument xmlWorksheetDoc = new XmlDocument();

            // Worksheet root
            /*
             *   <worksheet>
             */
            XmlElement xWorksheet = xmlWorksheetDoc.CreateElement("worksheet", nsSpreadsheetML);
            xmlWorksheetDoc.AppendChild(xWorksheet);
            xWorksheet.SetAttribute("xmlns:r", nsRelationships);

            /*
             *   <worksheet>
             *     <sheetViews>
             */
            XmlElement xSheetViews = xmlWorksheetDoc.CreateElement("sheetViews", nsSpreadsheetML);
            xWorksheet.AppendChild(xSheetViews);

            /*
             *   <worksheet>
             *     <sheetViews>
             *       <sheetView>
             */
            XmlElement xSheetView = xmlWorksheetDoc.CreateElement("sheetView", nsSpreadsheetML);
            xSheetViews.AppendChild(xSheetView);
            xSheetView.SetAttribute("workbookViewId", "0");

            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             */
            XmlElement xSheetData = xmlWorksheetDoc.CreateElement("sheetData", nsSpreadsheetML);
            xWorksheet.AppendChild(xSheetData);


            string stringvalue;
            XmlElement xRow;
            XmlElement xCol;
            XmlElement value;
            int currentRow = 0;
            int currentCol = 0;


            // row
            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             *       <row>
             */
            currentRow = 1;
            xRow = xmlWorksheetDoc.CreateElement("row", nsSpreadsheetML);
            xSheetData.AppendChild(xRow);
            xRow.SetAttribute("r", (currentRow).ToString());

            // col
            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             *       <row r="1">
             *         <c>
             */
            currentCol = 1;
            stringvalue = "abc";
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            xCol.SetAttribute("t", "s");
            xCol.SetAttribute("s", "0");
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            if (sharedStrings.Contains(stringvalue))
            {
                value.InnerText = sharedStrings.IndexOf(stringvalue).ToString();
            }
            else
            {
                value.InnerText = sharedStrings.Count.ToString();
                sharedStrings.Add(stringvalue);
            }

            // col
            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             *       <row r="1">
             *         <c> ... </c>
             *         <c>
             */
            currentCol = 2;
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            value.InnerText = "123.45";

            // row
            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             *       <row r="1"> ... </row>
             *       <row>
             */
            currentRow = 2;
            xRow = xmlWorksheetDoc.CreateElement("row", nsSpreadsheetML);
            xSheetData.AppendChild(xRow);
            xRow.SetAttribute("r", (currentRow).ToString());
            // col
            /*
             *   <worksheet>
             *     <sheetViews> ... </sheetViews>
             *     <sheetData>
             *       <row r="1"> ... </row>
             *       <row r="2">
             *         <c>
             */
            currentCol = 2;
            xCol = xmlWorksheetDoc.CreateElement("c", nsSpreadsheetML);
            xRow.AppendChild(xCol);
            xCol.SetAttribute("r", ExcelColumnName(currentCol) + currentRow.ToString());
            value = xmlWorksheetDoc.CreateElement("v", nsSpreadsheetML);
            xCol.AppendChild(value);
            value.InnerText = "99";

            // ************************************************************************************************************************
            // WRITE sheet1.xml
            // ************************************************************************************************************************
            Uri uriWorksheet = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
            PackagePart ppWorksheet = pkg.CreatePart(uriWorksheet, contentTypeWorksheet);
            StreamWriter swWorksheet = new StreamWriter(ppWorksheet.GetStream(FileMode.Create, FileAccess.Write));
            xmlWorksheetDoc.Save(swWorksheet);
            swWorksheet.Close();
            pkg.Flush();



            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // workbook.xml
            // ************************************************************************************************************************
            // ************************************************************************************************************************

            XmlDocument workbookXMLDoc = new XmlDocument();

            // Workbook root
            /*
             *   <workbook>
             */
            XmlElement xmlWorkbookRoot = workbookXMLDoc.CreateElement("workbook", nsSpreadsheetML);
            workbookXMLDoc.AppendChild(xmlWorkbookRoot);
            XmlAttribute xmlAttr = workbookXMLDoc.CreateAttribute("xmlns", "r", @"http://www.w3.org/2000/xmlns/");
            xmlAttr.Value = nsRelationships;
            xmlWorkbookRoot.Attributes.Append(xmlAttr);

            /*
             *   <workbook>
             *     <sheets>
             */
            XmlElement xSheets = workbookXMLDoc.CreateElement("sheets", nsSpreadsheetML);
            xmlWorkbookRoot.AppendChild(xSheets);

            /*
             *   <workbook>
             *     <sheets>
             *       <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
             */
            XmlElement xSheet = workbookXMLDoc.CreateElement("sheet", nsSpreadsheetML);
            xSheets.AppendChild(xSheet);
            xSheet.SetAttribute("name", "Sheet1");
            xSheet.SetAttribute("sheetId", "1");
            xSheet.SetAttribute("id", nsRelationships, "rId1");

            // ************************************************************************************************************************
            // WRITE workbook.xml
            // ************************************************************************************************************************
            Uri uriWorkbook = new Uri("/xl/workbook.xml", UriKind.Relative);
            PackagePart ppWorkbook = pkg.CreatePart(uriWorkbook, contentTypeMain);
            StreamWriter swWorkbook = new StreamWriter(ppWorkbook.GetStream(FileMode.Create, FileAccess.Write));
            workbookXMLDoc.Save(swWorkbook);
            swWorkbook.Close();
            pkg.Flush();



            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // SharedStrings document
            // ************************************************************************************************************************
            // ************************************************************************************************************************
            XmlDocument sharedStringsXMLDoc = new XmlDocument();

            // SharedStrings - Sst
            XmlElement xSst = sharedStringsXMLDoc.CreateElement("sst", nsSpreadsheetML);
            xSst.SetAttribute("count", sharedStrings.Count.ToString());
            xSst.SetAttribute("uniqueCount", sharedStrings.Count.ToString());
            sharedStringsXMLDoc.AppendChild(xSst);

            for (int i = 0; i < sharedStrings.Count; i++)
            {
                XmlElement xSi = sharedStringsXMLDoc.CreateElement("si", nsSpreadsheetML);
                XmlElement xt = sharedStringsXMLDoc.CreateElement("t", nsSpreadsheetML);
                xt.InnerText = sharedStrings[i];
                xSi.AppendChild(xt);
                xSst.AppendChild(xSi);
            }

            // ************************************************************************************************************************
            // WRITE sharedStrings.xml
            // ************************************************************************************************************************
            Uri uriStrings = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
            PackagePart ppSharedStrings = pkg.CreatePart(uriStrings, contentTypeSharedStrings);
            StreamWriter swSharedStrings = new StreamWriter(ppSharedStrings.GetStream(FileMode.Create, FileAccess.Write));
            sharedStringsXMLDoc.Save(swSharedStrings);
            swSharedStrings.Close();
            pkg.Flush();



            // ************************************************************************************************************************
            // ************************************************************************************************************************
            // WRITE relationships
            // ************************************************************************************************************************
            // ************************************************************************************************************************
            pkg.CreateRelationship(uriWorkbook, TargetMode.Internal, nsRelationships + "/officeDocument", "rId1");
            ppWorkbook.CreateRelationship(uriWorksheet, TargetMode.Internal, nsRelationships + "/worksheet", "rId1");
            ppWorkbook.CreateRelationship(uriStylesheet, TargetMode.Internal, nsRelationships + "/styles", "rId2");
            ppWorkbook.CreateRelationship(uriStrings, TargetMode.Internal, nsRelationships + "/sharedStrings", "rId3");
            pkg.Flush();



            // ************************************************************************************************************************
            // Close
            // ************************************************************************************************************************
            pkg.Close();

            ms.Position = 0;
            return ms;
        }

        /*
         *  Function: ExcelColumnName
         *     1 returns A
         *     2 returns B
         *     3 returns C
         *    26 returns Z
         *    27 returns AA
         *    28 returns AB
         *    etc.
         */
        public static string ExcelColumnName(int columnNumber)
        {
            string a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (columnNumber <= 26)
            {
                return (a[columnNumber - 1]).ToString();
            }
            else if (columnNumber <= 702)
            {
                int n1 = columnNumber - 27;
                int n2 = n1 / 26;
                int n3 = n1 - 26 * n2;

                char[] t = new char[2];
                t[0] = a[n2];
                t[1] = a[n3];
                return new string(t);
            }
            else
            {
                int n1 = columnNumber - 703;
                int n2 = n1 / 676;
                int n3 = n1 - 676 * n2;
                int n4 = n3 / 26;
                int n5 = n3 - 26 * n4;

                char[] t = new char[3];
                t[0] = a[n2];
                t[1] = a[n4];
                t[2] = a[n5];

                return new string(t);
            }
        }
    }
}