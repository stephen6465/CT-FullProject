using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace UCT.Models
{
    public class ExcelArcReportGenerator
    {
        private Dictionary<string, int> propertyToSharedStrings = new Dictionary<string, int>();
        private const string PROPERTY_KEY_ACTIVITY_POSITION_FORMAT = "Activity-{0}-Position";
        private const string PROPERTY_KEY_ACTIVITY_TITLE_FORMAT = "Activity-{0}-Title";
        private const string PROPERTY_KEY_ACTIVITY_SCENARIO_FORMAT = "Activity-{0}-Scenario";
        private const string PROPERTY_KEY_ACTIVITY_TOPICSREQUIRED_FORMAT = "Activity-{0}-TopicsRequired";
        private const string PROPERTY_KEY_ACTIVITY_WEEKS_FORMAT = "Activity-{0}-Weeks";
        private const string PROPERTY_KEY_GOAL_DESC_FORMAT = "Goal-{0}-Description";
        private const string PROPERTY_KEY_COMPETENCY_DESC_FORMAT = "Competency-{0}-Description";
        private const string PROPERTY_KEY_DESCRIPTOR_NUMBER_FORMAT = "Descriptor-{0}-Number";
        private const string PROPERTY_KEY_DESCRIPTOR_DESC_FORMAT = "Descriptor-{0}-Description";

        public string ProgramName { get; private set; }
        public string GeneratorUsername { get; private set; }

        public ExcelArcReportGenerator(string programName, string generatorUserName)
        {
            this.ProgramName = programName;
            this.GeneratorUsername = generatorUserName;
        }

        // Creates an Document instance and adds its children.
        public byte[] GenerateLearningActivityReport(List<LearningActivities_Archive> programLearningActivities)
        {
            byte[] reportBytes = null;
            using (MemoryStream stream = new MemoryStream())
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    CreateLearningActivityReportParts(package, programLearningActivities);
                }

                reportBytes = stream.ToArray();
            }

            return reportBytes;
        }

        // Creates an Document instance and adds its children.
        public byte[] GenerateCompetencyLearningActivitiesReport(List<LearningGoals_Archive> learningGoals, List<LearningActivities_Archive> programLearningActivities, List<Competencies_LearningActivities_Archive> competencyLearningActivities, List<Competencies_Archive> compsArchives, List<Descriptors_Archive> descriptorsArchives)
        {
            byte[] reportBytes = null;
            using (MemoryStream stream = new MemoryStream())
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    CreateCompetencyLearningActivityReportParts(package, learningGoals, programLearningActivities, competencyLearningActivities, compsArchives , descriptorsArchives);
                }

                reportBytes = stream.ToArray();
            }

            return reportBytes;
        }

        #region LearningActivity Report

        // Adds child parts and generates content of the specified part.
        private void CreateLearningActivityReportParts(SpreadsheetDocument document, List<LearningActivities_Archive> programLearningActivities)
        {
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateLearningActivityReportWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateLearningActivityReportWorkbookStylesPart1Content(workbookStylesPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            GenerateLearningActivityReportSharedStringTablePart1Content(sharedStringTablePart1, programLearningActivities);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateLearningActivityReportWorksheetPart1Content(worksheetPart1, programLearningActivities);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            GenerateLearningActivityReportThemePart1Content(themePart1);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateLearningActivityReportExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetLearningActivityReportPackageProperties(document);
        }

        // Generates content of workbookPart1.
        private void GenerateLearningActivityReportWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "5", BuildVersion = "24816" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { AutoCompressPictures = false };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)28800U, WindowHeight = (UInt32Value)16300U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Learning Activities", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)140001U, ConcurrentCalculation = false };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{7523E5D3-25F3-A5E0-1632-64F254C22452}" };
            workbookExtension1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:ArchID Flags=\"2\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            workbookExtension1.Append(openXmlUnknownElement1);

            workbookExtensionList1.Append(workbookExtension1);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateLearningActivityReportWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)10U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 14D };
            Color color3 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontScheme3);

            Font font4 = new Font();
            FontSize fontSize4 = new FontSize() { Val = 11D };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(fontSize4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontScheme4);

            Font font5 = new Font();
            Underline underline1 = new Underline();
            FontSize fontSize5 = new FontSize() { Val = 11D };
            Color color4 = new Color() { Theme = (UInt32Value)10U };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font5.Append(underline1);
            font5.Append(fontSize5);
            font5.Append(color4);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontScheme5);

            Font font6 = new Font();
            Underline underline2 = new Underline();
            FontSize fontSize6 = new FontSize() { Val = 11D };
            Color color5 = new Color() { Theme = (UInt32Value)11U };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font6.Append(underline2);
            font6.Append(fontSize6);
            font6.Append(color5);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontScheme6);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 8D };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font7.Append(fontSize7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontScheme7);

            Font font8 = new Font();
            Underline underline3 = new Underline();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(underline3);
            font8.Append(fontSize8);
            font8.Append(fontName8);
            font8.Append(fontScheme8);

            Font font9 = new Font();
            Underline underline4 = new Underline();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color6 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName9 = new FontName() { Val = "Calibri" };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(underline4);
            font9.Append(fontSize9);
            font9.Append(color6);
            font9.Append(fontName9);
            font9.Append(fontScheme9);

            Font font10 = new Font();
            Italic italic1 = new Italic();
            FontSize fontSize10 = new FontSize() { Val = 11D };
            Color color7 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName10 = new FontName() { Val = "Calibri" };
            FontScheme fontScheme10 = new FontScheme() { Val = FontSchemeValues.Minor };

            font10.Append(italic1);
            font10.Append(fontSize10);
            font10.Append(color7);
            font10.Append(fontName10);
            font10.Append(fontScheme10);

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

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

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
            LeftBorder leftBorder2 = new LeftBorder();
            RightBorder rightBorder2 = new RightBorder();
            TopBorder topBorder2 = new TopBorder();

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Auto = true };

            bottomBorder2.Append(color8);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)25U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

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

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)9U };
            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat27.Append(alignment1);

            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { WrapText = true };

            cellFormat28.Append(alignment2);

            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat29.Append(alignment3);

            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat30.Append(alignment4);

            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat31.Append(alignment5);
            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat33.Append(alignment6);

            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat34.Append(alignment7);

            cellFormats1.Append(cellFormat26);
            cellFormats1.Append(cellFormat27);
            cellFormats1.Append(cellFormat28);
            cellFormats1.Append(cellFormat29);
            cellFormats1.Append(cellFormat30);
            cellFormats1.Append(cellFormat31);
            cellFormats1.Append(cellFormat32);
            cellFormats1.Append(cellFormat33);
            cellFormats1.Append(cellFormat34);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)25U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)2U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle2 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)4U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle3 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)6U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle4 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)8U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle5 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)10U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle6 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)12U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle7 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)14U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle8 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)16U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle9 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)18U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle10 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)20U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle11 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)22U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle12 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)24U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle13 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle14 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)3U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle15 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)5U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle16 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)7U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle17 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)9U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle18 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)11U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle19 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)13U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle20 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)15U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle21 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)17U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle22 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)19U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle23 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)21U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle24 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)23U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle25 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

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
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateLearningActivityReportSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1, List<LearningActivities_Archive> programLearningActivities)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)53U, UniqueCount = (UInt32Value)41U };

            AddLearningActivityReportHeaders(sharedStringTable1, string.Empty);

            AddLearningActivitiesUniqueStrings(sharedStringTable1, programLearningActivities);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of worksheetPart1.
        private void GenerateLearningActivityReportWorksheetPart1Content(WorksheetPart worksheetPart1, List<LearningActivities_Archive> programLearningActivities)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetProperties sheetProperties1 = new SheetProperties() { EnableFormatConditionsCalculation = false };
            TabColor tabColor1 = new TabColor() { Theme = (UInt32Value)9U, Tint = -0.249977111117893D };
            PageSetupProperties pageSetupProperties1 = new PageSetupProperties() { FitToPage = true };

            sheetProperties1.Append(tabColor1);
            sheetProperties1.Append(pageSetupProperties1);
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = string.Concat("A1:E", (programLearningActivities.Count + 5).ToString()) };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, ZoomScale = (UInt32Value)150U, ZoomScaleNormal = (UInt32Value)150U, ZoomScalePageLayoutView = (UInt32Value)150U, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { BaseColumnWidth = (UInt32Value)10U, DefaultColumnWidth = 8.83203125D, DefaultRowHeight = 14D, DyDescent = 0D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 8D, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 28.33203125D, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 63.1640625D, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 46D, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 27D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);

            Cell cell2 = new Cell() { CellReference = "D1", DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell2.Append(cellValue2);

            row1.Append(cell1);
            row1.Append(cell2);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell3 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "2";

            cell3.Append(cellValue3);

            row2.Append(cell3);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell4 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "3";

            cell4.Append(cellValue4);

            row3.Append(cell4);

            Row row4 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 18D };

            Cell cell5 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "4";

            cell5.Append(cellValue5);

            Cell cell6 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "5";

            cell6.Append(cellValue6);

            Cell cell7 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)5U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "6";

            cell7.Append(cellValue7);

            Cell cell8 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "7";

            cell8.Append(cellValue8);

            Cell cell9 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "8";

            cell9.Append(cellValue9);

            row4.Append(cell5);
            row4.Append(cell6);
            row4.Append(cell7);
            row4.Append(cell8);
            row4.Append(cell9);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);

            int currentRowIndex = 5;
            KeyValuePair<string, int> entry;
            foreach (LearningActivities_Archive learningActivity in programLearningActivities)
            {
                currentRowIndex++;

                Row row = new Row() { RowIndex = (UInt32Value)(uint)currentRowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 42D };

                Cell cell10 = new Cell() { CellReference = string.Concat("A", currentRowIndex), StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
                CellValue cellValue10 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_POSITION_FORMAT, learningActivity.LearningActivityID)));
                cellValue10.Text = entry.Value.ToString();

                cell10.Append(cellValue10);

                Cell cell11 = new Cell() { CellReference = string.Concat("B", currentRowIndex), StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
                CellValue cellValue11 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_TITLE_FORMAT, learningActivity.LearningActivityID)));
                cellValue11.Text = entry.Value.ToString();

                cell11.Append(cellValue11);

                Cell cell12 = new Cell() { CellReference = string.Concat("C", currentRowIndex), StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
                CellValue cellValue12 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_SCENARIO_FORMAT, learningActivity.LearningActivityID)));
                cellValue12.Text = entry.Value.ToString();

                cell12.Append(cellValue12);

                Cell cell13 = new Cell() { CellReference = string.Concat("D", currentRowIndex), StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
                CellValue cellValue13 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_TOPICSREQUIRED_FORMAT, learningActivity.LearningActivityID)));
                cellValue13.Text = entry.Value.ToString();

                cell13.Append(cellValue13);

                Cell cell14 = new Cell() { CellReference = string.Concat("E", currentRowIndex), StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
                CellValue cellValue14 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_WEEKS_FORMAT, learningActivity.LearningActivityID)));
                cellValue14.Text = entry.Value.ToString();

                cell14.Append(cellValue14);

                row.Append(cell10);
                row.Append(cell11);
                row.Append(cell12);
                row.Append(cell13);
                row.Append(cell14);

                sheetData1.Append(row);
            }
            
            PhoneticProperties phoneticProperties1 = new PhoneticProperties() { FontId = (UInt32Value)6U, Type = PhoneticValues.NoConversion };
            PageMargins pageMargins1 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Scale = (UInt32Value)70U, FitToHeight = (UInt32Value)0U, Orientation = OrientationValues.Landscape, VerticalDpi = (UInt32Value)2U };

            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&D";

            headerFooter1.Append(oddHeader1);

            WorksheetExtensionList worksheetExtensionList1 = new WorksheetExtensionList();

            WorksheetExtension worksheetExtension1 = new WorksheetExtension() { Uri = "{64002731-A6B0-56B0-2670-7721B7C09600}" };
            worksheetExtension1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:PLV Mode=\"0\" OnePage=\"0\" WScale=\"0\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            worksheetExtension1.Append(openXmlUnknownElement2);

            worksheetExtensionList1.Append(worksheetExtension1);

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(phoneticProperties1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(worksheetExtensionList1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of themePart1.
        private void GenerateLearningActivityReportThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme31 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme31.Append(majorFont1);
            fontScheme31.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme31);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateLearningActivityReportExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Macintosh Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

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
            vTLPSTR2.Text = "Learning Activities";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

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

        private void SetLearningActivityReportPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Windows User";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-08-05T17:16:07Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-09-05T14:05:34Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "MichaelS Brown";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2014-09-03T12:58:32Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #endregion

        #region CompetencyLearningActivities Report

        // Adds child parts and generates content of the specified part.
        private void CreateCompetencyLearningActivityReportParts(SpreadsheetDocument document, List<LearningGoals_Archive> learningGoals, List<LearningActivities_Archive> programLearningActivities, List<Competencies_LearningActivities_Archive> competencyLearningActivities, List<Competencies_Archive> compsArchives, List<Descriptors_Archive> descriptorsArchives)
        {
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateCompetencyLearningActivityWorkbookPart1Content(workbookPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId3");
            GenerateCompetencyLearningActivityThemePart1Content(themePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId4");
            GenerateCompetencyLearningActivityWorkbookStylesPart1Content(workbookStylesPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId5");
            GenerateCompetencyLearningActivitySharedStringTablePart1Content(sharedStringTablePart1, learningGoals, programLearningActivities, compsArchives, descriptorsArchives );

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateCompetencyLearningActivityWorksheetPart1Content(worksheetPart1, learningGoals, programLearningActivities, competencyLearningActivities, compsArchives, descriptorsArchives);

            WorksheetPart worksheetPart2 = workbookPart1.AddNewPart<WorksheetPart>("rId2");
            GenerateCompetencyLearningActivityWorksheetPart2Content(worksheetPart2, programLearningActivities);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateCompetencyLearningActivityExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetCompetencyLearningActivityPackageProperties(document);
        }

        // Generates content of workbookPart1.
        private void GenerateCompetencyLearningActivityWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "6", BuildVersion = "25007" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { ShowInkAnnotation = false, AutoCompressPictures = false };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = -28060, YWindow = -640, WindowWidth = (UInt32Value)28800U, WindowHeight = (UInt32Value)16240U, TabRatio = (UInt32Value)601U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Program Competency Template", SheetId = (UInt32Value)25U, Id = "rId1" };
            Sheet sheet2 = new Sheet() { Name = "Learning Activities", SheetId = (UInt32Value)26U, Id = "rId2" };

            sheets1.Append(sheet1);
            sheets1.Append(sheet2);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)140001U, ConcurrentCalculation = false };

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{7523E5D3-25F3-A5E0-1632-64F254C22452}" };
            workbookExtension1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:ArchID Flags=\"2\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            workbookExtension1.Append(openXmlUnknownElement1);

            WorkbookExtension workbookExtension2 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
            workbookExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.WorkbookProperties workbookProperties2 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

            workbookExtension2.Append(workbookProperties2);

            workbookExtensionList1.Append(workbookExtension1);
            workbookExtensionList1.Append(workbookExtension2);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of themePart1.
        private void GenerateCompetencyLearningActivityThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateCompetencyLearningActivityWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)9U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme2);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme3);

            Font font3 = new Font();
            Underline underline1 = new Underline();
            FontSize fontSize3 = new FontSize() { Val = 11D };
            Color color3 = new Color() { Theme = (UInt32Value)10U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(underline1);
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontScheme4);

            Font font4 = new Font();
            Underline underline2 = new Underline();
            FontSize fontSize4 = new FontSize() { Val = 11D };
            Color color4 = new Color() { Theme = (UInt32Value)11U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(underline2);
            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontScheme5);

            Font font5 = new Font();
            Italic italic1 = new Italic();
            FontSize fontSize5 = new FontSize() { Val = 11D };
            Color color5 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font5.Append(italic1);
            font5.Append(fontSize5);
            font5.Append(color5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontScheme6);

            Font font6 = new Font();
            FontSize fontSize6 = new FontSize() { Val = 14D };
            Color color6 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font6.Append(fontSize6);
            font6.Append(color6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontScheme7);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 11D };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

            font7.Append(fontSize7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontScheme8);

            Font font8 = new Font();
            Underline underline3 = new Underline();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(underline3);
            font8.Append(fontSize8);
            font8.Append(fontName8);
            font8.Append(fontScheme9);

            Font font9 = new Font();
            Underline underline4 = new Underline();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color7 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName9 = new FontName() { Val = "Calibri" };
            FontScheme fontScheme10 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(underline4);
            font9.Append(fontSize9);
            font9.Append(color7);
            font9.Append(fontName9);
            font9.Append(fontScheme10);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);

            Fills fills1 = new Fills() { Count = (UInt32Value)3U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

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
            LeftBorder leftBorder2 = new LeftBorder();
            RightBorder rightBorder2 = new RightBorder();
            TopBorder topBorder2 = new TopBorder();

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Auto = true };

            bottomBorder2.Append(color8);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)13U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

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

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)30U };
            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat16.Append(alignment1);
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat18.Append(alignment2);

            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { WrapText = true };

            cellFormat19.Append(alignment3);

            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat20.Append(alignment4);
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat22.Append(alignment5);

            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat23.Append(alignment6);

            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat24.Append(alignment7);

            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat25.Append(alignment8);

            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat26.Append(alignment9);
            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };

            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat28.Append(alignment10);

            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };

            cellFormat29.Append(alignment11);

            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { WrapText = true };

            cellFormat30.Append(alignment12);

            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { WrapText = true };

            cellFormat31.Append(alignment13);

            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };

            cellFormat32.Append(alignment14);

            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat33.Append(alignment15);
            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };

            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment16 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat35.Append(alignment16);

            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment17 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat36.Append(alignment17);

            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment18 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat37.Append(alignment18);
            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };

            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment19 = new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat42.Append(alignment19);

            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment20 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top };

            cellFormat43.Append(alignment20);

            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);
            cellFormats1.Append(cellFormat19);
            cellFormats1.Append(cellFormat20);
            cellFormats1.Append(cellFormat21);
            cellFormats1.Append(cellFormat22);
            cellFormats1.Append(cellFormat23);
            cellFormats1.Append(cellFormat24);
            cellFormats1.Append(cellFormat25);
            cellFormats1.Append(cellFormat26);
            cellFormats1.Append(cellFormat27);
            cellFormats1.Append(cellFormat28);
            cellFormats1.Append(cellFormat29);
            cellFormats1.Append(cellFormat30);
            cellFormats1.Append(cellFormat31);
            cellFormats1.Append(cellFormat32);
            cellFormats1.Append(cellFormat33);
            cellFormats1.Append(cellFormat34);
            cellFormats1.Append(cellFormat35);
            cellFormats1.Append(cellFormat36);
            cellFormats1.Append(cellFormat37);
            cellFormats1.Append(cellFormat38);
            cellFormats1.Append(cellFormat39);
            cellFormats1.Append(cellFormat40);
            cellFormats1.Append(cellFormat41);
            cellFormats1.Append(cellFormat42);
            cellFormats1.Append(cellFormat43);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)13U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)2U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle2 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)4U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle3 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)6U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle4 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)8U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle5 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)10U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle6 = new CellStyle() { Name = "Followed Hyperlink", FormatId = (UInt32Value)12U, BuiltinId = (UInt32Value)9U, Hidden = true };
            CellStyle cellStyle7 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle8 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)3U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle9 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)5U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle10 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)7U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle11 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)9U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle12 = new CellStyle() { Name = "Hyperlink", FormatId = (UInt32Value)11U, BuiltinId = (UInt32Value)8U, Hidden = true };
            CellStyle cellStyle13 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

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
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

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
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateCompetencyLearningActivitySharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1, List<LearningGoals_Archive> learningGoals, List<LearningActivities_Archive> programLearningActivities, List<Competencies_Archive> compsArchives, List<Descriptors_Archive> descriptorsArchives)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)404U, UniqueCount = (UInt32Value)391U };
            
            //Learning Activities Report Fixed Strings
            AddLearningActivityReportHeaders(sharedStringTable1, string.Empty );
            
            //CompetencyLearningActivities Report Fixed Strings
            AddCompetencyLearningActivitiesReportHeaders(sharedStringTable1);

            //Add all unique Goal, Competency, and Descriptor strings
            foreach (var learningGoal in learningGoals.OrderBy(j=> j.Position))
            {
                uint index = AddSharedString(sharedStringTable1, string.Format("{0}: {1}", learningGoal.Title, learningGoal.Description)).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_GOAL_DESC_FORMAT, learningGoal.LearningGoalID), (int)index);

                foreach (var competency in compsArchives.Where(c => c.LearningGoalID == learningGoal.LearningGoalID).OrderBy(j => j.Position))
                {
                    index = AddSharedString(sharedStringTable1, competency.Description).Value;
                    propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_COMPETENCY_DESC_FORMAT, competency.CompetencyID), (int)index);

                    foreach (var descriptor in descriptorsArchives.Where( d => d.CompetencyID == competency.CompetencyID).OrderBy(j => j.Position))
                    {
                        index = AddSharedString(sharedStringTable1, descriptor.Position.ToString()).Value;
                        propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_DESCRIPTOR_NUMBER_FORMAT, descriptor.DescriptorID), (int)index); 

                        index = AddSharedString(sharedStringTable1, descriptor.Description).Value;
                        propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_DESCRIPTOR_DESC_FORMAT, descriptor.DescriptorID), (int)index);
                    }
                }
            }

            //Add all unique Learning Activities strings
            AddLearningActivitiesUniqueStrings(sharedStringTable1, programLearningActivities);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of worksheetPart1.
        private void GenerateCompetencyLearningActivityWorksheetPart1Content(WorksheetPart worksheetPart1, List<LearningGoals_Archive> learningGoals, List<LearningActivities_Archive> programLearningActivities, List<Competencies_LearningActivities_Archive> competencyLearningActivities, List<Competencies_Archive> compsArchives, List<Descriptors_Archive> descriptorsArchives)
        {
            List<MergeCell> mergedCells = new List<MergeCell>();
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetProperties sheetProperties1 = new SheetProperties() { EnableFormatConditionsCalculation = false };
            TabColor tabColor1 = new TabColor() { Rgb = "FFFFFF00" };

            sheetProperties1.Append(tabColor1);
            int totalRowCount = (learningGoals.Count + 5);

          
            foreach (var learningGoalArc in learningGoals)
            {
                foreach (var compArh in compsArchives.Where(c => c.LearningGoalID == learningGoalArc.LearningGoalID))
                {
                    totalRowCount +=1 ;
                }
            }

            foreach (var learningGoalArc in learningGoals)
            {
                foreach (var compArh in compsArchives.Where(c => c.LearningGoalID == learningGoalArc.LearningGoalID))
                {
                    foreach (var descripArc in descriptorsArchives)
                    {
                        totalRowCount += 1;    
                    }
                    
                }
            }


           // learningGoals.ForEach(lg => lg.Competencies.ToList().ForEach(c => totalRowCount += c.Descriptors.Count));
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = string.Concat("A1:N", totalRowCount.ToString()) };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1:A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { BaseColumnWidth = (UInt32Value)10U, DefaultColumnWidth = 8.83203125D, DefaultRowHeight = 14D, OutlineLevelRow = 2, DyDescent = 0D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 7.1640625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 9.1640625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 7.83203125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 106.6640625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)16384U, Width = 8.83203125D, Style = (UInt32Value)1U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);

            SheetData sheetData1 = new SheetData();
            int columnCount = (4 + programLearningActivities.Count);

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) }, StyleIndex = (UInt32Value)3U, CustomFormat = true };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "9";

            cell1.Append(cellValue1);

            row1.Append(cell1);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) }, StyleIndex = (UInt32Value)3U, CustomFormat = true };

            Cell cell2 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "2";

            cell2.Append(cellValue2);
            Cell cell3 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)6U };
            Cell cell4 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)6U };
            mergedCells.Add(new MergeCell() { Reference = "A2:D2" });

            row2.Append(cell2);
            row2.Append(cell3);
            row2.Append(cell4);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) }, StyleIndex = (UInt32Value)3U, CustomFormat = true };

            Cell cell5 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "3";

            cell5.Append(cellValue3);
            Cell cell6 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)6U };
            Cell cell7 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)6U };
            mergedCells.Add(new MergeCell() { Reference = "A3:D3" });

            row3.Append(cell5);
            row3.Append(cell6);
            row3.Append(cell7);

            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) } };

            Cell cell8 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "1";

            cell8.Append(cellValue4);

            Cell cell9 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "11";

            cell9.Append(cellValue5);

            row4.Append(cell8);
            row4.Append(cell9);

            Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) } };

            Cell cell10 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "10";

            cell10.Append(cellValue6);
            Cell cell11 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)13U };
            Cell cell12 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)13U };
            Cell cell13 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)13U };

            row5.Append(cell10);
            row5.Append(cell11);
            row5.Append(cell12);
            row5.Append(cell13);

            //Add Learning Activity Column headers
            char column = 'D';
            foreach (var  learningActivity in programLearningActivities)
            {
                column++;
                Cell cell = new Cell() { CellReference = string.Format("{0}5", column), StyleIndex = (UInt32Value)14U };
                CellValue cellValue = new CellValue();
                cellValue.Text = learningActivity.Position.ToString();
                cell.Append(cellValue);

                //Add Cell to row
                row5.Append(cell);
            }

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);

            //Display complete list
            uint currentRowIndex = 5;
            KeyValuePair<string, int> entry;
            
            foreach (var learningGoal in learningGoals)
            {
                currentRowIndex++;
                Row learningGoalRow = new Row() { RowIndex = (UInt32Value)currentRowIndex, Spans = new ListValue<StringValue>() { InnerText = string.Concat("1:", columnCount) } };

                Cell learningGoalNumberCell = new Cell() { CellReference = string.Format("A{0}", currentRowIndex), StyleIndex = (UInt32Value)14U };
                CellValue learningGoalNumberCellValue = new CellValue();
                learningGoalNumberCellValue.Text = learningGoal.Position.ToString();
                learningGoalNumberCell.Append(learningGoalNumberCellValue);

                Cell learningGoalDescriptionCell = new Cell() { CellReference = string.Format("B{0}", currentRowIndex), StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
                CellValue learningGoalDescriptionCellValue = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_GOAL_DESC_FORMAT, learningGoal.LearningGoalID)));
                learningGoalDescriptionCellValue.Text = entry.Value.ToString();
                learningGoalDescriptionCell.Append(learningGoalDescriptionCellValue);

                Cell learningGoalEmptyOneCell = new Cell() { CellReference = string.Format("C{0}", currentRowIndex), StyleIndex = (UInt32Value)13U };
                Cell learningGoalEmptyTwoCell = new Cell() { CellReference = string.Format("D{0}", currentRowIndex), StyleIndex = (UInt32Value)25U };
                mergedCells.Add(new MergeCell() { Reference = string.Format("B{0}:D{1}", currentRowIndex, currentRowIndex) });

                //Append first cells to row
                learningGoalRow.Append(learningGoalNumberCell);
                learningGoalRow.Append(learningGoalDescriptionCell);
                learningGoalRow.Append(learningGoalEmptyOneCell);
                learningGoalRow.Append(learningGoalEmptyTwoCell);
                
                //Loop for learning activities count
                column = 'D';
                foreach (var learningActivity in programLearningActivities)
                {
                    column++;
                    Cell learningGoalAssociateCell = new Cell() { CellReference = string.Format("{0}{1}", column, currentRowIndex), StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };

                    //If a competencyLearningActivity exists for this Learning Goal and LearningActivity
                    if(competencyLearningActivities.FirstOrDefault(r => r.CompetencyItemID == learningGoal.LearningGoalID && 
                                                                        r.CompetencyType == 1 &&
                                                                        r.LearningActivityID == learningActivity.LearningActivityID) != null)                    
                    {
                        CellValue learningGoalAssociateCellValue = new CellValue();
                        learningGoalAssociateCellValue.Text = "12";
                        learningGoalAssociateCell.Append(learningGoalAssociateCellValue);
                    }
                    learningGoalRow.Append(learningGoalAssociateCell);
                }

                //Add to Sheet
                sheetData1.Append(learningGoalRow);
                


                    foreach (var competency in compsArchives.Where(c => c.LearningGoalID == learningGoal.LearningGoalID))
                    {
                        currentRowIndex++;
                        Row competencyRow = new Row()
                        {
                            RowIndex = (UInt32Value) currentRowIndex,
                            Spans = new ListValue<StringValue>() {InnerText = string.Concat("1:", columnCount)},
                            Height = 14D,
                            CustomHeight = true,
                            OutlineLevel = 1
                        };
                        Cell competencyCell = new Cell()
                        {
                            CellReference = string.Format("A{0}", currentRowIndex),
                            StyleIndex = (UInt32Value) 14U
                        };

                        Cell competencyNumberCell = new Cell()
                        {
                            CellReference = string.Format("B{0}", currentRowIndex),
                            StyleIndex = (UInt32Value) 14U
                        };
                        CellValue competencyNumberCellValue = new CellValue();
                        competencyNumberCellValue.Text = competency.Position.ToString();
                        competencyNumberCell.Append(competencyNumberCellValue);

                        Cell competencyDescCell = new Cell()
                        {
                            CellReference = string.Format("C{0}", currentRowIndex),
                            StyleIndex = (UInt32Value) 15U,
                            DataType = CellValues.SharedString
                        };
                        CellValue competencyDescCellValue = new CellValue();
                        entry =
                            propertyToSharedStrings.FirstOrDefault(
                                pss =>
                                    pss.Key.Equals(string.Format(PROPERTY_KEY_COMPETENCY_DESC_FORMAT,
                                        competency.CompetencyID)));
                        competencyDescCellValue.Text = entry.Value.ToString();
                        competencyDescCell.Append(competencyDescCellValue);

                        Cell competencyEmptyCell = new Cell()
                        {
                            CellReference = string.Format("D{0}", currentRowIndex),
                            StyleIndex = (UInt32Value) 15U
                        };
                        mergedCells.Add(new MergeCell()
                        {
                            Reference = string.Format("C{0}:D{1}", currentRowIndex, currentRowIndex)
                        });

                        competencyRow.Append(competencyCell);
                        competencyRow.Append(competencyNumberCell);
                        competencyRow.Append(competencyDescCell);
                        competencyRow.Append(competencyEmptyCell);

                        //Loop for learning activities count
                        column = 'D';
                        foreach (var learningActivity in programLearningActivities)
                        {
                            column++;
                            Cell competencyAssociateCell = new Cell()
                            {
                                CellReference = string.Format("{0}{1}", column, currentRowIndex),
                                StyleIndex = (UInt32Value) 2U,
                                DataType = CellValues.SharedString
                            };

                            //If a competencyLearningActivity exists for this Competency and LearningActivity
                            if (competencyLearningActivities.FirstOrDefault(
                                r => r.CompetencyItemID == competency.CompetencyID &&
                                     r.CompetencyType == 2 &&
                                     r.LearningActivityID == learningActivity.LearningActivityID) != null)
                            {
                                CellValue competencyAssociateCellValue = new CellValue();
                                competencyAssociateCellValue.Text = "12";
                                competencyAssociateCell.Append(competencyAssociateCellValue);
                            }
                            competencyRow.Append(competencyAssociateCell);
                        }

                        //Add to Sheet
                        sheetData1.Append(competencyRow);

                        foreach (var descriptor in descriptorsArchives.Where(c=> c.CompetencyID == competency.CompetencyID) )
                        {
                            currentRowIndex++;
                            Row descriptorRow = new Row()
                            {
                                RowIndex = (UInt32Value) currentRowIndex,
                                Spans = new ListValue<StringValue>() {InnerText = string.Concat("1:", columnCount)},
                                OutlineLevel = 2
                            };
                            Cell descriptorFirstCell = new Cell()
                            {
                                CellReference = string.Format("A{0}", currentRowIndex),
                                StyleIndex = (UInt32Value) 14U
                            };
                            Cell descriptorSecondCell = new Cell()
                            {
                                CellReference = string.Format("B{0}", currentRowIndex),
                                StyleIndex = (UInt32Value) 14U
                            };

                            Cell descriptorNumberCell = new Cell()
                            {
                                CellReference = string.Format("C{0}", currentRowIndex),
                                StyleIndex = (UInt32Value) 14U,
                                DataType = CellValues.SharedString
                            };
                            CellValue descriptorNumberCellValue = new CellValue();
                            entry =
                                propertyToSharedStrings.FirstOrDefault(
                                    pss =>
                                        pss.Key.Equals(string.Format(PROPERTY_KEY_DESCRIPTOR_NUMBER_FORMAT,
                                            descriptor.DescriptorID)));
                            descriptorNumberCellValue.Text = entry.Value.ToString();

                            descriptorNumberCell.Append(descriptorNumberCellValue);

                            Cell descriptorDescCell = new Cell()
                            {
                                CellReference = string.Format("D{0}", currentRowIndex),
                                StyleIndex = (UInt32Value) 13U,
                                DataType = CellValues.SharedString
                            };
                            CellValue descriptorDescCellValue = new CellValue();
                            entry =
                                propertyToSharedStrings.FirstOrDefault(
                                    pss =>
                                        pss.Key.Equals(string.Format(PROPERTY_KEY_DESCRIPTOR_DESC_FORMAT,
                                            descriptor.DescriptorID)));
                            descriptorDescCellValue.Text = entry.Value.ToString();

                            descriptorDescCell.Append(descriptorDescCellValue);

                            descriptorRow.Append(descriptorFirstCell);
                            descriptorRow.Append(descriptorSecondCell);
                            descriptorRow.Append(descriptorNumberCell);
                            descriptorRow.Append(descriptorDescCell);

                            //Loop for learning activities count
                            column = 'D';
                            foreach (var learningActivity in programLearningActivities)
                            {
                                column++;
                                Cell descriptorAssociateCell = new Cell()
                                {
                                    CellReference = string.Format("{0}{1}", column, currentRowIndex),
                                    StyleIndex = (UInt32Value) 2U,
                                    DataType = CellValues.SharedString
                                };

                                //If a competencyLearningActivity exists for this Descriptor and LearningActivity
                                if (competencyLearningActivities.FirstOrDefault(
                                    r => r.CompetencyItemID == descriptor.DescriptorID &&
                                         r.CompetencyType == 3 &&
                                         r.LearningActivityID == learningActivity.LearningActivityID) != null)
                                {
                                    CellValue descriptorAssociateCellValue = new CellValue();
                                    descriptorAssociateCellValue.Text = "12";
                                    descriptorAssociateCell.Append(descriptorAssociateCellValue);
                                }
                                descriptorRow.Append(descriptorAssociateCell);
                            }

                            //Add to Sheet
                            sheetData1.Append(descriptorRow);
                        }
                    }
                
            }            

            MergeCells mergeCells = new MergeCells() { Count = (UInt32Value)(uint)mergedCells.Count };
            mergedCells.ForEach(mc => mergeCells.Append(mc));
           
            PageMargins pageMargins1 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape };

            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&D";

            headerFooter1.Append(oddHeader1);

            WorksheetExtensionList worksheetExtensionList1 = new WorksheetExtensionList();

            WorksheetExtension worksheetExtension1 = new WorksheetExtension() { Uri = "{64002731-A6B0-56B0-2670-7721B7C09600}" };
            worksheetExtension1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:PLV Mode=\"0\" OnePage=\"0\" WScale=\"0\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            worksheetExtension1.Append(openXmlUnknownElement2);

            worksheetExtensionList1.Append(worksheetExtension1);

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(worksheetExtensionList1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of worksheetPart2.
        private void GenerateCompetencyLearningActivityWorksheetPart2Content(WorksheetPart worksheetPart2, List<LearningActivities_Archive> programLearningActivities)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = string.Concat("A1:E", (programLearningActivities.Count + 5).ToString()) };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { ZoomScale = (UInt32Value)150U, ZoomScaleNormal = (UInt32Value)150U, ZoomScalePageLayoutView = (UInt32Value)150U, WorkbookViewId = (UInt32Value)0U };
            Selection selection2 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { BaseColumnWidth = (UInt32Value)10U, DefaultColumnWidth = 8.83203125D, DefaultRowHeight = 14D, DyDescent = 0D };

            Columns columns2 = new Columns();
            Column column6 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 15.6640625D, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 36.5D, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 36D, CustomWidth = true };

            columns2.Append(column6);
            columns2.Append(column7);
            columns2.Append(column8);

            SheetData sheetData2 = new SheetData();

            Row row337 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell2397 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue407 = new CellValue();
            cellValue407.Text = "0";

            cell2397.Append(cellValue407);

            Cell cell2398 = new Cell() { CellReference = "D1", DataType = CellValues.SharedString };
            CellValue cellValue408 = new CellValue();
            cellValue408.Text = "1";

            cell2398.Append(cellValue408);

            row337.Append(cell2397);
            row337.Append(cell2398);

            Row row338 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell2399 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue409 = new CellValue();
            cellValue409.Text = "2";

            cell2399.Append(cellValue409);

            row338.Append(cell2399);

            Row row339 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:5" } };

            Cell cell2400 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)7U, DataType = CellValues.SharedString };
            CellValue cellValue410 = new CellValue();
            cellValue410.Text = "3";

            cell2400.Append(cellValue410);

            row339.Append(cell2400);

            Row row340 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 18D };

            Cell cell2401 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue411 = new CellValue();
            cellValue411.Text = "4";

            cell2401.Append(cellValue411);

            Cell cell2402 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue412 = new CellValue();
            cellValue412.Text = "5";

            cell2402.Append(cellValue412);

            Cell cell2403 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)8U, DataType = CellValues.SharedString };
            CellValue cellValue413 = new CellValue();
            cellValue413.Text = "6";

            cell2403.Append(cellValue413);

            Cell cell2404 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue414 = new CellValue();
            cellValue414.Text = "7";

            cell2404.Append(cellValue414);

            Cell cell2405 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)9U, DataType = CellValues.SharedString };
            CellValue cellValue415 = new CellValue();
            cellValue415.Text = "8";

            cell2405.Append(cellValue415);

            row340.Append(cell2401);
            row340.Append(cell2402);
            row340.Append(cell2403);
            row340.Append(cell2404);
            row340.Append(cell2405);

            sheetData2.Append(row337);
            sheetData2.Append(row338);
            sheetData2.Append(row339);
            sheetData2.Append(row340);

            int currentRowIndex = 5;
            KeyValuePair<string, int> entry;
            foreach (var learningActivity in programLearningActivities)
            {
                currentRowIndex++;
                Row row = new Row() { RowIndex = (UInt32Value)(uint)currentRowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 168D };

                Cell cell2406 = new Cell() { CellReference = string.Concat("A", currentRowIndex), StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
                CellValue cellValue416 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_POSITION_FORMAT, learningActivity.LearningActivityID)));
                cellValue416.Text = entry.Value.ToString();

                cell2406.Append(cellValue416);

                Cell cell2407 = new Cell() { CellReference = string.Concat("B", currentRowIndex), StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
                CellValue cellValue417 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_TITLE_FORMAT, learningActivity.LearningActivityID)));
                cellValue417.Text = entry.Value.ToString();

                cell2407.Append(cellValue417);

                Cell cell2408 = new Cell() { CellReference = string.Concat("C", currentRowIndex), StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
                CellValue cellValue418 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_SCENARIO_FORMAT, learningActivity.LearningActivityID)));
                cellValue418.Text = entry.Value.ToString();

                cell2408.Append(cellValue418);

                Cell cell2409 = new Cell() { CellReference = string.Concat("D", currentRowIndex), StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
                CellValue cellValue419 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_TOPICSREQUIRED_FORMAT, learningActivity.LearningActivityID)));
                cellValue419.Text = entry.Value.ToString();

                cell2409.Append(cellValue419);

                Cell cell2410 = new Cell() { CellReference = string.Concat("E", currentRowIndex), StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
                CellValue cellValue420 = new CellValue();
                entry = propertyToSharedStrings.FirstOrDefault(pss => pss.Key.Equals(string.Format(PROPERTY_KEY_ACTIVITY_WEEKS_FORMAT, learningActivity.LearningActivityID)));
                cellValue420.Text = entry.Value.ToString();

                cell2410.Append(cellValue420);

                row.Append(cell2406);
                row.Append(cell2407);
                row.Append(cell2408);
                row.Append(cell2409);
                row.Append(cell2410);

                sheetData2.Append(row);
            }

            PageMargins pageMargins2 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

            WorksheetExtensionList worksheetExtensionList2 = new WorksheetExtensionList();

            WorksheetExtension worksheetExtension2 = new WorksheetExtension() { Uri = "{64002731-A6B0-56B0-2670-7721B7C09600}" };
            worksheetExtension2.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<mx:PLV Mode=\"0\" OnePage=\"0\" WScale=\"0\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" />");

            worksheetExtension2.Append(openXmlUnknownElement3);

            worksheetExtensionList2.Append(worksheetExtension2);

            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns2);
            worksheet2.Append(sheetData2);
            worksheet2.Append(pageMargins2);
            worksheet2.Append(worksheetExtensionList2);

            worksheetPart2.Worksheet = worksheet2;
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateCompetencyLearningActivityExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Macintosh Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "2";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)2U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Program Competency Template";
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Sheet1";

            vTVector2.Append(vTLPSTR2);
            vTVector2.Append(vTLPSTR3);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Hewlett-Packard";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void SetCompetencyLearningActivityPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Summer Atkinson";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2013-10-28T13:01:18Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-09-12T18:54:14Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "MichaelS Brown";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2014-08-08T19:49:18Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #endregion
                
        private void AddLearningActivityReportHeaders(SharedStringTable sharedStringTable1, string versionName)
        {
            //Add initial items with index from 0 to 8
            AddSharedString(sharedStringTable1, "TGS Learning Activities Template ");

            AddSharedString(sharedStringTable1, string.Format("Version: {0}", versionName));

            AddSharedString(sharedStringTable1, string.Format("Program Name: {0}", this.ProgramName));

            AddSharedString(sharedStringTable1, string.Format("Submitter: {0}", this.GeneratorUsername));

            AddSharedString(sharedStringTable1, "LA #");

            AddSharedString(sharedStringTable1, "LA Title");

            AddSharedString(sharedStringTable1, "LA Scenario");

            AddSharedString(sharedStringTable1, "Learning Topics Required");

            AddSharedString(sharedStringTable1, "Weeks ");
        }

        private void AddCompetencyLearningActivitiesReportHeaders(SharedStringTable sharedStringTable1)
        {
            //Add initial CompetencyLearningActivitiesReport items with index from 9 to 12
            AddSharedString(sharedStringTable1, "TGS Competency Template");

            AddSharedString(sharedStringTable1, "Goal / Competency / Descriptor");

            AddSharedString(sharedStringTable1, "Learning Activities");

            AddSharedString(sharedStringTable1, "X");
        }

        private void AddLearningActivitiesUniqueStrings(SharedStringTable sharedStringTable1, List<LearningActivities_Archive> programLearningActivities)
        {
            foreach (LearningActivities_Archive learningActivity in programLearningActivities)
            {
                uint index = AddSharedString(sharedStringTable1, learningActivity.Position.ToString()).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_ACTIVITY_POSITION_FORMAT, learningActivity.LearningActivityID), (int)index);

                index = AddSharedString(sharedStringTable1, learningActivity.Title).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_ACTIVITY_TITLE_FORMAT, learningActivity.LearningActivityID), (int)index);

                index = AddSharedString(sharedStringTable1, learningActivity.Scenario, true).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_ACTIVITY_SCENARIO_FORMAT, learningActivity.LearningActivityID), (int)index);

                index = AddSharedString(sharedStringTable1, learningActivity.TopicsRequired, true).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_ACTIVITY_TOPICSREQUIRED_FORMAT, learningActivity.LearningActivityID), (int)index);

                index = AddSharedString(sharedStringTable1, learningActivity.Weeks.ToString()).Value;
                propertyToSharedStrings.Add(string.Format(PROPERTY_KEY_ACTIVITY_WEEKS_FORMAT, learningActivity.LearningActivityID), (int)index);
            }
        }

        private UInt32Value AddSharedString(SharedStringTable sharedStringTable, string sharedString, bool preserveSpaces = false)
        {
            SharedStringItem sharedStringItem = (SharedStringItem)sharedStringTable.FirstOrDefault(d => d.InnerText.Equals(sharedString));
            if (sharedStringItem == null)
            {
                sharedStringItem = new SharedStringItem();
                Text text = new Text();
                if (preserveSpaces)
                    text.Space = SpaceProcessingModeValues.Preserve;
                text.Text = sharedString;
                sharedStringItem.Append(text);
                sharedStringTable.Append(sharedStringItem);
                return UInt32Value.FromUInt32((uint)sharedStringTable.ToList().Count-1);
            }
            else
            {
                return UInt32Value.FromUInt32((uint)GetSharedStringIndex(sharedStringTable, sharedString));
            }
        }

        private int GetSharedStringIndex(SharedStringTable sharedStringTable, string sharedString)
        {
            int stringIndex = -1;

            foreach (SharedStringItem item in sharedStringTable)
            {
                stringIndex += 1;
                if (item.InnerText.Equals(sharedString))
                {
                    break;
                }
            }
            return stringIndex;
        }

        

    }
}