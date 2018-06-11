// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module implements a collection of tables on sheets, in a single collection.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.report;

import reportd.table;
import reportd.database;
import reportd.cellformat;

import sdlang;
import dxml.writer;

import std.algorithm;
import std.range;
import std.file : readText;
import std.format : format;
import std.stdio : writeln, readln;
import std.string;
import std.conv : to;
import std.typecons : Tuple, tuple;


struct Sheet {
    string name;
    Table[][] tables;

    void addTable (Table t){
        tables ~= [t];
        if (t.hasWriteBack) t.toGrid; // force calculation
    }

    void addTable (Table t, int row){
        assert (row <= (tables.length+1));
        assert (row > 0);
        tables[row-1] ~= t;
        if (t.hasWriteBack) t.toGrid; // force calculation
    }

    ~this() {
        // Table has a pointer to database
        foreach (tableRow; tables) {
            foreach (table; tableRow) {
                //table.db=null;
            }
        }
    }
}

//alias CellMap = HashMap!(Point,GridCell);
alias CellMap = GridCell[Point];
class SheetCells {
    string name;
    CellMap cells;

    this (string n){
        name = n;
        //cells = CellMap(8); // TODO fix this
    }

    void clear(){
        import core.memory;
        GC.free(&cells);
        GC.collect;
    }
}

class Report {
    import d2sqlite3 : Database;

    string[] conditions;
    Sheet[] sheets;
    //Database db;

    immutable CellFormat[string] formats;


    SheetCells[] allCells;

    this (string reportSDL, string csvPath, string fieldsPath){
        this (reportSDL, createDBFromCSV (csvPath,fieldsPath));
    }

    this (string reportSDL, Database db){
        //db = d;


        auto root = parseSource(reportSDL.readText);

        // load sheets
        if (!root.maybe.tags["conditions"].empty) {
            foreach (tag; root.tags["conditions"][0].maybe.tags[""]) conditions ~= format("(%s)",tag.values[0].get!string);
        }

        // load formatting
        {
            CellFormat[string] _formats;
            auto formatRoot = parseSource(root.tags["format"][0].values[0].get!string.readText);
            foreach (elementTag; formatRoot.tags["element"]){
                CellFormat cf;
                auto elementName = elementTag.values[0].get!string;

                // font
                foreach (fontTag; elementTag.maybe.tags["font"]){
                    if (!fontTag.maybe.tags["name"].empty) cf.fontName = fontTag.tags["name"][0].values[0].get!string;
                    if (!fontTag.maybe.tags["size"].empty) cf.fontSize = fontTag.tags["size"][0].values[0].get!int;
                    if (!fontTag.maybe.tags["colour"].empty) cf.fontColour = fontTag.tags["colour"][0].values[0].get!string;
                    if (!fontTag.maybe.tags["weight"].empty) cf.fontWeight = cast(FontWeight)fontTag.tags["weight"][0].values[0].get!string;
                }

                // alignment
                foreach (alignTag; elementTag.maybe.tags["alignment"]){
                    if (!alignTag.maybe.tags["horizontal"].empty) cf.horizontal = cast(HAlignment)alignTag.tags["horizontal"][0].values[0].get!string;
                    if (!alignTag.maybe.tags["vertical"].empty) cf.vertical = cast(VAlignment)alignTag.tags["vertical"][0].values[0].get!string;
                }

                // misc
                if (!elementTag.maybe.tags["bg-colour"].empty) cf.bgColour = elementTag.tags["bg-colour"][0].values[0].get!string;
                if (!elementTag.maybe.tags["wrap"].empty) cf.wrap = elementTag.tags["wrap"][0].values[0].get!bool;

                _formats[elementName] = cf;
            }
            import std.exception : assumeUnique;
            formats = _formats.assumeUnique;
        }

        // load tables
        // --------------------------------------------------------------
        {
            foreach (sheetTag; root.tags["sheet"]){
                auto sheetName = sheetTag.values[0].get!string;
                Sheet sheet;
                sheet.name = sheetName;
                foreach (tableTag; sheetTag.tags["table"]){
                    writeln ("Loading ",tableTag.values[0].get!string);
                    //auto tableToAdd = new Table (tableTag.values[0].get!string,db,conditions,formats);
                    //tableToAdd.hasWriteBack
                    sheet.addTable (new Table (tableTag.values[0].get!string,db,conditions,formats));
                }
                sheets ~= sheet;
            }
        }
    }

    void toGrid (string sheet){
        writeln ("processing ",sheet);
        //GridCell[Point] sheetCells;
        allCells ~= new SheetCells(sheet);
        auto sheetCells = &(allCells[$-1].cells);


        auto currentRow=2;
        //foreach (i, tableRow; sheets[sheet].tables){ // go down through the rows
        foreach (i, tableRow; sheets.filter!(a => a.name == sheet).front.tables){ // go down through the rows
            auto currentCol=2;
            auto maxRow=0;
            foreach (j, table; tableRow){            // go across each row
                auto tableCells = table.toGrid;

                auto lastRow = tableCells.byKey.map!(a => a.row).maxElement;
                auto lastCol = tableCells.byKey.map!(a => a.col).maxElement;
                maxRow = max (maxRow,lastRow+currentRow);

                // remap

                foreach (cell; tableCells.byValue){
                    auto newPoint = cell.point;
                    newPoint.row += currentRow;
                    newPoint.col += currentCol;
                    cell.point = newPoint;
                    (*sheetCells)[newPoint] = cell;
                }
                currentCol += lastCol+2;
            }
            //currentRow += maxRow+2;
            currentRow = maxRow+2;
        }




    }

    void toGrid(){


        //sheets.map!(a => Tuple!(string,"name",GridCell[Point],"cells")(a.name,toGrid(a.name)));
        sheets.each!(a => toGrid(a.name));
        //return Tuple!(string,"name",CellMap*,"cells")(sheets[0].name, (toGrid(sheets[0].name)));
    }

    void clear(){
        foreach (sheet; allCells){
            sheet.clear;
        }
        allCells.length=0;
    }


    void printODS (string fileOutput=""){
            import std.stdio;
            import reportd.cellcalc;

            writeln ("Exporting to ODS");

            // generate grid cells

            if (allCells.length == 0) toGrid;

            writeln ("Writing document...");
            //xmlDocPtr doc;
            //auto doc = xmlWriter(appender!string);

            //-----------------------
            // setup styles

            // get unique number styles
            // get unique cell styles (formats)
            NumberFormat[] numberStyles;
            CellFormat[] formatStyles;


            foreach (sheet; allCells){
                foreach (sheetCell; sheet.cells.values){
                    if (!numberStyles.canFind (sheetCell.calculation.format)) numberStyles ~= sheetCell.calculation.format;
                    if (!formatStyles.canFind (sheetCell.format)) formatStyles ~= sheetCell.format;
                }
            }


            import std.typecons : Tuple;
            alias FullStyleIndex = Tuple!(int,"number",int,"format");
            auto getStyleIndex (GridCell rc){
                FullStyleIndex rvalue;
                //rvalue.number = (rc.calculation is null) ? -1 : numberStyles.countUntil(rc.calculation.format);
                rvalue.number = (rc.calculation.empty) ? -1 : numberStyles.countUntil(rc.calculation.format).to!int;
                rvalue.format = formatStyles.countUntil(rc.format).to!int;
                return rvalue;
            }

            // get combinations used

            FullStyleIndex[] fullStyleIndexes;
            foreach (sheet; allCells){
                foreach (sheetCell; sheet.cells.values){
                    fullStyleIndexes ~= getStyleIndex(sheetCell);
                }
            }
            fullStyleIndexes.sort();


            // write styles.xml in memory
            //auto writer = xmlNewTextWriterDoc (&doc,0);
            import std.array : appender;
            auto styleWriter = xmlWriter(appender!string);
            styleWriter.output.writeXMLDecl!string;
            styleWriter.openStartTag ("office:document-styles");
            styleWriter.writeAttr ("xmlns:table","urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            styleWriter.writeAttr ("xmlns:office" , "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            styleWriter.writeAttr ("xmlns:text" , "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            styleWriter.writeAttr ("xmlns:style","urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            styleWriter.writeAttr ("xmlns:draw","urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            styleWriter.writeAttr ("xmlns:fo","urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
            styleWriter.writeAttr ("xmlns:xlink","http://www.w3.org/1999/xlink");
            styleWriter.writeAttr ("xmlns:dc","http://purl.org/dc/elements/1.1/");
            styleWriter.writeAttr ("xmlns:number","urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
            styleWriter.writeAttr ("xmlns:svg","urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
            styleWriter.writeAttr ("xmlns:msoxl","http://schemas.microsoft.com/office/excel/formula");
            styleWriter.closeStartTag;
            styleWriter.writeStartTag ("office:font-face-decls");
            styleWriter.openStartTag ("style:font-face");
            styleWriter.writeAttr ("style:name","Calibri");
            styleWriter.writeAttr ("svg:font-family","Calibri");
            styleWriter.closeStartTag;
            styleWriter.writeEndTag;
            styleWriter.writeEndTag;

            styleWriter.writeStartTag ("office:styles");

            foreach (i, style; numberStyles){
                string numberStyleType;
                bool createStyleCopy=false;

                switch (style.style){
                    default: numberStyleType = "number"; break;
                    case NumberStyle.number: numberStyleType = "number-style"; break;
                    case NumberStyle.percent: numberStyleType = "percentage-style"; createStyleCopy=true; break;
                }

                void writeStyle (string styleName) {
                    styleWriter.openStartTag ("number:" ~ numberStyleType);
                    styleWriter.writeAttr ("style:name",styleName);
                    styleWriter.closeStartTag;
                    styleWriter.openStartTag ("number:number");
                    if (style.withComma) styleWriter.writeAttr("number:grouping","true");
                    styleWriter.writeAttr ("number:decimal-places",format("%s",style.decimalPlaces));
                    if (style.scale != NumberScale.none) styleWriter.writeAttr ("number:display-factor",format("%s",cast(string)style.scale));
                    styleWriter.closeStartTag;
                    styleWriter.writeEndTag;
                    if (createStyleCopy) {
                        styleWriter.openStartTag ("style:map");
                        styleWriter.writeAttr ("style:condition","value()>=0");
                        styleWriter.writeAttr ("style:apply-style-name",format ("numberstyle%sCopy",i));
                        styleWriter.closeStartTag;
                        styleWriter.writeEndTag;
                    }
                    if (style.style == NumberStyle.percent) styleWriter.writeTaggedText ("number:text","%");
                    styleWriter.writeEndTag;
                }

                if (createStyleCopy) writeStyle (format("numberStyle%sCopy",i));
                writeStyle (format ("numberstyle%s",i));

                //styleWriter.writeEndTag;
            }

            foreach (i, style; formatStyles){
                styleWriter.openStartTag ("style:style");
                styleWriter.writeAttr ("style:name",format("parentStyle%s",i));
                styleWriter.writeAttr ("style:family","table-cell");
                styleWriter.closeStartTag;
                styleWriter.openStartTag ("style:table-cell-properties");
                styleWriter.writeAttr ("fo:border","thin solid #808080");
                styleWriter.writeAttr ("style:vertical-align",(style.vertical.isNull) ? "automatic" : style.vertical);
                if (!style.bgColour.isNull) styleWriter.writeAttr ("fo:background-color",style.bgColour.to!string);
                if (!style.wrap.isNull) styleWriter.writeAttr ("fo:wrap-option","wrap");
                styleWriter.closeStartTag;
                styleWriter.writeEndTag;

                styleWriter.openStartTag ("style:text-properties");
                if (!style.fontName.isNull) styleWriter.writeAttr ("style:font-name",style.fontName.to!string);
                if (!style.fontSize.isNull) styleWriter.writeAttr ("fo:font-size",format("%spt",style.fontSize));
                if (!style.fontColour.isNull) styleWriter.writeAttr ("fo:color",style.fontColour.to!string);
                if (!style.fontWeight.isNull) styleWriter.writeAttr ("fo:font-weight",style.fontWeight.to!string);
                styleWriter.closeStartTag;
                styleWriter.writeEndTag;
                if (!style.horizontal.isNull){
                    styleWriter.openStartTag ("style:paragraph-properties");
                    if (style.horizontal==HAlignment.centre) {
                        styleWriter.writeAttr ("fo:text-align","center");
                    } else {
                        styleWriter.writeAttr ("fo:text-align",style.horizontal.to!string);
                    }
                    styleWriter.closeStartTag;
                    styleWriter.writeEndTag;
                }
                styleWriter.writeEndTag;
            }

            styleWriter.writeEndTag;
            styleWriter.writeEndTag;


            // write content.xml to memory
            auto contentWriter = xmlWriter(appender!string);
            contentWriter.output.writeXMLDecl!string;

            contentWriter.openStartTag ("office:document-content");
            contentWriter.writeAttr ("xmlns:table","urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            contentWriter.writeAttr ("xmlns:office" , "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            contentWriter.writeAttr ("xmlns:text" , "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            contentWriter.writeAttr ("xmlns:style","urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            contentWriter.writeAttr ("xmlns:draw","urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            contentWriter.writeAttr ("xmlns:fo","urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
            contentWriter.writeAttr ("xmlns:xlink","http://www.w3.org/1999/xlink");
            contentWriter.writeAttr ("xmlns:dc","http://purl.org/dc/elements/1.1/");
            contentWriter.writeAttr ("xmlns:number","urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
            contentWriter.writeAttr ("xmlns:svg","urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
            contentWriter.writeAttr ("xmlns:msoxl","http://schemas.microsoft.com/office/excel/formula");
            contentWriter.closeStartTag;

            contentWriter.writeStartTag ("office:automatic-styles");

            // full styles generated from all report cells
            string[FullStyleIndex] styleNames;
            foreach (index; fullStyleIndexes.uniq){
                contentWriter.openStartTag ("style:style");
                auto parentName = format("parentStyle%s",index.format);
                auto numberName = (index.number == -1) ? "na" : format("numberStyle%s",index.number);
                auto styleName = joiner (["style",parentName,numberName],"-");
                styleNames[index] = styleName.to!string; // store stylename for easy reference when printing out cells
                contentWriter.writeAttr ("style:name",styleName);
                contentWriter.writeAttr ("style:family","table-cell");
                contentWriter.writeAttr ("style:parent-style-name",parentName);
                if (numberName != "na") contentWriter.writeAttr ("style:data-style-name",numberName);
                contentWriter.closeStartTag;
                contentWriter.writeEndTag;
            }

            // column width styles
            void writeColumnWithStyle (string name, string width) {
                contentWriter.openStartTag ("style:style");
                contentWriter.writeAttr ("style:name",name);
                contentWriter.writeAttr ("style:family","table-column");
                contentWriter.closeStartTag;
                contentWriter.openStartTag ("style:table-column-properties");
                contentWriter.writeAttr ("style:column-width",width);
                contentWriter.writeAttr ("style:use-optimal-column-width","true");
                contentWriter.closeStartTag;
                contentWriter.writeEndTag;
                contentWriter.writeEndTag;
            }

            auto columnWidthStyles = [ "smallWidth" : "0.6cm",
                                       "mediumWidth" : "1.1cm",
                                       "largeWidth" : "2.1cm",
                                       "extraWidth" : "4.1cm",
                                       "XXLWidth" : "6.1cm"];

            columnWidthStyles.byPair
                             .each!(a => writeColumnWithStyle(a[0],a[1]));

            contentWriter.writeEndTag;

            contentWriter.writeStartTag ("office:body");
            contentWriter.writeStartTag ("office:spreadsheet");


            foreach (sheet; allCells){
                auto cells = &sheet.cells;
                contentWriter.openStartTag ("table:table");
                contentWriter.writeAttr ("table:name",sheet.name);
                contentWriter.closeStartTag;

                auto height = cells.byKey.map!(a => a.row).maxElement;
                auto width = cells.byKey.map!(a => a.col).maxElement;

                // write out column setup

                void writeTableColumnStyle (string styleName) {
                    assert (styleName in columnWidthStyles);
                    contentWriter.openStartTag ("table:table-column");
                    contentWriter.writeAttr ("table:style-name",styleName);
                    contentWriter.closeStartTag;
                    contentWriter.writeEndTag;
                }

                foreach (j; 0..width+1){
                    auto columnCells = cells.byKey.filter!(a => a.col == j);
                    if (columnCells.count == 0){
                        writeTableColumnStyle ("smallWidth");
                    } else {
                        auto m = columnCells.map!(a => (*cells)[a].text.count)
                                        .array
                                        .maxCount[0];

                        if (m > 25) {
                            writeTableColumnStyle ("XXLWidth");
                        } else if (m > 20) {
                            writeTableColumnStyle ("extraWidth");
                        } else {
                            writeTableColumnStyle ("largeWidth");
                        }
                    }
                }


                // write out rows
                foreach (i; 0..height+1){
                    contentWriter.openStartTag ("table:table-row");
                    contentWriter.closeStartTag;
                    foreach (j; 0..width+1){
                        auto point = Point(i,j);
                        if (point in (*cells)){
                            auto reportCell = (*cells)[point];
                            contentWriter.openStartTag ("table:table-cell");

                            // write type
                            auto valueType = (reportCell.isStringType) ? "string" :
                                                    reportCell.isPercentType ? "percentage" : "float";

                            contentWriter.writeAttr ("office:value-type",valueType);

                            // write span
                            // for ODS, both row and col span needs to be written, or none
                            if (reportCell.rowSpan>0 || reportCell.colSpan>0){
                                contentWriter.writeAttr ("table:number-rows-spanned",format("%s",reportCell.rowSpan+1));
                                contentWriter.writeAttr ("table:number-columns-spanned",format("%s",reportCell.colSpan+1));
                            }


                            // write style
                            contentWriter.writeAttr ("table:style-name",styleNames[getStyleIndex(reportCell)]);

                            if (valueType == "percentage" || valueType == "float") {
                                contentWriter.writeAttr ("office:value",format("%.8f",reportCell.result));
                            }
                            contentWriter.closeStartTag;

                            // write value
                            switch (valueType) {
                                default: break;
                                case "string":
                                    contentWriter.writeTaggedText ("text:p", reportCell.isNAType ? "-" : reportCell.text);
                                    break;
                                case "percentage":
                                    // TODO make this all better

                                    auto formatString = format ("%%0.%df%%%%",reportCell.calculation.format.decimalPlaces);
                                    auto numericValue = reportCell.result;
                                    import std.math : abs;
                                    contentWriter.writeTaggedText ("text:p",format(formatString,abs(numericValue*100)));
                                    break;
                            }
                            contentWriter.writeEndTag;
                        } else {
                            contentWriter.writeTaggedText ("table:covered-table-cell","");
                        }
                    }
                    contentWriter.writeEndTag;
                }
                contentWriter.writeEndTag;
            }

            contentWriter.writeEndTag;
            contentWriter.writeEndTag;
            contentWriter.writeEndTag;

            // write to ODS
            import std.zip;

            auto odsPackage = new ZipArchive();
            auto archive = new ArchiveMember();

            // styles.xml
            archive.name = "styles.xml";
            archive.compressionMethod = CompressionMethod.deflate;
            archive.expandedData = cast(ubyte[])styleWriter.output.data;
            odsPackage.addMember (archive);

            archive = new ArchiveMember();
            archive.name = "content.xml";
            archive.compressionMethod = CompressionMethod.deflate;
            archive.expandedData = cast(ubyte[])contentWriter.output.data;
            odsPackage.addMember (archive);


            if (fileOutput.empty){
                string fileName;
                int index=0;
                auto done = false;
                import std.file : exists, remove;
                import std.process : executeShell;
                while (!done){
                    fileName = format ("output%s.ods",index);
                    if (!exists(fileName)){
                        done=true;
                    } else {
                        try {
                            remove (fileName);
                            done=true;
                        } catch (Exception ex){
                            // file open in another program
                            // try to save as next filename
                        }
                    }
                    index++;
                }

                auto file = File(fileName,"w");
                file.rawWrite (odsPackage.build());
                file.close;
                executeShell(fileName);
            } else {
                auto file = File(fileOutput,"w");
                file.rawWrite (odsPackage.build());
                file.close;
            }
        }
}

unittest {





}
