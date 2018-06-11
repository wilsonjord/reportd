// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
Implements "pivot table"-like functionality via the Table object.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)
*/

module table;

import std.algorithm;
import std.file : readText;
import std.stdio : readln, writeln, writefln;
import std.format : format;
import std.range;
import std.typecons : Tuple, Flag, Yes, No;
import std.conv : to;
import std.functional : memoize;
import std.exception : enforce;

import d2sqlite3;
import sdlang;
import luad.all;

import reportd.database;
import reportd.cell;
import reportd.cellcalc;
import reportd.cellformat;
import reportd.node;


enum Orientation {row,column}
alias TableVariableResult = Tuple!(CachedResults,"result",int[],"reference");

auto fitsReference (TableVariableResult dtv, int[] r) {
    if (r.equal(dtv.reference)){
        return true;
    }

    foreach (v; dtv.reference){
        bool found = false;
        foreach (t; r){
            if (t==v) found=true;
        }
        if (!found) return false;
    }
    return true;
}

auto fitsReference2(T) (TableVariableResult dtv, T r) {
    if (r.equal(dtv.reference)){
        return true;
    }

    if (dtv.reference.count!(a => !r.canFind(a)) == 0) return true;
    return false;
}

unittest {
    TableVariableResult a;
    a.reference = [1,2,3];

    assert (a.fitsReference ([1,2,3]));
    assert (a.fitsReference ([1,2,3,4]));
    assert (a.fitsReference ([4,2,1,3]));
    assert (!a.fitsReference ([0,5,2,6]));

}

struct TableVariable {
    string name;
    string select;
    string[] by;
    string where;
    int sqlCount=0;

    @property @nogc {
        auto isConstant() { return by.empty; }
        auto count() { return sqlCount; }
    }

    TableVariableResult[] results;

    void addResult (CachedResults r, int[] refs) {
        refs = refs.sort().array;
        results ~= TableVariableResult(r,refs);
        sqlCount++;
    }
}

struct TableVariables {
    TableVariable[string] data;

    @property @nogc {
        auto count() { return data.byValue.map!(a => a.count).sum; }
    }

    void add (TableVariable dtv) { data[dtv.name] = dtv;  }

    auto getVariable (string n) { return data[n]; }

    auto getResult (string n, int[] reference) {
//        return (data[n].isConstant) ? data[n].results[0].result :
//                                      data[n].results.filter!(a => a.fitsReference(reference)).map!(a => a.result).front;

        return (data[n].isConstant) ? [data[n].results[0].result] :
                                      data[n].results.filter!(a => a.fitsReference(reference)).map!(a => a.result).array;

    }


    auto getResult2(T) (string n, T reference) {
        return (data[n].isConstant) ? data[n].results[0].result :
                                      data[n].results.filter!(a => a.fitsReference2(reference)).map!(a => a.result).front;
    }

    auto names() { return data.keys; }

}

auto readCalc (Tag tag){
    CellCalc rvalue;
    rvalue.formula = tag.tags["formula"][0].values[0].get!string;
    if (!tag.maybe.tags["style"].empty) rvalue.format.style = cast (NumberStyle) tag.tags["style"][0].values[0].get!string;
    if (!tag.maybe.tags["scale"].empty) {
        switch (tag.tags["scale"][0].values[0].get!string) {
            default: throw new Exception ("Error: invalid scale value");
            case "none": rvalue.format.scale = NumberScale.none; break;
            case "thousand": rvalue.format.scale = NumberScale.thousand; break;
            case "million": rvalue.format.scale = NumberScale.million; break;
        }
    } else {
        rvalue.format.scale = NumberScale.none;
    }

    if (!tag.maybe.tags["decimal-place"].empty) {
        rvalue.format.decimalPlaces = tag.tags["decimal-place"][0].values[0].get!int;
    } else {
        rvalue.format.decimalPlaces = 0;
    }
    if (!tag.maybe.tags["with-comma"].empty) rvalue.format.withComma = tag.tags["with-comma"][0].values[0].get!bool;

    return rvalue;
}


alias Point = Tuple!(int,"row",int,"col");
//struct Point {
//    int row;
//    int col;
//
//    this (int r, int c){
//        row=r;
//        col=c;
//    }
//}

alias Result = Tuple!(double,"value",bool,"resultOK");
enum GridCellType {label,data,na,title,text}
alias GridCell = Tuple!(GridCellType,"type",Point,"point",string,"text",double,"result",CellFormat,"format",CellCalc,"calculation",int,"rowSpan",int,"colSpan");

auto isStringType (GridCell c){
    return (c.type==GridCellType.label || c.type==GridCellType.na || c.type==GridCellType.title || c.type==GridCellType.text);
}

auto isNumberType (GridCell c){
    return (c.type==GridCellType.data);
}

auto isPercentType (GridCell c){
    return (c.type==GridCellType.data && c.calculation.format.style==NumberStyle.percent);
}

auto isNAType (GridCell c){
    return (c.type==GridCellType.na);
}


auto gridLabel (Point p, string t, CellFormat f){
    GridCell cell;
    cell.type = GridCellType.label;
    cell.point = p;
    cell.text = t;
    cell.format = f;
    return cell;
}

auto gridTitle (Point p, string t, CellFormat f){
    GridCell cell;
    cell.type = GridCellType.title;
    cell.point = p;
    cell.text = t;
    cell.format = f;
    return cell;
}

auto gridData (Point p, double r, CellFormat f, CellCalc c){
    GridCell cell;
    cell.type = GridCellType.data;
    cell.point = p;
    cell.result = r;
    cell.format = f;
    cell.calculation = c;
    return cell;
}

auto gridText (Point p, string r, CellFormat f){
    GridCell cell;
    cell.type = GridCellType.text;
    cell.point = p;
    cell.format = f;
    cell.text = r;
    return cell;
}

auto gridNA (Point p, CellFormat f){
    GridCell cell;
    cell.point = p;
    cell.type = GridCellType.na;
    cell.format = f;
    return cell;
}

alias ConditionFormat = Tuple!(string,"condition",CellFormat,"format");

auto getGroupDepthLevel (LinkCell start, LinkCell cell){
    auto rvalue=0;
    auto childLevel=0;
    auto group = cell.groupName;
    auto currentMax = 0;
    void traverse (LinkCell c){
        if (cell == c) {
            rvalue = childLevel;
            return;
        }
        if (c.child) {
            if (sameGroup(group,c.child)) childLevel++;
            traverse (c.child);
            if (sameGroup(group,c.child)) childLevel--;
        }
        if (c.next) traverse (c.next);
    }
    traverse (start);
    return rvalue;
}
alias fastGetGroupDepthLevel = memoize!getGroupDepthLevel;

auto getGroupDepth (LinkCell start, string n){
    auto rvalue=0;
    auto childLevel=0;
    void traverse (LinkCell c){
        if (sameGroup (n,c)) rvalue = max(childLevel,rvalue);
        if (c.child) {
            if (sameGroup(n,c.child)) childLevel++;
            traverse (c.child);
            if (sameGroup(n,c.child)) childLevel--;
        }
        if (c.next) traverse (c.next);
    }
    traverse (start);
    return rvalue;
}

alias fastGetGroupDepth = memoize!getGroupDepth;

auto getDepth (LinkCell start, LinkCell cell){
    if (cell.child && sameGroup(cell,cell.child)) return 0;

    auto groupDepth = fastGetGroupDepth (start,cell.groupName);

    auto rvalue = 1;
    auto childLevel=0;
    void traverse (LinkCell c){
        if (cell == c) {
            rvalue = groupDepth - childLevel;
            return ;
        }

        if (c.child) {
            if (sameGroup(c.child,cell)) childLevel++;
            traverse (c.child);
            if (sameGroup(c.child,cell)) childLevel--;
        }

        if (c.next) traverse (c.next);
    }
    traverse (start);

    return rvalue;
}

alias fastGetDepth = memoize!getDepth;

auto getBreadth(bool delegate(LinkCell,LinkCell) sameOrientation, LinkCell cell){
    auto rvalue=0;
    void traverse (LinkCell c){
        // in case of row cells, count "terminal" cells as where it transitions to column cells
        // in case of column cells, count terminls as where child is null
        if (sameOrientation(cell,c) && (!c.child || (c.child && !sameOrientation(cell,c.child)))) rvalue++;
        if (c.child) traverse (c.child);
        if (c.next) traverse (c.next);
    }
    //if (cell.child && sameGroup(cell,cell.child)){
    if (cell.child && sameOrientation(cell,cell.child)){
        traverse(cell.child);
    } else {
        return 0;
    }
    return rvalue-1;
}

alias fastGetBreadth = memoize!getBreadth;

//------------------------------
// Functions for LUA
//------------------------------

auto round (double num, int place){
    import std.math : quantize, round;
    if (place==1) return num.quantize(0.1);
    if (place==2) return num.quantize(0.01);
    return round(num);
}

auto commatise(char[] txt) {
    auto step = 3;
    auto ins = ",";
    auto start = 0;
    import std.regex;
    if (start > txt.length || step > txt.length)
        return txt;

    // First number may begin with digit or decimal point. Exponents ignored.
    enum decFloField = ctRegex!("[0-9]*\\.[0-9]+|[0-9]+");

    auto matchDec = matchFirst(txt[start .. $], decFloField);
    if (!matchDec)
        return txt;

    // Within a decimal float field:
    // A decimal integer field to commatize is positive and not after a point.
    enum decIntField = ctRegex!("(?<=\\.)|[1-9][0-9]*");
    // A decimal fractional field is preceded by a point, and is only digits.
    enum decFracField = ctRegex!("(?<=\\.)[0-9]+");

    return txt[0 .. start] ~ matchDec.pre ~ matchDec.hit
        .replace!(m => m.hit.retro.chunks(step).join(ins).retro)(decIntField)
        .replace!(m => m.hit.chunks(step).join(ins))(decFracField)
        ~ matchDec.post;
}

auto contains (char[] a, char[] b){
    return a.canFind(b);
}

class Table {
    alias Span = Tuple!(int,"rowSpan",int,"colSpan");

    string[] conditions; // table level conditions
    Database db;

    LinkCell[string] rootCells;
    LinkCell[CellID] cellLookup;
    Span[CellID] spans;
    TableVariables variables;
    string[][Orientation] labelGroupOrientation;
    string title;
    CellFormat[string] formats;
    ConditionFormat[][CellID] conditionalFormats;

    string[] headers;

    LuaState lua;
    string[] luaFunctions;

    string[CellID] writeBack;

    double sqlsPerSecond=0;



    auto isOrientation (LinkCell c, Orientation ori){
        return labelGroupOrientation[ori].canFind!(a => a==c.groupName);
    }
    auto isRow (LinkCell c) { return isOrientation(c, Orientation.row); }
    auto isRow (GridCell c) { return isRow (c.point); }
    auto isRow (Point p) { return p.col<=rowLabelWidth && p.row >= columnLabelHeight; }

    auto isColumn (LinkCell c) { return isOrientation(c, Orientation.column); }
    auto isColumn (GridCell c) { return isColumn (c.point); }
    auto isColumn (Point p) { return p.row <=columnLabelHeight && p.col >= rowLabelWidth; }

    auto isHeader (GridCell c) { return isHeader (c.point); }
    auto isHeader (Point p) { return p.row<= columnLabelHeight && p.col <= columnLabelHeight; }
    //auto isRow (Point p) { return p.col<=rowLabelWidth; }

    auto sameOrientation (LinkCell c1, LinkCell c2) { return (isRow(c1) == isRow(c2)); }

    @property {
        auto rootCell() { return rootCells[labelGroupOrientation[Orientation.row][0]]; }
        //auto columnLabelHeight() { return labelGroupOrientation[Orientation.column].map!(a => getGroupDepth(rootCell,a)).sum; }
        auto columnLabelHeight() { return labelGroupOrientation[Orientation.column].map!(a => fastGetGroupDepth(rootCell,a)).sum; }

        //auto rowLabelWidth() { return labelGroupOrientation[Orientation.row].map!(a => getGroupDepth(rootCell,a)).sum; }
        auto rowLabelWidth() { return labelGroupOrientation[Orientation.row].map!(a => fastGetGroupDepth(rootCell,a)).sum; }

        auto hasWriteBack() { return writeBack.byKey.count > 0; }

        auto count() { return variables.count; }
    }

    this (string tableSDL, Database d, string[] reportConditions, immutable CellFormat[string] reportFormats){
        db = d;
        if (!reportConditions.empty) conditions ~= reportConditions;
        if (reportFormats !is null) {
            foreach (key; reportFormats.keys){
                formats[key]=reportFormats[key];
            }
        }
        this(tableSDL);
        //init(tableSDL);
    }

    this (string tableSDL){
        init(tableSDL);
    }

    this(){}

    auto getWhere (CellID[] branch) {
        return chain(branch.map!(a => cellLookup[a].where),[conditions.joiner(" AND ").array.to!(string)]) // apply where clause to main branch clause
                    .filter!(a => !a.empty) // filter out cases where "where" may be empty
                    .joiner (" AND ");

    }

    auto columnSpan (LinkCell cell){
        return (isRow(cell)) ? 1 : rootCell.fastGetDepth(cell);
    }

    auto rowSpan (LinkCell cell){
        return (isRow(cell)) ? fastGetBreadth(&sameOrientation,cell) : rootCell.fastGetDepth(cell);
    }

    void updateLua (ref int[] branchReference) {

        auto branchContains (string label){
            return branchReference.map!(a => cellLookup[a].text).canFind(label);
        }

        //lua = new LuaState;
        //lua.openLibs; // TODO - check if we need this
        lua["branchContains"] = &branchContains;


        foreach (f; luaFunctions) lua.doString (f); // TODO - don't need it to run every time


        foreach (var; variables.data.values){
            auto results = variables.getResult(var.name,branchReference);

            if (results.count == 0){
                // keep value as nil
            } else if (results.count == 1) {
                lua[var.name] = results[0][0][0].as!double;
            } else {
                lua[var.name] = results[0].map!(a => a[0].as!double).array;
            }
        }


    }

    void updateLua3 (ref int[] branchReference) {
        auto branchContains (string label){
            return branchReference.map!(a => cellLookup[a].text).canFind(label);
        }

        //lua = new LuaState;
        //lua.openLibs; // TODO - check if we need this
        lua["branchContains"] = &branchContains;

        foreach (f; luaFunctions) lua.doString (f); // TODO - don't need it to run every time

        foreach (var; variables.data.values){
            auto results = variables.getResult(var.name,branchReference);
            if (results.count == 0){
                // keep value as nil
            } else if (results.count == 1) {
                lua[var.name] = results[0][0][0].as!double;
            } else {
                lua[var.name] = results[0].map!(a => a[0].as!double).array;
            }
        }
    }

    void updateLua2 (int[] branchReference) {
        auto branchContains (string label){
            return branchReference.map!(a => cellLookup[a].text).canFind(label);
        }

        return;

        //lua = new LuaState;
        //lua.openLibs; // TODO - check if we need this

        lua["branchContains"] = &branchContains;

        foreach (f; luaFunctions) lua.doString (f); // TODO - don't need it to run every time

        foreach (var; variables.data.values){
            auto results = variables.getResult2(var.name,branchReference);
            if (results.count == 0){
                // keep value as nil
            } else if (results.count == 1) {
                lua[var.name] = results[0][0].as!double;
            } else {
                lua[var.name] = results.map!(a => a[0].as!double).array;
            }
        }
    }

    private void init (string tableSDL){
        lua = new LuaState;
        lua.openLibs;
        lua["round"] = &round;
        lua["commatise"] = &commatise;
        lua["contains"] = &contains;

        auto root = parseSource(tableSDL.readText);
        LinkCell[] noChildren;

        // load title
        if (!root.maybe.tags["title"].empty) {
            title = root.tags["title"][0].values[0].get!string;
        }


        // load table level conditions
        if (!root.maybe.tags["conditions"].empty) {
            foreach (tag; root.tags["conditions"][0].maybe.tags[""]) conditions ~= format("(%s)",tag.values[0].get!string);
        }

        // load label groups
        // --------------------------------------------------------------
        {
            foreach (labelGroupTag; root.tags["labels"]){
                CellCalc currentCalc;
                LinkCell currentRoot;
                LinkCell[][] allCells;
                string[] fieldNames;

                if (!labelGroupTag.maybe.tags["calculation"].empty) currentCalc = labelGroupTag.tags["calculation"][0].readCalc;

                auto labelGroupName = labelGroupTag.values[0].get!string;

                {
                    auto getLabelsFromFields (Tag[] fields, string condition){
                        if (fields.empty) return null;
                        //auto condTag = field.maybe.attributes["condition"];
                        //auto whereCondition = (condition.empty) ? "1=1" : condition;

                        //auto whereCondition = [fields[0].getAttribute!string("condition",""),condition].filter!(a => !a.empty);
                        auto whereCondition = [fields[0].getAttribute!string("condition",""),
                                               condition,
                                               conditions.joiner(" AND ").to!string].filter!(a => !a.empty);

                        string sql;
                        switch (whereCondition.count) {
                            default:
                                sql = format ("select distinct %s from data where %s;",fields[0].values[0],whereCondition.joiner(" AND "));
                                break;
                            case 0:
                                sql = format ("select distinct %s from data;",fields[0].values[0]);
                                break;
                            case 1:
                                sql = format ("select distinct %s from data where %s;",fields[0].values[0],whereCondition.front);
                                break;
                        }

                        auto cells = db.execute (sql).map!(a => a.peek!string(0))
                                                       .array
                                                       .sort()
                                                       .map!(a => createLinkCell (labelGroupName,a,format("%s='%s'",fields[0].values[0],a.replace("'","''")))).array;

                        if (cells.empty) return null;

                        // add Total cell
                        if (fields[0].getAttribute!bool("total",true)) cells ~= createLinkCell (labelGroupName,"Total","(" ~ cells.map!(a => a.where).joiner(" OR ").array.to!string ~ ")",Yes.isTotal);

                        // link cells
                        foreach (i; 0..cells.length-1){
                            cells[i].next = cells[i+1];
                        }

                        // set children
                        foreach (cell; cells){
                            cell.child = (whereCondition.empty) ? getLabelsFromFields (fields[1..$],cell.where) :
                                                                  getLabelsFromFields (fields[1..$],whereCondition.joiner(" AND ").to!string ~ " AND " ~ cell.where);
                        }
                        return cells[0];
                    }
                    // create indexes
                    fieldNames = labelGroupTag.maybe.tags["field"].map!(a => a.values[0].get!string).array;
                    if (!fieldNames.empty) db.createIndex (fieldNames);
                    currentRoot = getLabelsFromFields (labelGroupTag.maybe.tags["field"].array,"");
                }

                // translate any database fields into labels
//                {
//                    foreach (field; labelGroupTag.maybe.tags["field"]){
//                        fieldNames ~= field.values[0].to!string;
//                        auto condTag = field.maybe.attributes["condition"];
//                        auto generateTotal = field.getAttribute!bool("total",true);
//
//                        auto condition = chain([(condTag.empty ? null : condTag[0].value.get!string)],conditions).filter!(a => !a.empty).joiner(" AND ");
//
//                        void generateLabels (LinkCell parent=null){
//                            LinkCell[] _cells;
//                            auto finalCondition = [condition.to!string,parent.getWhere].filter!(a => !a.empty).joiner(" AND ");
//                            auto sql = (finalCondition.empty) ? format ("select distinct %s from data;",field.values[0]) :
//                                                                format ("select distinct %s from data where %s;",field.values[0],finalCondition);
//
//                            try {
//                                sql.writeln;
//                                auto results = db.execute(sql);
//                                foreach (row; results){
//                                    auto toAdd = (row.columnType(0)==SqliteType.TEXT) ? createLinkCell (labelGroupName,row.peek!string(0),format ("%s='%s'",field.values[0],row.peek!string(0).replace("'","''"))) :
//                                                                                        createLinkCell (labelGroupName,row.peek!string(0),format ("%s=%s",field.values[0],row.peek!double(0)));
//                                    if (!currentCalc.empty) toAdd.calculation = currentCalc;
//                                    _cells ~= toAdd;
//                                }
//                                _cells = _cells.sort!((a,b) => a.text < b.text).array;
//                                if (!_cells.empty){
//                                    if (generateTotal) _cells ~= createLinkCell (labelGroupName,"Total","(" ~ _cells.map!(a => a.where).joiner(" OR ").array.to!string ~ ")",Yes.isTotal);
//                                    for (int i=0; i<_cells.length-1; i++){ // ignore last cell
//                                        _cells[i].next = _cells[i+1];
//                                    }
//                                    allCells ~= _cells;
//                                    if (parent !is null) {
//                                        parent.child = _cells[0];
//                                    }
//                                }
//                            } catch (Exception ex){
//                                sql.writeln;
//                                ex.msg.writeln;
//                                throw (ex);
//                            }
//                        }
//
//                        if (allCells.empty){
//                            generateLabels;
//                        } else {
//                            writeln ("allcells: ",allCells);
//                            foreach (cell; allCells[$-1]) {
//                                generateLabels (cell);
//                            }
//                        }
//                    }
//
//                    if (!allCells.empty) currentRoot = allCells[0][0];
//
//                    // create indexes
//                    if (!fieldNames.empty) db.createIndex (fieldNames);
//
//                }



                // manual labels
                {
                    LinkCell processLabel(Tag label) {
                        // create cell for label
                        auto cell = createLinkCell (labelGroupName,label.tags["display-text"][0].values[0].get!string,
                                                                   label.maybe.tags["condition"].empty ? "" : label.tags["condition"][0].values[0].get!string,
                                                                  (label.getAttribute!bool("total",false)) ? Yes.isTotal : No.isTotal);

                        cell.noChildren = label.getAttribute!bool("no-children",false);

                        // check if write back is set
                        if (!label.maybe.tags["write-back"].empty){
                            writeBack[cell.id] = label.maybe.tags["write-back"][0].values[0].get!string;
                        }

                        // read in a new calculation and set, or use the already set calculation
                        if (!label.maybe.tags["calculation"].empty) {
                            cell.calculation = label.tags["calculation"][0].readCalc;
                        } else {
                            if (!currentCalc.empty) cell.calculation = currentCalc;
                        }

                        // read in conditional formats
                        foreach (tag; label.maybe.tags["conditional-format"]){
                            auto cond = tag.tags["condition"][0].values[0].get!string;
                            CellFormat cf;

                            // font
                            foreach (fontTag; tag.maybe.tags["font"]){
                                if (!fontTag.maybe.tags["font-name"].empty) cf.fontName = fontTag.tags["font-name"][0].values[0].get!string;
                                if (!fontTag.maybe.tags["font-size"].empty) cf.fontSize = fontTag.tags["font-size"][0].values[0].get!int;
                                if (!fontTag.maybe.tags["font-colour"].empty) cf.fontColour = fontTag.tags["font-colour"][0].values[0].get!string;
                                if (!fontTag.maybe.tags["font-weight"].empty) cf.fontWeight = cast(FontWeight)fontTag.tags["font-weight"][0].values[0].get!string;
                            }

                            // alignment
                            foreach (alignTag; tag.maybe.tags["alignment"]){
                                if (!alignTag.maybe.tags["horizontal"].empty) cf.horizontal = cast(HAlignment)alignTag.tags["horizontal"][0].values[0].get!string;
                                if (!alignTag.maybe.tags["vertical"].empty) cf.vertical = cast(VAlignment)alignTag.tags["vertical"][0].values[0].get!string;
                            }

                            // misc
                            if (!tag.maybe.tags["bg-colour"].empty) cf.bgColour = tag.tags["bg-colour"][0].values[0].get!string;
                            if (!tag.maybe.tags["wrap"].empty) cf.wrap = tag.tags["wrap"][0].values[0].get!bool;

                            conditionalFormats[cell.id] ~= ConditionFormat(cond,cf);

                        }

                        // create cells for any nested labels
                        LinkCell rootCell = null;
                        LinkCell currentCell = null;
                        foreach (l; label.maybe.tags["label"]){
                            auto nestedCell = processLabel (l);

                            if (rootCell is null) {
                                rootCell = nestedCell;
                                currentCell = rootCell;
                            } else {
                                currentCell.next = nestedCell;
                                currentCell = currentCell.next;
                            }
                        }

                        if (rootCell !is null) cell.child = rootCell;

                        return cell;
                    }

                    LinkCell rootCell = null;
                    LinkCell currentCell = null;
                    foreach (label; labelGroupTag.maybe.tags["label"]){
                        auto cell = processLabel (label);
                        if (rootCell is null){
                            rootCell = cell;
                            currentCell = cell;
                        } else {
                            currentCell.next = cell;
                            currentCell = cell;
                        }
                    }

                    if (rootCell !is null) {
                        assert (currentRoot is null);
                        currentRoot = rootCell;
                    }
                }
                rootCells[labelGroupName] = currentRoot;
            }
        }



        // load variables
        {
            foreach (variableTag; root.maybe.tags["variable"]){
                TableVariable dtv;
                dtv.name = variableTag.values[0].get!string;
                dtv.select = variableTag.tags["select"][0].values[0].get!string;
                if ("where" in variableTag.tags) dtv.where = variableTag.tags["where"][0].values[0].get!string;
                foreach (tag; variableTag.maybe.tags["by"]) {
                    assert (tag.values[0].get!string in rootCells);
                    dtv.by ~= tag.values[0].get!string;
                }
                generateSqls (dtv);
                variables.add (dtv);
            }
        }

        // load custom functions
        {
            foreach (functionTag; root.maybe.tags["functions"]){
                foreach (tag; functionTag.tags[""]){
                    luaFunctions ~= tag.values[0].get!string;
                }
            }
        }

        // set row/column orientation
        {
            auto currentGroupName="";
            foreach (ori; ["rows","columns"]){
                foreach (tag; root.tags[ori][0].tags[""]){
                    auto groupName = tag.values[0].get!string;
                    enforce (groupName in rootCells, "Label group " ~ groupName ~ " not found.");
                    labelGroupOrientation[(ori=="rows") ? Orientation.row : Orientation.column] ~= groupName;

                    if (!currentGroupName.empty) {
                        setChildren(rootCells[currentGroupName],rootCells[groupName]);
                    }
                    currentGroupName = groupName;
                }
            }
        }



        // set cell lookups
        {
            void traverse (LinkCell cell){
                cellLookup[cell.id] = cell;
                if (cell.child) {
                    traverse (cell.child);
                }

                if (cell.next){
                    traverse (cell.next);
                }
            }
            traverse (rootCell);
        }

        // read headers
        {
            foreach (tag; root.tags["row-headers"][0].tags[""]){
                headers ~= tag.values[0].get!string;
            }
        }
    }

    void generateSqls (ref TableVariable dtv){
        writeln ("Generating SQL");

        alias WhereResult = Tuple!(int,"id",string,"where");
        auto getSqlTree (LinkCell c){
            WhereResult[][] branches;
            WhereResult[] branch;
            void process (LinkCell cell){
                branch ~= WhereResult(cell.id,cell.where);
                if (cell.child !is null) {
                    process (cell.child);
                } else {
                    branches ~= branch;
                    branch.length--;
                }

                if (cell.next !is null) {
                    process (cell.next);
                } else {
                    if (branch.length > 0)
                    branch.length--;
                }
            }
            process (c);
            return branches;
        }

        WhereResult[][][] trees;

        auto tree = dtv.by.map!(a => getSqlTree(rootCells[a]));

        // TODO - limit is up to 7 nested labels for a table. Is there a better way to do the below?

        if (tree.count == 1){
            trees = tree[0].map!(a => [a]).array;
        } else if (tree.count == 2){
            trees = cartesianProduct(tree[0],tree[1]).map!(a => [a[0],a[1]]).array;
        } else if (tree.count == 3){
            trees = cartesianProduct(tree[0],tree[1],tree[2]).map!(a => [a[0],a[1],a[2]]).array;
        } else if (tree.count == 4){
            trees = cartesianProduct(tree[0],tree[1],tree[2],tree[3]).map!(a => [a[0],a[1],a[2],a[3]]).array;
        } else if (tree.count == 5){
            trees = cartesianProduct(tree[0],tree[1],tree[2],tree[3],tree[4]).map!(a => [a[0],a[1],a[2],a[3],a[4]]).array;
        } else if (tree.count == 6){
            trees = cartesianProduct(tree[0],tree[1],tree[2],tree[3],tree[4],tree[5])
                                    .map!(a => [a[0],a[1],a[2],a[3],a[4],a[5]]).array;
        } else if (tree.count == 7){
            trees = cartesianProduct(tree[0],tree[1],tree[2],tree[3],tree[4],tree[5],tree[6])
                                    .map!(a => [a[0],a[1],a[2],a[3],a[4],a[5],a[6]]).array;
        } else if (tree.count > 7){
            assert(0);
        }

        if (tree.count == 0){
            // means variable is constant
            auto sql = format ("select %s from data where %s",dtv.select,
                                            chain([dtv.where],[conditions.joiner(" AND ").array.to!(string)]) // apply where clause to main branch clause
                                            .filter!(a => !a.empty) // filter out cases where "where" may be empty
                                            .joiner(" AND "));
            try {
                dtv.addResult (db.execute (sql).cached,[]);
            } catch (Exception ex){
                sql.writeln;
                ex.msg.writeln;
                throw (ex);
            }
        } else {
            foreach (branch; trees.map!(a => a.joiner)){
                auto sql = format ("select %s from data where %s",dtv.select,
                                                                chain(branch.map!(a => a.where),[dtv.where],[conditions.joiner(" AND ").array.to!(string)]) // apply where clause to main branch clause
                                                                .filter!(a => !a.empty) // filter out cases where "where" may be empty
                                                                .joiner(" AND "));


                try {
                    dtv.addResult (db.execute (sql).cached,branch.map!(a => a.id).array);
                } catch (Exception ex){
                    sql.writeln;
                    ex.msg.writeln;
                    throw (ex);
                }
            }
        }
    }

    auto toGrid(){
        writeln ("toGrid");

        auto titleOffset() {
            return (title.empty) ? 0 : 1;
        }

        int[] currentBranchReference;
        currentBranchReference.length=1_00;
        currentBranchReference.length=0;
        currentBranchReference.assumeSafeAppend;

        LinkCell currentCalcCell;

        int[100] currentBranchReference2;
        auto cellIndex=0;

        //auto row=columnLabelHeight;
        auto row=columnLabelHeight + titleOffset;

        auto savedRow = titleOffset;
        auto col=0;

        GridCell[Point] rvalue;

        int[Point] colSpan;
        int[Point] rowSpan;

        void traverse (LinkCell cell){
            auto point = Point(row,col);

            if (point !in rvalue){
                rvalue[point] = gridLabel (point,cell.text,formats["label"]);

                if (cell.isTerminal){
                    // last column label
                    //rowSpan[point] = rootCell.fastGetDepth(cell);
                    rowSpan[point] = (columnLabelHeight-point.row);
                }
            }

            if (!cell.calculation.empty) currentCalcCell = cell;

            if (cell.isTerminal) {
                currentBranchReference ~= cell.id;

                // COMPLETED BRANCH

                // update lua state

                updateLua (currentBranchReference);

                bool resultOK;
                alias Result = Tuple!(double,"number",string,"text");
                Result result;
                try {
                    lua.doString(format("result = %s",currentCalcCell.calculation.formula));
                    if (currentCalcCell.calculation.format.style==NumberStyle.text){
                        result.text = lua.get!string("result");
                        resultOK=true;
                    } else {
                        result.number = lua.get!double("result");
                        import std.math : isInfinity, isNaN;
                        if (result.number.isNaN || result.number.isInfinity){
                            resultOK = false;
                        } else {
                            resultOK = true;
                        }
                    }
                } catch (Exception ex) {
                    resultOK=false;
                }

                auto dataPoint = Point(savedRow,col);

                assert (dataPoint !in rvalue);

                // apply data, total, or alternate row formatting
                auto dataFormat = formats["data"];
                if (dataPoint.row % 2) dataFormat.mergeCellFormat (formats["alternate"]);
                auto cf = currentBranchReference.canFind!(a => cellLookup[a].isTotal) ? formats["total"] : dataFormat;

                if (resultOK){
                    // check cond. formatting
                    lua["CellValue"] = result.number; // TODO what about text?
                    lua["CellText"] = result.text;
                    foreach (formatRng; currentBranchReference.filter!(a => a in conditionalFormats).map!(a => conditionalFormats[a])){
                        foreach (condFormat; formatRng){
                            lua.doString(format ("condEval = (%s)",condFormat.condition));
                            if (lua.get!bool("condEval")){
                                cf.mergeCellFormat (condFormat.format);
                            }
                        }
                    }

                    // create gridded point
                    if (currentCalcCell.calculation.format.style==NumberStyle.text){
                        rvalue[dataPoint] = gridText(dataPoint,result.text,cf);
                    } else {
                        rvalue[dataPoint] = gridData(dataPoint,result.number,cf,currentCalcCell.calculation);
                    }



                    // check if need to write back to database
                    foreach (cellid; currentBranchReference){
                        auto fieldToUpdate = writeBack.get (cellid,"");
                        if (!fieldToUpdate.empty){
                            // TODO only works when updating numeric values
                            auto sql = format ("update data set %s=%s where %s",fieldToUpdate,
                                                                                result.number,
                                                                                getWhere (currentBranchReference));
                            db.execute (sql);
                        }
                    }
                } else {
                    cf.mergeCellFormat (formats["na"]);
                    rvalue[dataPoint] = gridNA (dataPoint,cf);
                }


                currentBranchReference.length--;
                currentBranchReference.assumeSafeAppend;
            }

            if (cell.child) {
                currentBranchReference ~= cell.id;
                if (isRow(cell) && isColumn(cell.child)) {
                    auto depth = rootCell.fastGetDepth(cell);
                    col += depth+1;
                    colSpan[point]=depth;
                    savedRow=row;
                    row = titleOffset;
                }

                if (isRow(cell) && isRow(cell.child)) {
                    auto depth = rootCell.fastGetDepth(cell);
                    col += depth+1;
                    colSpan[point]=depth;
                    rowSpan[point]=fastGetBreadth(&sameOrientation,cell);
                }

                if (isColumn(cell)) {
                    auto depth = rootCell.fastGetDepth(cell);
                    row += depth+1;
                    //rowSpan[point]=depth;
                    colSpan[point]=fastGetBreadth(&sameOrientation,cell);
                }

                traverse (cell.child);



                if (isRow(cell) && isColumn(cell.child)) {
                    col -= rootCell.fastGetDepth(cell)+1;
                    row=savedRow;
                }
                if (isRow(cell) && isRow(cell.child)) col -= rootCell.fastGetDepth(cell)+1;
                if (isColumn(cell)) row -= rootCell.fastGetDepth(cell)+1;

                currentBranchReference.length--;
                currentBranchReference.assumeSafeAppend;
            }

            if (cell.next) {
                if (isRow(cell)) {
                    auto breadth=fastGetBreadth(&sameOrientation,cell);
                    row += breadth+1;
                    rowSpan[point] = breadth;
                }
                if (isColumn(cell)) {
                    auto breadth=fastGetBreadth(&sameOrientation,cell);
                    col += breadth+1;
                    colSpan[point] = breadth;
                    //rowSpan[point] = rootCell.fastGetDepth(cell);
                }
                traverse (cell.next);
                if (isRow(cell)) row -= fastGetBreadth(&sameOrientation,cell)+1;
                if (isColumn(cell)) col -= fastGetBreadth(&sameOrientation,cell)+1;
            }
        }

        import std.datetime.stopwatch;
        auto sw = StopWatch();
        sw.start;

        writeln ("traversing");

        traverse (rootCell);
        writeln ("done");
        sw.stop;

        //writeln (sw.total.msecs," total time");
        //writeln (tsw.peek.msecs," lua time.");


        // add header row labesl
        foreach (j,header; headers){
            auto i = j.to!int; // TODO
            rvalue[Point(0+titleOffset,0+i)] = gridLabel (Point(0+titleOffset,0+i),header,formats["label"]);
            rvalue[Point(0+titleOffset,0+i)].rowSpan=columnLabelHeight-1;

            // check if the last header needs to span across multiple columns
            //if (i==headers.length && i<rowLabelWidth){
            if (header == headers[$-1] && i<rowLabelWidth){
                rvalue[Point(0+titleOffset,0+i)].colSpan = rowLabelWidth-i;
            }
        }

        foreach (ref gridCell; rvalue.byValue){
            if (gridCell.point in colSpan){
                gridCell.colSpan = colSpan[gridCell.point];
            }

            if (gridCell.point in rowSpan){
                gridCell.rowSpan = rowSpan[gridCell.point];
            }
        }

        // add title
        if (!title.empty) {
            rvalue[Point(0,0)] = gridTitle (Point(0,0),title,formats["title"]);
            rvalue[Point(0,0)].colSpan = rvalue.byKey.map!(a => a.col).maxElement;
        }

//        rvalue.values.map!(a => Tuple!(Point,string,int,int)(a.point,a.text,a.rowSpan,a.colSpan)).array
//                                                                                                 .multiSort!("a[0].row < b[0].row",
//                                                                                                             "a[0].col < b[0].col")
//                                                                                                 .array
//                                                                                                 .each!(a => a.writeln);
//
//        readln;




        return rvalue;
    }
}
