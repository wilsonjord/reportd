// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module basically wraps the d2sqlite3 database module, for the purpose of mapping a csv to a in memory database.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.database;

import std.algorithm;
import std.range;
import std.stdio : writeln;
import std.typecons : Tuple, Flag, Yes, No;;
import std.file : readText;
import std.format: format;
import std.array : array, replace;
import std.conv : to;
import std.string : isNumeric;

import reportd.fastcsv;
import d2sqlite3;
import sdlang;

enum FieldType {text="text",numeric="numeric"}
alias Field = Tuple!(string,"name",int,"column",FieldType,"type");

auto isNumeric (Field field) { return field.type==FieldType.numeric; }
auto isText (Field field) { return field.type==FieldType.numeric; }
auto nullValue (Field field) {
    return (field.isNumeric) ? "null" : "''";
}

auto createDB(T) (string config, T data, Flag!"ignoreConvErrors" flag = No.ignoreConvErrors){
    auto stringToSqlValue(S) (S v) {
        if (v.empty) return "null";

        if (v.isNumeric) {
            return v[];
        } else {
            if (flag == Yes.ignoreConvErrors) {
                return format("\"%s\"",v);
            } else {
                throw new Exception ("Error coverting text to numeric. value is: "~v.to!string);
            }
        }
    }

    assert (!data.empty);

    // load data as a database
    // data will most usual be a result of reading in a CSV file, or data copied from clipboard, etc.

    auto db = Database(":memory:");

    // load extensions, so can use "median" in SQL
//    sqlite3_enable_load_extension(db.handle, 1);
//    if (SQLITE_ERROR == sqlite3_load_extension (db.handle,"extension-functions.dll","sqlite3_extension_init",null))
//        throw new Exception ("DB extension not loaded");

    auto root = parseSource(config.readText);

    // read data-based fields first
    import std.exception : enforce;
    Field[] fields;
    foreach (tag; root.tags["field"]){
        if (!tag.maybe.attributes["column"].empty){
            Field field;
            field.name = tag.values[0].get!string;
            enforce ("type" in tag.attributes, field.name ~ " needs a type, please check your database configuration SDL file.");
            auto t=tag.attributes["type"][0].value.get!string;
            switch (t) {
                default: throw new Exception (format("Type '%s' for field '%s' not understood.",t,field.name));
                case "text": field.type=FieldType.text; break;
                case "numeric": field.type=FieldType.numeric; break;
            }
            field.column = tag.attributes["column"][0].value.get!int;
            fields ~= field;
        }
    }

    // create table
    try {
        db.execute (fields.createSqls);
    } catch (Exception ex){
        fields.createSqls.writeln;
        ex.msg.writeln;
        throw (ex);
    }

    // populate table

    db.begin;
    import std.exception : assumeUnique;
    immutable fieldNames = fields.map!(a => a.name).joiner(",").array.assumeUnique;
    foreach (row; data){
        auto fieldValues = fields.map!(a => a.column > row.length ? a.nullValue :
                                                                   (a.isNumeric) ? stringToSqlValue!(const char[])(row[a.column-1]) :
                                                                                   format("'%s'",row[a.column-1].replace("'","''"))).joiner(",");

        auto sql = format ("INSERT INTO data (%s) VALUES (%s);",fieldNames,
                                                                fieldValues);

        try {
            db.execute (sql);
        } catch (Exception ex){
            sql.writeln;
            ex.msg.writeln;
            throw (ex);
        }
    }
    db.commit;


    db.begin;
    foreach (tag; root.tags["field"]){
        if (!tag.maybe.attributes["formula"].empty){
            auto name = tag.values[0].get!string;
            // create field
            db.execute (format ("ALTER TABLE data ADD %s %s;",name,(tag.values[0].get!string=="numeric") ? "REAL" : "TEXT"));

            // populate field
            db.execute (format("UPDATE data SET %s=%s;",name,tag.attributes["formula"][0].value.get!string));
        }
    }
    db.commit;

    // finally, process any custom data manipulation
    db.begin;
    foreach (tag; root.maybe.tags["update"]){
        db.execute (format("UPDATE data SET %s;",tag.values[0].get!string));
    }
    db.commit;

    return db;
}

auto createDBFromCSV (string csvFile, string configFile, Flag!"ignoreConvErrors" flag = No.ignoreConvErrors){
    auto data = csvFromUtf8File (csvFile);
    data.popFront;
    return createDB (configFile,data,flag);
}

auto createSqls (Field[] f){
    return format ("CREATE TABLE data(%s);",f.map!(a => format("%s %s",a.name,(a.type==FieldType.numeric) ? "real" : "text")).joiner(","));
}

auto hasField (Database db, string f){
    return db.execute("PRAGMA table_info(data);").map!(a => a.peek!string(1)).canFind(f);
}

auto createIndexName (string[] fields){
    return fields.joiner("_").array ~ "_index";
}

void createIndex(T) (auto ref T db, string[] fields){
    auto indexName = fields.createIndexName;

    auto currentIndexes = db.execute ("select * FROM sqlite_master WHERE type = 'index'");
    foreach (row; currentIndexes){
        if (indexName.to!string == row.peek!string(1)) return; // index already exsist
    }

    // create index
    auto sql = format ("create index %s on data (%s)",indexName,fields.joiner(","));
    sql.writeln;
    db.execute (sql);

}

void dump (Database db, string filePath){
    import std.stdio : File;
    auto file = File (filePath,"w");

    // write header
    file.writeln (db.execute ("PRAGMA table_info(data);").map!(a => a.peek!string(1)).joiner("\t"));

    // write contents
    foreach (row; db.execute ("select * from data")){
        foreach (val; row){
            file.write (val);
            file.write ("\t");
        }
        file.writeln;
    }
}


