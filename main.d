module main;

import std.stdio;
import std.typecons;
import d2sqlite3;

import reportd;

void main() {
    auto db = createDBFromCSV (`C:\development\sentencing-analysis\data2.csv`,`C:\development\sentencing-analysis\fields.sdl`);
    //auto db = createDBFromCSV (`C:\development\health.csv`,"better-health-fields.sdl");

    auto report = new Report (`C:\development\sentencing-analysis\report.sdl`,db);

    report.printODS (`C:\development\sentencing-analysis\sentencing.ods`);

}
