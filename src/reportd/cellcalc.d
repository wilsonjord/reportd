// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module implements cell calculation and formatting.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.cellcalc;

import std.typecons : Tuple, Nullable;

enum NumberStyle : string {number="number", percent="percent", text="text"}
enum NumberScale : string {none="1",thousand="1000",million="1000000"}

/**
    Stores how a numeric value should be displayed.
*/
alias NumberFormat = Tuple!(NumberStyle,"style",
                            NumberScale,"scale",
                            int,"decimalPlaces",
                            bool,"withComma");
///
unittest {
    NumberFormat a;
    a.style = NumberStyle.number;
    NumberFormat b;
    b.style = NumberStyle.number;
    assert (a==b);

    b.style = NumberStyle.percent;
    assert (a!=b);
}

/**
    Stores how a cell should be calculated, and how the result of that calculation should be displayed.
*/
struct CellCalc {
    private:
        NumberFormat _format;
        string _formula;
        bool _overRide=false;

    public:
        @property auto overRide() {return _overRide;}
        @property void overRide(bool newval) {_overRide = newval;}
        @property ref format() {return _format;}
        @property format(NumberFormat newval) {_format = newval;}
        @property string formula() const {
            return _formula;
        }

        @property void formula(string newval) {
            _formula = newval;
        }

        @property auto empty() {
            import std.range : empty;
            return _formula.empty;
        }

        /**
        Constructor taking a text representation of a calculation/formula.
        Params:
            formula = range or string representing the file _name
        */

        this (string formula){
            _formula = formula;
        }

        bool opEquals (in CellCalc c) const {
            return (_formula == c.formula);
        }

        string toString() {
            return _formula;
        }
}

///
unittest {
    CellCalc a;
    a.formula = "1+1";
    auto b = CellCalc("1+1");
    assert (a == b);

    a.format.scale = NumberScale.none;
    b.format.scale = NumberScale.thousand;
    assert (a == b); // equality is formula based only
}
