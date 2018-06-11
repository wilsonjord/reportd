// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module implements colouring and formatting for a cell.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.cellformat;

import std.typecons;
import std.file;
import std.conv;

/** Controls the horizontal alignment of the cell content. */
enum HAlignment {centre="centre",left="left",right="right"}
/** Controls the vertical alignment of the cell content. */
enum VAlignment {middle="middle",top="top",bottom="bottom"}
/** Controls the weight of a font. */
enum FontWeight {normal="normal",bold="bold"}

/**
    A CellFormat object contains information on how a cell looks, using font (type, size, colour, and weight), text alignment, and border colour.
*/

// TODO: upcoming release will see "apply" for Nullable type, investigate if that can simplify use of Nullable below
struct CellFormat {
    private:
        Nullable!string _fontName;
        Nullable!int _fontSize;
        Nullable!string _fontColour;
        Nullable!FontWeight _fontWeight;
        Nullable!string _bgColour;
        Nullable!bool _wrap;
        Nullable!HAlignment _horizontal;
        Nullable!VAlignment _vertical;

    public:
        @property {
            /** Font used to display cell contents. */
            auto fontName() { return _fontName; }
            void fontName(string newval) { _fontName = newval; }

            /** Font size */
            auto fontSize() { return _fontSize; }
            void fontSize(int newval) { _fontSize = newval; }

            /** Font colour

                Params:
                    colour = The font colour of a cell, as a string based hex representation.

                Example:
                    ```d
                     CellFormat style;
                     style.fontColour = "#0000ff" // blue text
                     style.fontColour = "#ff0000" // red text```
             */
            void fontColour(string colour) { _fontColour = colour; }
            ///
            auto fontColour() { return _fontColour; }

            /** Font weight
                Params:
                    weight = The weight of the font

            */
            void fontWeight(FontWeight weight) { _fontWeight = weight; }
            ///
            auto fontWeight() { return _fontWeight; }

            /** Background colour
                Params:
                    colour = The background colour of a cell, as a string based hex representation.

                Example:
                    ```d
                     CellFormat style;
                     style.bgColour = "##e8e8e8" // light grey background```
            */

            void bgColour(string colour) { _bgColour = colour; }
            ///
            auto bgColour() { return _bgColour; }

            auto wrap() { return _wrap; }
            void wrap(bool newval) { _wrap = newval; }

            auto horizontal() { return _horizontal; }
            void horizontal(HAlignment newval) { _horizontal = newval; }

            auto vertical() { return _vertical; }
            void vertical(VAlignment newval) { _vertical = newval; }

            auto isBold() { return _fontWeight==FontWeight.bold; }
        }


        bool opEquals (in CellFormat a) const {
            auto b = cast(CellFormat)a;
            if (_fontName.isNull != b.fontName.isNull) return false;
            if (_fontSize.isNull != b.fontSize.isNull) return false;
            if (_fontColour.isNull != b.fontColour.isNull) return false;
            if (_fontWeight.isNull != b.fontWeight.isNull) return false;
            if (_bgColour.isNull != b.bgColour.isNull) return false;
            if (_wrap.isNull != b.wrap.isNull) return false;
            if (_horizontal.isNull != b.horizontal.isNull) return false;
            if (_vertical.isNull != b.vertical.isNull) return false;

            if (!_fontName.isNull && (_fontName != b.fontName)) return false;
            if (!_fontSize.isNull && (_fontSize != b.fontSize)) return false;
            if (!_fontColour.isNull && (_fontColour != b.fontColour)) return false;
            if (!_fontWeight.isNull && (_fontWeight != b.fontWeight)) return false;
            if (!_bgColour.isNull && (_bgColour != b.bgColour)) return false;
            if (!_wrap.isNull && (_wrap != b.wrap)) return false;
            if (!_horizontal.isNull && (_horizontal != b.horizontal)) return false;
            if (!_vertical.isNull && (_vertical != b.vertical)) return false;

            return true;
        }

        int opCmp (CellFormat b) {
            if (opEquals(b)) return 0;
            if (fontName.isNull != b.fontName.isNull) return (fontName.isNull > b.fontName.isNull) ? 1 : -1;
            if (fontSize.isNull != b.fontSize.isNull) return (fontSize.isNull > b.fontSize.isNull) ? 1 : -1;
            if (fontColour.isNull != b.fontColour.isNull) return (fontColour.isNull > b.fontColour.isNull) ? 1 : -1;
            if (fontWeight.isNull != b.fontWeight.isNull) return (fontWeight.isNull > b.fontWeight.isNull) ? 1 : -1;
            if (bgColour.isNull != b.bgColour.isNull) return (bgColour.isNull > b.bgColour.isNull) ? 1 : -1;
            if (wrap.isNull != b.wrap.isNull) return (wrap.isNull > b.wrap.isNull) ? 1 : -1;
            if (horizontal.isNull != b.horizontal.isNull) return (horizontal.isNull > b.horizontal.isNull) ? 1 : -1;
            if (vertical.isNull != b.vertical.isNull) return (vertical.isNull > b.vertical.isNull) ? 1 : -1;

            if (!fontName.isNull && (fontName != b.fontName)) return (fontName > b.fontName) ? 1 : -1;
            if (!fontSize.isNull && (fontSize != b.fontSize)) return (fontSize > b.fontSize) ? 1 : -1;
            if (!fontColour.isNull && (fontColour != b.fontColour)) return (fontColour > b.fontColour) ? 1 : -1;
            if (!fontWeight.isNull && (fontWeight != b.fontWeight)) return (fontWeight > b.fontWeight) ? 1 : -1;
            if (!bgColour.isNull && (bgColour != b.bgColour)) return (bgColour > b.bgColour) ? 1 : -1;
            if (!wrap.isNull && (wrap != b.wrap)) return (wrap > b.wrap) ? 1 : -1;
            if (!horizontal.isNull && (horizontal != b.horizontal)) return (horizontal > b.horizontal) ? 1 : -1;
            if (!vertical.isNull && (vertical != b.vertical)) return (vertical > b.vertical) ? 1 : -1;

            assert(false);

        }

        string toString(){
            return (!fontName.isNull ? fontName.to!string : "") ~ " " ~
                   (!fontSize.isNull ? fontSize.to!string : "") ~ " " ~
                   (!fontColour.isNull ? fontColour.to!string : "") ~ " " ~
                   (!fontWeight.isNull ? fontWeight.to!string : "") ~ " ";
        }
}

/**
    Function to merge the formats of two different cells.
    The resultant cell format will be the result of overlaying the null fields of the first cell with the second cell, and
    overlaying the null fields of the second cell with the first cell.
*/
void mergeCellFormat (ref CellFormat a, ref CellFormat b) {
    a.fontName = (b.fontName.isNull) ? a.fontName : b.fontName;
    a.fontSize = (b.fontSize.isNull) ? a.fontSize : b.fontSize;
    a.fontColour = (b.fontColour.isNull) ? a.fontColour : b.fontColour;
    a.fontWeight = (b.fontWeight.isNull) ? a.fontWeight : b.fontWeight;
    a.bgColour = (b.bgColour.isNull) ? a.bgColour : b.bgColour;
    a.wrap = (b.wrap.isNull) ? a.wrap : b.wrap;
    a.horizontal = (b.horizontal.isNull) ? a.horizontal : b.horizontal;
    a.vertical = (b.vertical.isNull) ? a.vertical : b.vertical;
}

alias ConditionalFormat = Tuple!(string,"condition",CellFormat,"cellFormat",string,"branchContains",string,"overrideText");

string getCellFormatName (ref CellFormat[string] formats, CellFormat format) {
    foreach (formatName, formatValue; formats){
        if (formatValue == format) return formatName;
    }
    return "";
}

alias FormatProperty = Tuple!(string,"name",string,"value");
struct FormatStyle {
    private:
        string[string] _properties;
        string[string][] _conditions;

    public:
        @property {
            auto properties() { return _properties;}
            void properties(string[string] newval) { _properties=newval;}
            auto conditions() {return _conditions;}
            void conditions(string[string][] newval) {_conditions=newval;}
        }

        void addProperty (string name, string value) {
            _properties[name] = value;
        }

        void addCondition (string[string] conditionMap) {
            _conditions ~= conditionMap;
        }

        string getProperty (string name) {
            auto ptr = name in _properties;
            if (ptr is null) return null;
            return *ptr;
        }

        bool opEquals (FormatStyle a) {
            //auto fsa = cast(FormatStyle)a;
            auto fsa = a;
            return (properties == fsa.properties) && (conditions == fsa.conditions);
        }
}

struct FormatStyles {
    private:
        FormatStyle[string] _styles;
        int index=0;

    public:
        @property styles() { return _styles; }

        string addStyle (FormatStyle fs){
            foreach (key; _styles.keys){
                if (fs == _styles[key]) return key;
            }
            string id;
            do {
                import std.format;
                id = format("NS%s",index);
                index++;
            } while (id in _styles);

            _styles[id] = fs;
            return id;
        }

        bool getStyle (string name, ref FormatStyle fs ) @nogc {
            auto ptr = name in _styles;
            if (ptr is null) return false;
            fs = *ptr;
            return true;
        }
}
