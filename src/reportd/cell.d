// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module provides the Cell functionality.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.cell;

import reportd.node;
import reportd.cellcalc;

import std.typecons : Flag,Yes,No;
import std.algorithm : equal;

alias CellID = int;

// Class gives nice reference semantics that's easy to reason about
class Cell {
    immutable string groupName;
    string text;
    string where;
    bool isTotal;
    bool noChildren;

    CellCalc calculation;

    static CellID _ID=0;
    immutable CellID id;

    this (string g, string t="", string w="", Flag!"isTotal" flag = No.isTotal){
        groupName = g;
        text = t;
        where = w;
        isTotal = (flag==Yes.isTotal) ? true : false;
        id = _ID++;
    }

    override string toString(){
        return groupName ~ " " ~ text ~ " " ~ where;
    }
}

alias LinkCell = Node!Cell;

auto createLinkCell (string g, string t="", string w="",Flag!"isTotal" flag = No.isTotal) {
    return new LinkCell (new Cell (g,t,w,flag));
}

auto isTerminal (LinkCell c) { return c.child is null; }
auto groupValue(T : LinkCell)(T t) { return t.groupName; }
auto groupValue(T : string)(T t) { return t; }
auto sameGroup(T,S) (T a, S b) {
    static assert (is(T == string) || is(T == LinkCell));
    static assert (is(S == string) || is(S == LinkCell));
    return equal (groupValue(a),groupValue(b));
}

void prettyPrint (LinkCell c) {
    import std.stdio : writeln;
    c.writeln;

    if (c.child) {
        writeln ("printing children of ",c.text);
        prettyPrint (c.child);
    }

    if (c.next) prettyPrint (c.next);
}

unittest {
    auto a = new LinkCell (new Cell ("test")); // "test" is groupName
    auto b = new LinkCell (new Cell ("test"));
    assert (sameGroup("test","test"));
    assert (sameGroup("test",a));
    assert (sameGroup(b,"test"));
    assert (sameGroup(a,b));
}

auto getWhere (LinkCell c) { return (c is null) ? "" : c.where; }

unittest {
    auto a = new LinkCell (new Cell (""));
    assert (a.isTerminal);
}


/++ Sets c as the lowest level child of the Cell, and any next +/
void setChildren (LinkCell dc, LinkCell c) {
    if (dc is c) return;  // don't link cell to itself

    if (!dc.noChildren){
        if (dc.child is null){
            dc.child = c;
        } else {
            dc.child.setChildren(c);
        }
    }

    if (dc.next !is null){
        dc.next.setChildren (c);
    }
}

///
unittest {
    auto a = createLinkCell ("a");
    auto b = createLinkCell ("b");
    auto c = createLinkCell ("c","some text");

    a.next = b;
    a.setChildren (c);
    assert (a.child.text == "some text");
    assert (a.child == b.child);

    auto d = createLinkCell ("d");

    a.setChildren (d);
    assert (a.child.groupName == "c");
    assert (a.child.child.groupName == "d");
    assert (b.child.child.groupName == "d");
}
