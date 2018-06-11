// Written in the D programming language
// Jordan K. Wilson https://wilsonjord.github.io/

/**
This module implements a generic node-like object.

Author: $(HTTP wilsonjord.github.io, Jordan K. Wilson)

*/

module reportd.node;


class Node(T) {
    private:
        Node!T _next;
        Node!T _previous;
        Node!T _parent;
        Node!T _child;
        T _payload;

    public:
        @property @nogc {
            auto next() { return _next; }
            void next(Node!T newval) {
                assert (this !is newval);
                _next=newval;
                if (newval !is null){
                    newval.previous=this;
                    newval.parent = this.parent;
                }
            }
            auto child() { return _child; }
            void child(Node!T newval) {
                assert (this !is newval);
                _child=newval;
                if (newval !is null){
                    newval.parent = this;
                }
            }
            auto previous()  { return _previous; }
            void previous(Node!T newval)  { _previous=newval; }
            auto parent()  { return _parent; }
            void parent(Node!T newval)  { _parent=newval; }
            auto payload()  { return _payload; }
            void payload(T newval)  { _payload = newval; }
        }

        this(){}

        this (T val) @nogc {
            _payload = val;
        }

        override string toString() {
            import std.conv : to;
            return _payload.to!string;
        }

        alias payload this;
}

///
unittest {
    auto node = new Node!int(1);
    assert (node.payload == 1);

    node.next = new Node!int(2);
    assert (node.next.payload == 2);
    assert (node.next.previous.payload == 1);
}

