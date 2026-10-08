% Family Tree Program
% Replace the sample people or add new facts as needed.

% parent(Parent, Child).
parent(george, john).
parent(helen, john).
parent(george, susan).
parent(helen, susan).

parent(john, mary).
parent(linda, mary).
parent(john, robert).
parent(linda, robert).

parent(susan, emily).
parent(david, emily).
parent(susan, michael).
parent(david, michael).

parent(mary, sophia).
parent(paul, sophia).
parent(robert, liam).
parent(anna, liam).

% male(Person).
male(george).
male(john).
male(david).
male(robert).
male(michael).
male(paul).
male(liam).

% female(Person).
female(helen).
female(susan).
female(linda).
female(mary).
female(emily).
female(anna).
female(sophia).

% child(Child, Parent): Child is a child of Parent.
child(Child, Parent) :-
    parent(Parent, Child).

% grandparent(Grandparent, Grandchild): two parent links connect them.
grandparent(Grandparent, Grandchild) :-
    parent(Grandparent, Parent),
    parent(Parent, Grandchild).

% sibling(Person1, Person2): the people share at least one parent.
% once/1 prevents duplicate success when they share two parents.
sibling(Person1, Person2) :-
    dif(Person1, Person2),
    once((parent(Parent, Person1), parent(Parent, Person2))).

% cousin(Person1, Person2): their parents are siblings.
cousin(Person1, Person2) :-
    dif(Person1, Person2),
    parent(Parent1, Person1),
    parent(Parent2, Person2),
    sibling(Parent1, Parent2).

% descendant(Descendant, Ancestor): direct descendants form the base case.
descendant(Descendant, Ancestor) :-
    parent(Ancestor, Descendant).

% A person is also a descendant when the person's parent is a descendant.
descendant(Descendant, Ancestor) :-
    parent(Ancestor, Intermediate),
    descendant(Descendant, Intermediate).

