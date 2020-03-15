grammar Language;

/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :
	LPAREN expression RPAREN #ParenthesizedExpr
	| operatorToken=SUBTRACT expression #MinusExpr
	| operatorToken=ADD expression #PlusExpr
	| operatorToken=(MAX | MIN) LPAREN expression DESP expression RPAREN #MaxMinExpr
	| expression EXPONENT expression #ExponentialExpr
    | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
	| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
	| operatorToken=(INC | DEC) LPAREN expression RPAREN #IncDecExpr
	| operatorToken=(MOD | DIV) LPAREN expression DESP expression RPAREN #ModDivExpr
	| NUMBER #NumberExpr
	| IDENTIFIER #IdentifierExpr
	; 

/*
 * Lexer Rules
 */

NUMBER : INT ('.' INT)?; 
//IDENTIFIER : [a-zA-Z]+[1-9][0-9]+;

INT : ('0'..'9')+;


MOD: 'mod';
DIV: 'div';
INC: 'inc';
DEC: 'dec';
EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
DESP: ';'|',';
MAX: 'max';
MIN: 'min';

WS : [ \t\r\n] -> channel(HIDDEN);

