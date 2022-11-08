grammar Calculator;


/*
 * Parser Rules
 */

compileUnit : expr EOF;

expr
	:'(' expr ')' #Parent
	| ('-') right = expr #Neg
	| left = expr op = ('*' | ':' | '/' | '%') right = expr #MulDiv
	| left = expr op = ('+' | '-') right = expr #AddSub
	| left = expr op = ('>' | '>=' | '==' | '<' | '<=') right = expr #Compare
	| NUMBER #Num
	| IDENTIFIER #Ident
	;


/*
 * Lexer Rules
 */

NUMBER : INT;
IDENTIFIER : [A-Z]+[0-9]+;

INT : ('-')? ('0'..'9')+;

MULTIPLY: '*';
DIVIDE: ':';
DIV: '/';
MOD: '%';
ADD: '+';
SUBSTRUCT: '-';
EQUAL: '==';
GREATER: '>';
GREATOREQ: '>=';
LESS: '<';
LESSOREQ: '<=';
LPAREN : '(';
RPAREN : ')';

WS : [ \t\r\n] -> skip;
