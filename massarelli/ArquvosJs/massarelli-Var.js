//ɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚɚ
var Cells = {}
var Texts = {}
var Menus = {}

// ------ PLANILHAS
Cells['LicitPubi'] = [			[   {'nLinPla':53,'nColPla':5,'lplan0':10,'cplan0':10,'lFrz':2,'cFrz':1,'iEl':1,'iEl':1,'cursor': 'outline','enterMove': 'baixo','scrollBars':0,'header': 'header','LinMod': 1,'autoload': 'autoload','autosize': 'autosize','arqDados': 'Editais.js','sincr': '*',} , 0],   			[0, {'valor': '&#211;rg&#227;o','backgroundColor': '#C6E0B4','color': '#000000','fontFamily': 'Calibri','fontSize': '14','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 2, 1, 2, 2,],'mergedCols': 1,'mergedLins': 1,'top': 0,'left': 0,'width': 157,'height': 29}, {'valor': 'Cidade 2','backgroundColor': '#C6E0B4','color': '#000000','fontFamily': 'Calibri','fontSize': '14','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 2, 1, 2, 0,],'mergedCols': 1,'mergedLins': 1,'top': 0,'left': 157,'width': 138,'height': 29}, {'valor': 'Estado','backgroundColor': '#C6E0B4','color': '#000000','fontFamily': 'Calibri','fontSize': '14','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'center','bordas': [ 2, 1, 2, 0,],'mergedCols': 1,'mergedLins': 1,'top': 0,'left': 295,'width': 48,'height': 29}, {'valor': 'Abertura','backgroundColor': '#C6E0B4','color': '#000000','fontFamily': 'Calibri','fontSize': '14','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'center','bordas': [ 2, 1, 2, 0,],'mergedCols': 1,'mergedLins': 1,'top': 0,'left': 343,'width': 72,'height': 29}, {'valor': 'Site','backgroundColor': '#C6E0B4','color': '#000000','fontFamily': 'Calibri','fontSize': '14','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 2, 2, 2, 0,],'mergedCols': 1,'mergedLins': 1,'top': 0,'left': 415,'width': 94,'height': 29} ],			[0, {'valor': '1','backgroundColor': '#8EA9DB','color': '#000000','fontFamily': 'Calibri','fontSize': '11','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 0, 1, 1, 2,],'mergedCols': 1,'mergedLins': 1,'top': 29,'left': 0,'width': 157,'height': 19}, {'valor': 'A','backgroundColor': '#FFC000','color': '#000000','fontFamily': 'Calibri','fontSize': '11','fontWeight': 'normal','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 0, 1, 1, 0,],'mergedCols': 1,'mergedLins': 1,'top': 29,'left': 157,'width': 138,'height': 19}, {'valor': 'B','backgroundColor': '#FFFF00','color': '#000000','fontFamily': 'Calibri','fontSize': '11','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'center','bordas': [ 0, 1, 1, 0,],'mergedCols': 1,'mergedLins': 1,'top': 29,'left': 295,'width': 48,'height': 19}, {'valor': '01/01/2023','backgroundColor': '#FFFFFF','color': '#000000','fontFamily': 'Calibri','fontSize': '11','fontWeight': 'normal','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'center','bordas': [ 0, 1, 1, 0,],'mergedCols': 1,'mergedLins': 1,'top': 29,'left': 343,'width': 72,'height': 19}, {'valor': 'com','backgroundColor': '#FFFFFF','color': '#305496','fontFamily': 'Calibri','fontSize': '12','fontWeight': 'bold','fontStyle': 'normal','textDecorationLine': '','readonly': 'true','textAlign': 'left','bordas': [ 0, 2, 1, 0,],'mergedCols': 1,'mergedLins': 1,'top': 29,'left': 415,'width': 94,'height': 19} ],]

// ------[PLANILHAS]


// ------ TEXTOS

// ------[TEXTOS]



// ------ MENUS
Menus['divMenuHori'] = [
                    // ... format geral
                    [2, 'autoload', 'aba', 'click', 'fitDiv', '', , , ],
                    // ...
                    [0,  1,  'Em An&#225;lise',  '',  ['#808080', 18, 'B', 'I', 'u', 'Arial'] ,['#BFBFBF', 'center', 59, 166, [ 1, 1, 1, 1,]],  ['#595959', 16, 'B', 'I', 'u', 'Arial'] ,['#F2F2F2', 'center', 59, 74, [ 3, 3, 3, 3,]], 0, 0 ], 
                    [1,  1,  'Em Andamento',  '',  ['#808080', 18, 'B', 'I', 'u', 'Arial'] ,['#BFBFBF', 'center', 34, 166, [ 1, 1, 1, 1,]],  ['#595959', 16, 'B', 'I', 'u', 'Arial'] ,['#F2F2F2', 'center', 34, 74, [ 3, 1, 1, 1,]], 0, 0 ], 

                    ]
                    // ... fim de menu
Menus['pop1'] = [
                    // ... format geral
                    [13, 'autoload', 'pop', 'hover', '', '', , , ],
                    // ...
                    [0,  2,  'Cabelo &#8221; a &#8220; b &#8226; c &#8212; d &#8217; e',  '',  ['#000000', 16, 'B', 'i', 'u', 'Calibri'] ,['#92D050', 'left', 23, 364, [ 3, 3, 3, 3,]],  ['#0070C0', 20, 'B', 'I', 'U', 'Arial'] ,['#A9D08E', 'center', 23, 74, [ 3, 1, 1, 1,]], 0, 0 ], 
                    [1,  2,  'Cabe&#231;a &#8221; a &#8220; b &#8226; c &#8226; d',  'YY',  ['#FFC000', 20, 'B', 'i', 'u', 'Arial'] ,['#FFE699', 'left', 34, 364, [ 3, 2, 2, 2,]],  ['#FFC000', 20, 'B', 'i', 'u', 'Arial'] ,['#A9D08E', 'center', 34, 74, [ 1, 1, 2, 1,]], 0, 0 ], 
                    [3,  2,  '&#8226;c&#227;o',  'c&#227;o',  ['#000000', 26, 'B', 'i', 'u', 'Arial'] ,['#C6E0B4', 'left', 45, 364, [ 2, 2, 1, 2,]],  ['#000000', 26, 'B', 'i', 'u', 'Arial'] ,['#A9D08E', 'center', 45, 74, [ 2, 1, 0, 1,]], 0, 0 ], 
                    [1,  2,  '&#230;b . s&#601; . lu&#720; . tli',  '>',  ['#00B050', 20, 'B', 'i', 'u', 'Arial'] ,['#FEFE9A', 'left', 26, 364, [ 1, 2, 1, 2,]],  ['#00B050', 20, 'B', 'i', 'u', 'Arial'] ,['#A9D08E', 'center', 26, 74, [ 0, 1, 0, 1,]], 0, 0 ], 
                    [21,  3,  'Corpo e bra&#231;os',  '>',  ['#D9D9D9', 14, 'b', 'i', 'u', 'Calibri'] ,['#808080', 'left', 26, 192, [ 1, 1, 0, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#D9E1F2', 'center', 26, 74, [ 0, 0, 0, 1,]], 0, 0 ], 
                    [22,  4,  'A',  '',  ['#000000', 11, 'b', 'i', 'u', 'Calibri'] ,['#FFFFFF', 'left', 26, 40, [ 1, 1, 0, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#A9D08E', 'center', 26, 74, [ 0, 0, 0, 0,]], 0, 0 ], 
                    [23,  4,  'i',  '',  ['#000000', 11, 'b', 'i', 'u', 'Calibri'] ,['#FFFFFF', 'left', 26, 40, [ 0, 1, 1, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#A9D08E', 'center', 26, 74, [ 0, 0, 0, 0,]], 0, 0 ], 
                    [24,  3,  'M&#227;o',  '',  ['#D9D9D9', 14, 'B', 'i', 'u', 'Calibri'] ,['#808080', 'left', 26, 192, [ 0, 1, 1, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#8497B0', 'center', 26, 74, [ 0, 0, 1, 1,]], 0, 0 ], 
                    [5,  3,  'unhas',  '',  ['#D9D9D9', 14, 'b', 'i', 'u', 'Calibri'] ,['#808080', 'left', 26, 192, [ 1, 1, 1, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#8497B0', 'center', 26, 74, [ 1, 0, 0, 0,]], 0, 0 ], 
                    [33,  3,  'dedos',  '>',  ['#D9D9D9', 14, 'b', 'i', 'u', 'Calibri'] ,['#808080', 'left', 18, 192, [ 1, 1, 1, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#8497B0', 'center', 18, 74, [ 0, 0, 0, 1,]], 0, 0 ], 
                    [34,  4,  'a',  '',  ['#000000', 11, 'b', 'i', 'u', 'Calibri'] ,['#FEFE9A', 'left', 15, 40, [ 1, 1, 0, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#8497B0', 'center', 15, 74, [ 0, 0, 0, 1,]], 0, 0 ], 
                    [35,  4,  'b',  '',  ['#000000', 11, 'b', 'i', 'u', 'Calibri'] ,['#FEFE9A', 'left', 19, 40, [ 0, 1, 1, 1,]],  ['#FFFFFF', 11, 'b', 'i', 'u', 'Calibri'] ,['#8497B0', 'center', 19, 74, [ 0, 0, 0, 1,]], 0, 0 ], 
                    [36,  2,  'Pernas',  '',  ['#0070C0', 20, 'B', 'i', 'u', 'Arial'] ,['#F8CBAD', 'left', 27, 364, [ 1, 2, 2, 2,]],  ['#F2F2F2', 11, 'B', 'i', 'u', 'Calibri'] ,['#C65911', 'center', 27, 74, [ 0, 1, 0, 1,]], 0, 0 ], 

                    ]
                    // ... fim de menu

// ------[MENUS]



