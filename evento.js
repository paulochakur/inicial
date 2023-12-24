/*
    REFERÊNCIAS
    *****
    https://www.educative.io/answers/how-to-dynamically-load-a-js-file-in-javascript

    *****
    https://www.w3schools.com/html/html5_webstorage.asp


    https://www.javatpoint.com/javascript-special-characters

    https://html.spec.whatwg.org/multipage/syntax.html

    https://www.digitalocean.com/community/tutorials/how-to-work-with-strings-in-javascript
    
    https://www.htmlsymbols.xyz/arrow-symbols

    Deep copy: https://code.tutsplus.com/articles/the-best-way-to-deep-copy-an-object-in-javascript--cms-39655
            clone = Object.assign({}, original)
            clone = JSON.parse(JSON.stringify(user))

            https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/assign

    Operator	Description
    ==	        Equal to: true if the operands are equal
    !=	        Not equal to: true if the operands are not equal
    ===	        Strict equal to: true if the operands are equal and of the same type
    !==	        Strict not equal to: true if the operands are equal but of different type or not equal at all


    // ******* Promise
    let loadImage = new Promise(function(resolve, reject) {
        el(imgId).onload = () => resolve(imgId)
    })
    loadImage.then(
        result => { formIm(imgId,  x, y, w, h, fitS, faEscX, faEscY, adjDiv) },
      )
    // *******[Promise]

*/


//  -------------- Variáveis Globais -------------
//var module = require('fs')
var mobFlag = 0 , scrTurn = 0, divSheetId = '' , divSheetAtiId = '', fatX = 1, fatY = 1
var ctx = 0, dictCoods = {}, verScr = 0
var linPrint = 0

// ... load arq Js
var nomeProjS = '', LoadedJS={}
var ListPlan = {name: [3]}  ,  nLoad = 0 , N = 0

// ... Array e Painel
var Painel = {'nome':{'toplin':1}} , Array = {'nome':{'toplin':1}}

// ... console
var eventoAc = '<br> .......'
var proprHab = 0 ; leftCon = 500    // ... flag de habilitação de janela de propriedades


// ... Zoom
var divOriId        = '' , divOri       = 0     // div da img original
var parentZoomId    = '' , parentZoom   = 0     // parent de divOri
var imgZooOriId     = '' , imgZooOri    = 0     // imagem original de divOr a ser amostrada

var divZooId        = '' , divZoo       = 0     // div display da amostra aumentada
var imgZoom         = 0                         // image aumentada, dentro de divZoo - el('ZoomImg')
var divCursor       = 0                         // div cursor                        - el("ZoomDivCur")

var fatZoom         = 1     // fator de aumento
var menuZoom        = 0     // flag de zoom ativo

var topAbs  = 0 , leftAbs = 0 , leftImO   = 0 , topImO   = 0 , zooParam  = 0 , cursLado = 0
var scaleOriX = 0 , scaleOriY = 0
// ...[Zoom]




// ... evento
var evento  = '', eventoA = '', lastIn = ''  ; var Leventos  = []

var eleOn = 'O'  ; var eleOnId  = 'O' ; var eleOnTy     = 1 ; var onStyle    = 1 ; var eleOnClass   = 1 ; var onPar = 1
var eleTa  = 1   ; var eleTaId  = 'T' ; var eleTaTy     = 1 ; var taStyle    = 1 ; var eleTaClass   = 1 ; var taPar = 1
var eleFo  = '@' ; var eleFoId  = 'F' ; var eleFoTy     = 1 ; var foStyle    = 1 ; var eleFoClass   = 1 ; var foPar = 1

var eleTaA = 1   ; var eleTaIdA = 'T' ; var eleTaTyA    = 1 ; var taStyleA   = 1 ; var eleTaClassA  = 1
var eleFoA = '@' ; var eleFoIdA = 'F' ; var eleFoTyA    = 1 ; var foStyleA   = 1 ; var eleFoClassA  = 1

var eleWget = 0 ; eleTget = 0; eleWrec = 0; eleTrec = 0; eleWoff = 0; eleToff = 0; eleWcli = 0; eleTcli = 0; eleWscr = 0; eleTscr = 0
var eleHget = 0 ; eleLget = 0; eleHrec = 0; eleLrec = 0; eleHoff = 0; eleLoff = 0; eleHcli = 0; eleLcli = 0; eleHscr = 0; eleLscr = 0

var foco = '' , entraInp = ''
var eleTaIdAnt = 0
var eleScrTop = 1

// ... coordenadas
var xMe  = 1    , yMe   = 1 , yMe   = 1 , yMe   = 1                             // relativas ao elemento
var xMw  = 1    , yMw   = 1 , Ih    = 1 , Iw    = 1                             // relativas a  window
var xMs  = 1    , yMs   = 1 , Sh    = 1 , Sw    = 1  , Ah    = 1 , Aw    = 1    // relativas a  screen
var xMp  = 1    , yMp   = 1 , Ph    = 1 , Pw    = 1                             // relativas a  page
var Dh   = 1    , Dw    = 1 , Oh    = 1 , Ow    = 1                             // relativas a  document
var Oh   = 1    , Ow    = 1
var xPs  = 1    , yPs   = 1                                                     // scroll de page

var lastX = 0, lastY = 0, delX = 0, delY = 0 ; divSheetId = '' ; divSheetAtiId = ''

var keyCode                         // keydown event
var ctrK                            // keydown event
// ...[evento]

// ... input
var habClic = 0 ; var iniValueInput = 0 , iniValueInputA = 0 , finValueInputA = 0
var toques=0
var inpCurId = '', inpAntId = '', inpCur = ''  , inpAnt = '', painelNome = '', celElA = ''
var iPaiCurr = 0 , jPaiCurr = 0 , cellCurr = 0

// ... abas
var paraAba = {} , iAbaAti = 1

// ... menu
var IsubOpen   = [1,1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]
var divMenuAti = '' ; var popMenuAti = ''
var iMenAnt = 0 , linMeIdAnt = ''   , scrollHab = 0

// ... planilha
var elAnt = '' , curAnt = '', downPla = 0
var iEl = 0     , jEl = 0           , divPlanCurr = 0   , divPlanCurrId = ''
var iSheetSel = 0     , jSheetSel = 0 , iSheetSelA = 0 , jSheetSelA = 0
var blocTrap = 0

// ************ TESTES

/*
chave1 = "nome"  ;     chave2 = "nLin"  ; chave3 = "nCol"  ; chT = "nTot"

dictJs = {chave1:"Plan" , "nLin":30 , chave3:60}
dictJs.nomePlan = "planilha"
dictJs[chT] = 456

dictMt["primeira"] = dictJs
*/

//  --------------[Variáveis Globais] -------------


//  --------------Listener de Eventos -------------
Leventos = ["abort", "afterprint", "animationend", "animationiteration", "animationstart", "beforeprint", "beforeunload", "blur", "canplay", "canplaythrough", "change", "click", "contextmenu", "copy", "cut", "dblclick", "drag", "dragend", "dragenter", "dragleave", "dragover", "dragstart", "drop", "durationchange", "ended", "error", "focus", "focusin", "focusout", "fullscreenchange", "fullscreenerror", "hashchange", "input", "invalid", "keydown", "keypress", "keyup", "load", "loadeddata", "loadedmetadata", "loadstart", "message", "mousedown", "mouseenter", "mouseleave", "mousemove", "mouseover", "mouseout", "mouseup", "mousewheel", "offline", "online", "open", "pagehide", "pageshow", "paste", "pause", "play", "playing", "popstate", "progress", "ratechange", "resize", "reset", "scroll", "search", "seeked", "seeking", "select", "show", "stalled", "storage", "submit", "suspend", "timeupdate", "toggle", "touchcancel", "touchend", "touchmove", "touchstart", "transitionend", "unload", "volumechange", "waiting", "wheel"]

// ----- funções de addEventListener
document.addEventListener("DOMContentLoaded", evIni)        ;  function evIni(){ eventoA = evento; evento = 'inicio'    ; iniSys() }

document.addEventListener(Leventos[00], ev00) ;  function ev00() { eventoA = evento; evento = Leventos[00]; eventTrap() }
document.addEventListener(Leventos[01], ev01) ;  function ev01() { eventoA = evento; evento = Leventos[01]; eventTrap() }
document.addEventListener(Leventos[02], ev02) ;  function ev02() { eventoA = evento; evento = Leventos[02]; eventTrap() }
document.addEventListener(Leventos[03], ev03) ;  function ev03() { eventoA = evento; evento = Leventos[03]; eventTrap() }
document.addEventListener(Leventos[04], ev04) ;  function ev04() { eventoA = evento; evento = Leventos[04]; eventTrap() }
document.addEventListener(Leventos[05], ev05) ;  function ev05() { eventoA = evento; evento = Leventos[05]; eventTrap() }
document.addEventListener(Leventos[06], ev06) ;  function ev06() { eventoA = evento; evento = Leventos[06]; eventTrap() }
document.addEventListener(Leventos[07], ev07) ;  function ev07() { eventoA = evento; evento = Leventos[07]; eventTrap() }
document.addEventListener(Leventos[08], ev08) ;  function ev08() { eventoA = evento; evento = Leventos[08]; eventTrap() }
document.addEventListener(Leventos[09], ev09) ;  function ev09() { eventoA = evento; evento = Leventos[09]; eventTrap() }
document.addEventListener(Leventos[10], ev10) ;  function ev10() { eventoA = evento; evento = Leventos[10]; eventTrap() }
document.addEventListener(Leventos[11], ev11) ;  function ev11() { eventoA = evento; evento = Leventos[11]; eventTrap() }
document.addEventListener(Leventos[12], ev12) ;  function ev12() { eventoA = evento; evento = Leventos[12]; eventTrap() }
document.addEventListener(Leventos[13], ev13) ;  function ev13() { eventoA = evento; evento = Leventos[13]; eventTrap() }
document.addEventListener(Leventos[14], ev14) ;  function ev14() { eventoA = evento; evento = Leventos[14]; eventTrap() }
document.addEventListener(Leventos[15], ev15) ;  function ev15() { eventoA = evento; evento = Leventos[15]; eventTrap() }
document.addEventListener(Leventos[16], ev16) ;  function ev16() { eventoA = evento; evento = Leventos[16]; eventTrap() }
document.addEventListener(Leventos[17], ev17) ;  function ev17() { eventoA = evento; evento = Leventos[17]; eventTrap() }
document.addEventListener(Leventos[18], ev18) ;  function ev18() { eventoA = evento; evento = Leventos[18]; eventTrap() }
document.addEventListener(Leventos[19], ev19) ;  function ev19() { eventoA = evento; evento = Leventos[19]; eventTrap() }
document.addEventListener(Leventos[20], ev20) ;  function ev20() { eventoA = evento; evento = Leventos[20]; eventTrap() }
document.addEventListener(Leventos[21], ev21) ;  function ev21() { eventoA = evento; evento = Leventos[21]; eventTrap() }
document.addEventListener(Leventos[22], ev22) ;  function ev22() { eventoA = evento; evento = Leventos[22]; eventTrap() }
document.addEventListener(Leventos[23], ev23) ;  function ev23() { eventoA = evento; evento = Leventos[23]; eventTrap() }
document.addEventListener(Leventos[24], ev24) ;  function ev24() { eventoA = evento; evento = Leventos[24]; eventTrap() }
document.addEventListener(Leventos[25], ev25) ;  function ev25() { eventoA = evento; evento = Leventos[25]; eventTrap() }
document.addEventListener(Leventos[26], ev26) ;  function ev26() { eventoA = evento; evento = Leventos[26]; eventTrap() }
document.addEventListener(Leventos[27], ev27) ;  function ev27() { eventoA = evento; evento = Leventos[27]; eventTrap() }
document.addEventListener(Leventos[28], ev28) ;  function ev28() { eventoA = evento; evento = Leventos[28]; eventTrap() }
document.addEventListener(Leventos[29], ev29) ;  function ev29() { eventoA = evento; evento = Leventos[29]; eventTrap() }
document.addEventListener(Leventos[30], ev30) ;  function ev30() { eventoA = evento; evento = Leventos[30]; eventTrap() }
document.addEventListener(Leventos[31], ev31) ;  function ev31() { eventoA = evento; evento = Leventos[31]; eventTrap() }
document.addEventListener(Leventos[32], ev32) ;  function ev32() { eventoA = evento; evento = Leventos[32]; eventTrap() }
document.addEventListener(Leventos[33], ev33) ;  function ev33() { eventoA = evento; evento = Leventos[33]; eventTrap() }
document.addEventListener(Leventos[34], ev34) ;  function ev34() { eventoA = evento; evento = Leventos[34]; eventTrap() }
document.addEventListener(Leventos[35], ev35) ;  function ev35() { eventoA = evento; evento = Leventos[35]; eventTrap() }
document.addEventListener(Leventos[36], ev36) ;  function ev36() { eventoA = evento; evento = Leventos[36]; eventTrap() }
document.addEventListener(Leventos[37], ev37) ;  function ev37() { eventoA = evento; evento = Leventos[37]; eventTrap() }
document.addEventListener(Leventos[38], ev38) ;  function ev38() { eventoA = evento; evento = Leventos[38]; eventTrap() }
document.addEventListener(Leventos[39], ev39) ;  function ev39() { eventoA = evento; evento = Leventos[39]; eventTrap() }
document.addEventListener(Leventos[40], ev40) ;  function ev40() { eventoA = evento; evento = Leventos[40]; eventTrap() }
document.addEventListener(Leventos[41], ev41) ;  function ev41() { eventoA = evento; evento = Leventos[41]; eventTrap() }
document.addEventListener(Leventos[42], ev42) ;  function ev42() { eventoA = evento; evento = Leventos[42]; eventTrap() }
document.addEventListener(Leventos[43], ev43) ;  function ev43() { eventoA = evento; evento = Leventos[43]; eventTrap() }
document.addEventListener(Leventos[44], ev44) ;  function ev44() { eventoA = evento; evento = Leventos[44]; eventTrap() }
document.addEventListener(Leventos[45], ev45) ;  function ev45() { eventoA = evento; evento = Leventos[45]; eventTrap() }
document.addEventListener(Leventos[46], ev46) ;  function ev46() { eventoA = evento; evento = Leventos[46]; eventTrap() }
document.addEventListener(Leventos[47], ev47) ;  function ev47() { eventoA = evento; evento = Leventos[47]; eventTrap() }
document.addEventListener(Leventos[48], ev48) ;  function ev48() { eventoA = evento; evento = Leventos[48]; eventTrap() }
document.addEventListener(Leventos[49], ev49) ;  function ev49() { eventoA = evento; evento = Leventos[49]; eventTrap() }
document.addEventListener(Leventos[50], ev50) ;  function ev50() { eventoA = evento; evento = Leventos[50]; eventTrap() }
document.addEventListener(Leventos[51], ev51) ;  function ev51() { eventoA = evento; evento = Leventos[51]; eventTrap() }
document.addEventListener(Leventos[52], ev52) ;  function ev52() { eventoA = evento; evento = Leventos[52]; eventTrap() }
document.addEventListener(Leventos[53], ev53) ;  function ev53() { eventoA = evento; evento = Leventos[53]; eventTrap() }
document.addEventListener(Leventos[54], ev54) ;  function ev54() { eventoA = evento; evento = Leventos[54]; eventTrap() }
document.addEventListener(Leventos[55], ev55) ;  function ev55() { eventoA = evento; evento = Leventos[55]; eventTrap() }
document.addEventListener(Leventos[56], ev56) ;  function ev56() { eventoA = evento; evento = Leventos[56]; eventTrap() }
document.addEventListener(Leventos[57], ev57) ;  function ev57() { eventoA = evento; evento = Leventos[57]; eventTrap() }
document.addEventListener(Leventos[58], ev58) ;  function ev58() { eventoA = evento; evento = Leventos[58]; eventTrap() }
document.addEventListener(Leventos[59], ev59) ;  function ev59() { eventoA = evento; evento = Leventos[59]; eventTrap() }
document.addEventListener(Leventos[60], ev60) ;  function ev60() { eventoA = evento; evento = Leventos[60]; eventTrap() }
document.addEventListener(Leventos[61], ev61) ;  function ev61() { eventoA = evento; evento = Leventos[61]; eventTrap() }
document.addEventListener(Leventos[62], ev62) ;  function ev62() { eventoA = evento; evento = Leventos[62]; eventTrap() }
document.addEventListener(Leventos[63], ev63) ;  function ev63() { eventoA = evento; evento = Leventos[63]; eventTrap() }
document.addEventListener(Leventos[64], ev64) ;  function ev64() { eventoA = evento; evento = Leventos[64]; eventTrap() }
document.addEventListener(Leventos[65], ev65) ;  function ev65() { eventoA = evento; evento = Leventos[65]; eventTrap() }
document.addEventListener(Leventos[66], ev66) ;  function ev66() { eventoA = evento; evento = Leventos[66]; eventTrap() }
document.addEventListener(Leventos[67], ev67) ;  function ev67() { eventoA = evento; evento = Leventos[67]; eventTrap() }
document.addEventListener(Leventos[68], ev68) ;  function ev68() { eventoA = evento; evento = Leventos[68]; eventTrap() }
document.addEventListener(Leventos[69], ev69) ;  function ev69() { eventoA = evento; evento = Leventos[69]; eventTrap() }
document.addEventListener(Leventos[70], ev70) ;  function ev70() { eventoA = evento; evento = Leventos[70]; eventTrap() }
document.addEventListener(Leventos[71], ev71) ;  function ev71() { eventoA = evento; evento = Leventos[71]; eventTrap() }
document.addEventListener(Leventos[72], ev72) ;  function ev72() { eventoA = evento; evento = Leventos[72]; eventTrap() }
document.addEventListener(Leventos[73], ev73) ;  function ev73() { eventoA = evento; evento = Leventos[73]; eventTrap() }
document.addEventListener(Leventos[74], ev74) ;  function ev74() { eventoA = evento; evento = Leventos[74]; eventTrap() }
document.addEventListener(Leventos[75], ev75) ;  function ev75() { eventoA = evento; evento = Leventos[75]; eventTrap() }
document.addEventListener(Leventos[76], ev76) ;  function ev76() { eventoA = evento; evento = Leventos[76]; eventTrap() }
document.addEventListener(Leventos[77], ev77) ;  function ev77() { eventoA = evento; evento = Leventos[77]; eventTrap() }
document.addEventListener(Leventos[78], ev78) ;  function ev78() { eventoA = evento; evento = Leventos[78]; eventTrap() }
document.addEventListener(Leventos[79], ev79) ;  function ev79() { eventoA = evento; evento = Leventos[79]; eventTrap() }
document.addEventListener(Leventos[80], ev80) ;  function ev80() { eventoA = evento; evento = Leventos[80]; eventTrap() }
document.addEventListener(Leventos[81], ev81) ;  function ev81() { eventoA = evento; evento = Leventos[81]; eventTrap() }
document.addEventListener(Leventos[82], ev82) ;  function ev82() { eventoA = evento; evento = Leventos[82]; eventTrap() }
document.addEventListener(Leventos[83], ev83) ;  function ev83() { eventoA = evento; evento = Leventos[83]; eventTrap() }
document.addEventListener(Leventos[84], ev84) ;  function ev84() { eventoA = evento; evento = Leventos[84]; eventTrap() }

document.addEventListener('DOMMouseScroll', ev84) ; function ev85() { eventoA = evento; evento = 'DOMMouseScroll'; eventTrap() }
// -----[funções de addEventListener]
// [--------------Listener de Eventos -------------]


// ----- inicio Sys
function iniSys(){
    evento = ''

    // Window
    Ih  = window.innerHeight                    ; Iw  = window.innerWidth
    Oh  = window.outerHeight                    ; Ow  = window.outerWidth        
    Sh  = window.screen.height                  ; Sw  = window.screen.width                     // inclui task bar
    Ah  = window.screen.availHeight             ; Aw  = window.screen.availWidth                // exculi task bar
    Ph  = document.body.clientHeight            ; Pw  = document.body.clientWidth
    xPs = Math.round(window.pageXOffset)        ; yPs = Math.round(window.pageYOffset)          //  page
    Dh  = document.documentElement.clientHeight ; Dw  = document.documentElement.clientWidth    // relativo a  document  - não inclui barras de scroll

    if(Aw<1000){mobFlag = 1 }

    // ---------- cria elementos de sistema

    // ... cria divs de Prevent Scroll
    para = document.createElement("div")
    el('Fundo').appendChild(para)
    para.id     = "divPrevScr"
    el("divPrevScr").style = 'position: absolute; left: -10000px; top: 74.0px; width: 1162.0px; height: 300.0px; box-sizing: content-box; margin: 0px; padding: 0px; overflow: scroll; border-color: black; border-width: 3.0px; border-style: solid; z-index: 1000;   '    
    el('divPrevScr').style.pointerEvents="none"

    para = document.createElement("div")
    el('divPrevScr').appendChild(para)
    para.id     = "divPrevMov"
    el("divPrevMov").style = 'position: absolute; left: -2000.0px; top: -2000.0px; width: 10000.0px; height: 10000.0px; box-sizing: content-box; margin: 0px; padding: 0px; border-color: black; border-width: 0px; border-style: solid'    
    
    el("divPrevMov").style.top  = '-2000.0px'
    el("divPrevMov").style.left = '-2000.0px'
    el("divPrevMov").setAttribute('wheelPrev', true)

    el('divPrevScr').style.top  = '-10000px'
    el('divPrevScr').style.left = '-10000px'
    el("divPrevScr").setAttribute('wheelPrev', true)

    el('divPrevScr').scrollTop  = 50000
    el('divPrevScr').scrollLeft = 50000
    // ...[cria divs de Prevent Scroll]
        
    // ... cria div - "console"
    para = document.createElement("div")
    el('Corpo').appendChild(para)
    para.id     = "console"
    el("console").style = 'position: fixed; left: 10px; top: 10px; width: 330px; height: 590px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: text; border-color: black; border-radius: 1.0%; border-width: 1.0px; border-style: solid; background-color: #d1cb47; opacity: 1; overflow: auto; z-index: 101; white-space: pre-wrap; font-family: Courier, sans-serif; font-weight: bold; font-size: 12.0px; '    
    // ...[cria div - "console"]

    // ... cria div - "proprBox"
    para = document.createElement("div") ; el('Corpo').appendChild(para) ; para.id     = "proprBox"
    el("proprBox").style = 'position: fixed; left: 10px; top: 135px; width: 320px; height: 420px; box-sizing: content-box; margin: 0.0px; padding: 20px; cursor: text; border-color: black; border-radius: 1.0%; border-width: 1.0px; border-style: solid; background-color: #fae8af; opacity: 1; overflow: hidden; z-index: 100 '

    padY = 11
    Llabls = ['Elem On :', 'Tipo     EleOn :', 'Parent EleOn :', 'Class   EleOn :', 'Focus Ele :', 'Target Ele :', 'Evento :', 'Mouse Ele On  :', 'Mouse Window:', 'Mouse Screen :', 'Mouse Page    :', 'Scroll Page :', '', ]
    LidPar = ['eleon2', 'tipoDiv2', 'parentDiv2', 'classDiv2', 'focusDiv2', 'eleTarDiv2', 'eventoDiv2', 'mouseEleDiv2', 'mouseWinDiv2', 'mouseCscDiv2', 'mousePageDiv2', 'scrollPageDiv2', '', '',]
    for (i=0; i < 12; i++) {
        if (i==1)  { padY = padY+25 }
        if (i==4)  { padY = padY+5 }
        if (i==6)  { padY = padY+15 }
        if (i==11) { padY = padY+30 }
        topL = padY + i*29

        // . . . param
        strStyL = 'position: absolute; left: 8px; top: 180px; width: 120px; height: 24.0px; box-sizing: border-box; margin: 0px; margin-top: 3.0px; padding: 0px; border-color: black; border-width: 0px; border-style: solid; background-color: transparent; white-space: pre-wrap; '
        strFonL = 'font-family: Arial, sans-serif; font-size: 15.0px; text-align: right; color: #0000e6; '
        pddL    = 'padding-top: 5px;'

        strStyP = 'position: absolute; left: 130px; box-sizing: border-box; margin: 0px;  overflow: hidden; border-color: black; border-width: 1.0px; border-style: solid; text-indent: 10.0px; text-align: left;  white-space: pre-wrap; '        
        strFonP = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 12.0px; color: black; '
        bcP     = 'background-color: #c1c1ff;'
        pddP    = 'padding-top: 5px;'

        heiL = 24 ; widL = 200
        if (i==0)           { 
            heiL = 40  ; widL = 220 ; bcP     = 'background-color: ;'     ; pddP = 'padding-top: 8px'      ; strFonP = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 18.0px; '
            strFonL = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 20.0px; text-align: right;   color: rgb(0, 0, 230); '
        }
        
        if (i==4 || i==5)   { 
            widL = 220 ; bcP = 'background-color: #a4a4ff;'  ; strFonP = 'font-family: Arial, sans-serif; font-weight: ; font-size: 14.0px; color: black; '
        }
        
        if (i==6)           { 
            widL = 220 ; bcP = 'background-color: #b8efa0;'  ; strFonP = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 14.0px; color: black; '
            strFonL = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 20.0px; text-align: right; color: #00b900; '
            pddL    = 'padding-top: 0px;'
        }
        if (i>6 && i<11)    { 
            widL = 220 ; bcP = 'background-color: #b8efa0;'  ; strFonP = 'font-family: Arial, sans-serif; font-weight: ; font-size: 12.0px; color: black; '
            strFonL = 'font-family: Arial, sans-serif; font-weight: ; font-size: 15.0px; text-align: right; color: #00b900; '
        }
        if (i==11)          {  
            widL = 220 ; bcP = 'background-color: #ff0fff;'  ; strFonP = 'font-family: Arial, sans-serif; font-weight: ; font-size: 14.0px; color: black; '
            strFonL = 'font-family: Arial, sans-serif; font-weight: bold; font-size: 18.0px; text-align: right; color: #ff0fff; '
            pddL    = 'padding-top: 0px;'
        }
        
        strStyP = strStyP + strFonP + bcP + pddP
        strStyL = strStyL + strFonL  + pddL
        // . . .[param]

        // ... labels
        para = document.createElement("div")    ; el('proprBox').appendChild(para)
        para.style          = strStyL
        para.style.top      = topL+'px'
        para.style.height   = heiL+'px'
        para.style.width    = 120+'px'
        para.innerHTML      = Llabls[i]

        // ... param
        para = document.createElement("div")    ; el('proprBox').appendChild(para)        
        para.id             = LidPar[i]
        para.style          = strStyP
        para.style.top      = topL+'px'
        para.style.height   = heiL+'px'
        para.style.width    = widL+'px'
    }
    // ...[cria div - "proprBox"]
            
    // ... cria canvas - "atuJib"
    para = document.createElement("canvas")
    el('Corpo').appendChild(para)
    para.id     = "canTrans"
    el("canTrans").style = 'position: fixed; left: 0px; top: 0px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: text; border-color: black; border-radius: 0%; border-width: 0.0px; border-style: solid; background-color: transparent; opacity: 1; overflow: auto; z-index: 100'
    para.width  = 500 ; para.height = 1
    ctx  = el("canTrans").getContext("2d")

    // .  .  .  nome do Proj
    nomeProjS = el("Corpo").getAttribute('nomeProj')
    tamNom = nomeProjS.length - 1 ; if (tamNom>30) { tamNom = 30 }
    for (i=0; i<=tamNom+1 ; i++) {
        if (i==0)  { corS    = 'rgb('+61+','+58+','+tamNom+')'}
        if (i>0)   { carI = nomeProjS.substr(i-1, 1) ; asc = carI.charCodeAt() ; corS    = 'rgb('+asc+',10,0)'}
        ctx.fillStyle = corS ; ctx.fillRect(i+10, 0, 1, 1)
        // .....
    }
    // .  .  . [nome do Proj]

    // ...[cria canvas - "atuJib"]

    // ... cria div cursor - "ZoomDivCur"
    para = document.createElement("div")
    el('Corpo').appendChild(para)
    para.id = "ZoomDivCur"
    strSty  = 'position: absolute; left: -3000px; top: -3000px; width: 40.0px; height: 40px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: none; border-color: black; border-radius: 0.0%; border-width: 0.0px; border-style: solid; background-color: gray; opacity: 0.4; z-index: 40'
    el("ZoomDivCur").style          = strSty
    el("ZoomDivCur").style.overflow          = 'hidden'
    el("ZoomDivCur").style.opacity  = 0.4
    el("ZoomDivCur").style.zIndex   = 40
    // ...[cria div cursor - "ZoomDivCur"]

    // ... cria IMG de zoom  - "ZoomImg"
    para = document.createElement("img")
    para.id     = "ZoomImg"
    el('Corpo').appendChild( para )
    el("ZoomImg").style = 'position: absolute; left: -3000px; top: -3000px; width: 40.0px; height: 40px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: none; border-color: black; border-radius: 0.0%; border-width: 0.0px; border-style: solid; background-color: gray; opacity: 1'
    // ...[cria IMG de zoom  - "ZoomImg"]
    
    // ... cria div para abrigar img de zoom, se não existente
    para = document.createElement("div")
    el('Corpo').appendChild(para)       // div de zoom é irmã de divZoom
    para.id     = 'Zoom'
    el("Zoom").style = 'position: absolute; left: -3000px; top: -3000px; width: 150px; height: 150px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: none; border-color: black; border-radius: 0.0%; border-width: 0.0px; border-style: solid; background-color: gray; opacity: 1 ; overflow: hidden'
    // ........[Zoom]
    
    // ---------- [cria elementos de sistema]

    // . . . monta dictCoods - originais e aplica xl e yt e corrige overflow para mob
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    Ldivs = []  ; for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    // . . . "iw"
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        divId = Ldivs[iDiv]
        try{
            iw = el(divId).getAttribute("iw")
            if (iw>0){
                tDiv    = cssUnitToNr(window.getComputedStyle(el(divId)).top)
                lDiv    = cssUnitToNr(window.getComputedStyle(el(divId)).left)
                scrDiv  = el(divId).getAttribute("scroll")
                yt      = el(divId).getAttribute("yt")
                xl      = el(divId).getAttribute("xl")
                scY     = el(divId).getAttribute("scY")
                scX     = el(divId).getAttribute("scX")

                dictCoods[iw] = [tDiv, lDiv]
                // . . . aplica xl e yt - somente deskTop, em desenvovimento
                if (yt!=null  && mobFlag==0){ el(divId).style.top  = (tDiv-yt)+'px' }
                if (xl!=null  && mobFlag==0){ el(divId).style.left = (lDiv-xl)+'px' }
                if (scY!=null && mobFlag==0){ el(divId).scrollTop  = scY }
                if (scX!=null && mobFlag==0){ el(divId).scrollLeft = scX }

                // . . . somente mob
                if (scrDiv=='roll' && mobFlag==1){ el(divId).style.overflow = 'scroll' }
            }
        }catch{}
    }
    // . . .[monta dictCoods - originais]

    // ....... Inicial abas
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    Ldivs = []  ; for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    // . . . parâmetros de Aba 1
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        divId = Ldivs[iDiv]
        iAba = 0 ; eAba1 = divId.includes('-Aba-1')
        
        if (eAba1){ iA = divId.indexOf("-Aba-") ; iAba  = parseInt(divId.slice(iA+5)) ; divAbaId = divId.slice(0, iA)}
        if (iAba==1){
            wAba    = parseInt(window.getComputedStyle(el(divId)).width)
            hAba    = parseInt(window.getComputedStyle(el(divId)).height)
            tAba    = parseInt(window.getComputedStyle(el(divId)).top)
            lAba    = parseInt(window.getComputedStyle(el(divId)).left)
            paraAba[divAbaId] = [ tAba, lAba, wAba, hAba, 1]
        }
    }
    // . . .[parâmetros de Aba 1]

    // . . . parâmetros de Aba n
    nAbas = 1
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        divId = Ldivs[iDiv]
        iAba = 0 ; eAba = divId.includes('-Aba-')
        if (eAba){ iA = divId.indexOf("-Aba-") ; iAba  = parseInt(divId.slice(iA+5)) ; divAbaId = divId.slice(0, iA)}
        if (iAba>1){
            iw = el(divId).getAttribute("iw")
            el(divId).style.zIndex  = 0
            el(divId).style.top     = (paraAba[divAbaId][0])+'px' ; dictCoods[iw][0] = paraAba[divAbaId][0]
            el(divId).style.left    = (paraAba[divAbaId][1])+'px' ; dictCoods[iw][1] = paraAba[divAbaId][1]
            el(divId).style.width   = (paraAba[divAbaId][2])+'px'
            el(divId).style.height  = (paraAba[divAbaId][3])+'px'
            nAbas = nAbas + 1 ;     paraAba[divAbaId][4] = nAbas            
        }
    }
    // . . . ativa aba de cada div
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        divId = Ldivs[iDiv]
        iAba = 0 ; eAba = divId.includes('-Aba-')
        if (eAba){ iA = divId.indexOf("-Aba-") ; iAba  = parseInt(divId.slice(iA+5)) ; divAbaId = divId.slice(0, iA)}
        if(iAba==1){
            try{ abaAti    = el(divId).getAttribute("aba") }catch{ abaAti = 0 }
            if (abaAti>0) { showAba(divAbaId, iAba=abaAti)  }
        }
    }
    // . . .[ativa aba de cada div]
    
    // . . .[parâmetros de Aba n]

    // .......[Inicial abas]

    // ........ cria Array
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    Ldivs = []  ; for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        try  {
            divId = Ldivs[iDiv] ; divE = el(divId)
        
            nLar = parseInt(divE.getAttribute("nLar")) ; if (nLar==null){ nLar = 0 }
            nCar = parseInt(divE.getAttribute("nCar")) ; if (nCar==null){ nCar = 0 }
            hPar = parseInt(divE.getAttribute("hPar")) ; if (hPar==null){ hPar = 0 }
            vPar = parseInt(divE.getAttribute("vPar")) ; if (vPar==null){ vPar = 0 }
            liCo = divE.getAttribute("liCo")
            
            styles = window.getComputedStyle(divE)
            lOri    = parseInt(styles.left)   ; tOri    = parseInt(styles.top)
            wOri    = parseInt(styles.width)  ; hOri    = parseInt(styles.height)
            divE.setAttribute('lOri', lOri) ; divE.setAttribute('tOri', tOri) ; divE.setAttribute('wOri', wOri) ; divE.setAttribute('hOri', hOri)

            // . . . se div definida como array
            if ((nLar*nCar)>1){        
                divPainelId = divE.parentElement.id         
                // . . . preenche atributos de Array
                divE.className  = divId
                Array[divId]    = { 'nLar':nLar , 'nCar':nCar , 'hPar':hPar , 'vPar':vPar , 'topLin': 1 , 'lefCol': 1 ,
                                    'lOri':lOri , 'tOri':tOri , 'wOri':wOri , 'hOri':hOri , 'liCo': liCo,
                                    'parDivId': divE.parentElement.id
                                  }

                for (iArr = 0; iArr<=nLar-1; iArr++){
                    for (jArr = 0; jArr<=nCar-1; jArr++){
                        oriId = divId ; idApp = '_'+(iArr+1)+':'+(jArr+1)
                        
                        // ------ function clonaEle (oriId, idApp)
                        // . . . dissemina class por todos descendentes
                        Descen = el(oriId).querySelectorAll('*')
                        for (iCh = 0; iCh<=Descen.length-1; iCh++){ Descen[iCh].className = el(oriId).className}

                        clonaEle (oriId, idApp)
                        // ... redefine posição
                        elCel = el(divId+idApp)
                        elCel.style.left = (lOri+ (jArr-0)*(wOri+hPar))+'px'   
                        elCel.style.top  = (tOri+ (iArr-0)*(hOri+vPar))+'px'
                        // ------[function clonaEle (oriId, idApp)]
                    }
                }
                divE.remove()
            }

        }catch{}
    }
    // ........[cria Array]

    // ........ abre Planilhas
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    Ldivs = []  ; for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        try  {
            divSheetId = Ldivs[iDiv]
            auto        = Cells[divSheetId][0][0]['autoload']
            autoSize    = Cells[divSheetId][0][0]['autosize']
            nLinPla     = Cells[divSheetId][0][0]['nLinPla']
            nColPla     = Cells[divSheetId][0][0]['nColPla']
            arqDados    = Cells[divSheetId][0][0]['arqDados']
            sincrArq    = Cells[divSheetId][0][0]['sincr']
            hSheet      = el(divSheetId).offsetHeight
            wSheet      = el(divSheetId).offsetWidth
            if (auto=='autoload'){ 
                preencheSheet(lplanIni=0, cplanIni=0, divSheetId, prTotal=0)
                if (arqDados!='' && arqDados!=undefined) { loadDadosJS(divSheetId) }
            }
        }
        catch{}
    }
    // ........[abre Planilhas]
    
    // ........ Inclui TEXTOS
    Ldivs = []
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    for (iDiv = 0; iDiv <= nDivs-1; iDiv++){
        divSheetId = Ldivs[iDiv]
        textForm   = Texts[divSheetId]
        
        if (textForm!=undefined){ icluiTxtemDivWord(divSheetId) }
    }
    // ........[Inclui TEXTOS]

    // ........ Monta Menus
    for (me in Menus){
        formMenu    = Menus[me]
        auto        = formMenu[0][1]
        pop = me.includes("pop")
        if (auto=='autoload'){ criaMenu(me) }
    }
    // ........[Monta Menus]

    // ....
    // -------- finaliza rotina inicial

    
    // . . . scroll inicial
    deslX = el('Corpo').getAttribute("deslX") ; deslY = el('Corpo').getAttribute("deslY")
    window.scrollTo(deslX,deslY)

    iniLoc()
    atuJib()

    // ... situação inicial de Console
    if(hab==0)      { el('console').style.left = '-3000px'      ; el('proprBox').style.left = '-3000px'}
    if(hab==1)      { el('console').style.left = leftCon+'px'   ; el('proprBox').style.left = leftCon+'px'}
    if(proprHab==0) { el('console').style.zIndex = '101' }
    if(proprHab==1) { el('console').style.zIndex = '99' }

    // ...... ajusta mob
    hF = cssUnitToNr(window.getComputedStyle(el('Fundo')).height)
    wF = cssUnitToNr(window.getComputedStyle(el('Fundo')).width)
    
    fatX    =  Ih/wF ;    fatY = fatX
    delV    = (wF-hF)/2 + (hF/2)*(1-fatY)
    delH    = (hF-wF)/2 + (wF/2)*(1-fatX)

    if (mobFlag==1 && wF>400){
        el('Fundo').style.transform = "rotateZ(-90deg)"+" scaleX("+fatX+")"+" scaleY("+fatY+")"
        el('Fundo').style.left  = (-delV)+'px'
        el('Fundo').style.top   = (-delH)+'px'
        scrTurn = 1
    }
    // ......[ajusta mob]

    document.activeElement.blur()

    // ....
}
// -----[inicio Sys]

//  -------------- Trap de Eventos  -------------
function eventTrap() {
    // ---- Parâmetros de EVENTO
    // elementos de evento - On, Target, Foco
    eleTa  = event.target
    eleFo  = document.activeElement
    if (eleOn=='O')          { eleOn  = eleFo }
    if (evento=="mousemove") { eleOn  = eleTa }
    if(evento=='click' && eleTa!=eleOn) { eleTa = eleFo } // "click through"
    // . .

    // . . . anteriores
    foco=''
    if (eleFoA!=eleFo) { eleFoA = eleFo  ;   eleFoIdA = eleFoId  ;  eleFoClassA = eleFoClass  ;   eleFoTyA = eleFoTy ; foco='MUDOUfoco'} // ; print(' .....    MUDA FOCO evento:'+evento+' eleFo.tagName:'+eleFo.tagName) }
    if (eleTaA!=eleTa) { eleTaA = eleTa  ;   eleTaIdA = eleTaId  ;  eleTaClassA = eleTaClass  ;   eleTaTyA = eleTaTy }

    // On     Element
    eleOnId    = eleOn.id      ;   eleOnClass  = eleOn.className   ;   eleOnTy     = eleOn.tagName ;   onStyle = getComputedStyle(eleOn)    ;   onPar = eleOn.parentElement
    // Target Element
    eleTaId    = eleTa.id      ;   eleTaClass  = eleTa.className   ;   eleTaTy     = eleTa.tagName ;   taStyle = getComputedStyle(eleTa)    ;   taPar = eleTa.parentElement
    // Focus  Element
    eleFoId    = eleFo.id      ;   eleFoClass  = eleFo.className   ;   eleFoTy     = eleFo.tagName ;   foStyle = getComputedStyle(eleFo)    ;   foPar = eleFo.parentElement

    if (eleOnClass==undefined){ eleOnClass = 'Und'}
    if (eleTaClass==undefined){ eleTaClass = 'Und'}
    if (eleFoClass==undefined){ eleFoClass = 'Und'}

    // . . . ele Geom
    eleWget = cssUnitToNr(window.getComputedStyle(eleTa).width) ; eleHget = cssUnitToNr(window.getComputedStyle(eleTa).height) ; eleTget = cssUnitToNr(window.getComputedStyle(eleTa).top) ; eleLget = parseInt(window.getComputedStyle(eleTa).left)
    eleWrec = parseInt(eleTa.getBoundingClientRect().width)     ; eleHrec = parseInt(eleTa.getBoundingClientRect().height)     ; eleTrec = parseInt(eleTa.getBoundingClientRect().top)     ; eleLrec = parseInt(eleTa.getBoundingClientRect().left)
    eleWoff = eleTa.offsetWidth                                 ; eleHoff = eleTa.offsetHeight                                 ; eleToff = eleTa.offsetTop                                 ; eleLoff = eleTa.offsetLeft
    eleWcli = eleTa.clientWidth                                 ; eleHcli = eleTa.clientHeight                                 ; eleTcli = eleTa.clientTop                                 ; eleLcli = eleTa.clientLeft
    eleWscr = eleTa.scrollWidth                                 ; eleHscr = eleTa.scrollHeight

    // elementos de evento - On, Target, Foco
    // if (evento!="focusout"){ eleTa  = event.target} // ?????
    
    // .  .  .[tipo INPUT]

    // --- Mouse Coords 
    //  Window  -  Não atuliza se evento=="keyup"
    if (evento!="keyup" && evento!="keydown" && evento!="keypress") {
        
        // relativo a  window  - inclui scroll bars
        xMw = event.clientX ; yMw = event.clientY
        // relativo ao elemento On
        xMe = event.offsetX ; yMe = event.offsetY
        // relativo a  screen - inclui task bar
        xMs = event.screenX ; yMs = event.screenY
        // relativo a  page / body
        xMp = event.pageX   ; yMp = event.pageY

        // Movimento - X, Y e deltas
        if (evento=='touchmove' || evento=='touchstart')  {   
            xMw = parseInt(event.touches[0].clientX)    ; yMw = parseInt(event.touches[0].clientY)
            xMe = xMw - eleLrec                         ; yMe = yMw - eleTrec
            xMs = parseInt(event.touches[0].screenX)    ; yMs = parseInt(event.touches[0].screenY)
            xMp = parseInt(event.touches[0].pageX)      ; yMp = parseInt(event.touches[0].pageY)
            if(scrTurn==1){ [xMs, yMs] = [yMs, xMs] ; [xMp, yMp] = [yMp, xMp] }
        }
        if (evento=='touchmove' && lastX!=0)    { delX  = lastX - xMs  ; delY  = lastY - yMs }
        if (evento=='touchstart')               { lastX = xMs          ; lastY = yMs ; delX =0 ; delY = 0 }
        if (evento=='touchend')                 { lastX = 0            ; lastY = 0   ; delX =0 ; delY = 0}

        //if (evento=='touchmove' && lastX!=0)    { el('ConsAma-Txt').innerHTML = ' delY: '+delY}

        // Wheel
        if (evento=='wheel')      { delX  = event.deltaX                                 ; delY  = event.deltaY }
    }
    // ---[Mouse Coords]

    // Scroll 
    //  elemento
    eleOnScrTop = eleOn.scrollTop   ;   eleOnScrLeft = eleOn.scrollLeft
    eleTaScrTop = eleTa.scrollTop   ;   eleTaScrLeft = eleTa.scrollLeft
    eleFoScrTop = eleFo.scrollTop   ;   eleFoScrLeft = eleFo.scrollLeft
    //  page
    xPs = Math.round(window.pageXOffset)        ; yPs = Math.round(window.pageYOffset)
    
    // Keys
    keyCode = event.keyCode         ; ctrK = event.ctrlKey

    // . . . Painel
    painelNome = ''
    try{ 
        if( evento!='wheel' || evento!='click') { painelNome = eleTaClass }
        if( evento=='keydown')                  { painelNome = eleOnClass }
    }catch{}
    // . . .[Painel]

    // ..... atualiza Jiboia
    if(keyCode==88){ atuJib() }  // ctr X

    // .... inibe menu do browser em rightClick
    if(evento=='contextmenu'){ evento = 'rightclick' ; event.preventDefault() }

    // ..... último'focusin'
    if(evento=='focusin')                    { lastIn = eleTa }


    //if (evento=='dblclick')  { event.preventDefault() ; event.stopPropagation()}
    //event.stopPropagation()
    /*
    if(evento=='wheel'){ 
        cc = event.cancelable
        print('cc: '+cc)
        //event.preventDefault()     
    }
    */

        // ...
        
    // ----[Parâmetros de EVENTO]

    // **************************************

    // ---------- Scroll de página e panilha
    // ..... prevent default de scroll
    wheelPrevOn   = ( (onPar.getAttribute('wheelPrev')) || (eleOn.getAttribute('wheelPrev')) )
    if( !(wheelPrevOn) && scrollHab==1){
        scrollHab = 0
        el("divPrevMov").style.top  = '-2000.0px'
        el("divPrevMov").style.left = '-2000.0px'
    
        el('divPrevScr').style.top  = '-10000px'
        el('divPrevScr').style.left = '-10000px'
    }
    if(wheelPrevOn && scrollHab==0){
        scrollHab = 1   ;   divPrevScr = el('divPrevScr')
        if(onPar.getAttribute('wheelPrev')) { divSheet  = onPar }
        if(eleOn.getAttribute('wheelPrev')) { divSheet  = eleOn }
        divSheetId = divSheet.id ; parDiv = divSheet.parentElement
        
        // . . . geometria de divPrevScr
        parDiv.appendChild(divPrevScr)
        divPrevScr.style.top    = parseInt(window.getComputedStyle(divSheet).top)    +'px'
        divPrevScr.style.left   = parseInt(window.getComputedStyle(divSheet).left)   +'px'
        divPrevScr.style.width  = parseInt(window.getComputedStyle(divSheet).width)  +'px'
        divPrevScr.style.height = parseInt(window.getComputedStyle(divSheet).height) +'px'

        // ..... recoloca Prevent Scroll
        divPrevScr.scrollTop  = 5000 ; divPrevScr.scrollLeft = 5000
        divPrevScr.style.pointerEvents="auto"
    }
    // . . . reposiciona Prevent Scroll
    if(scrollHab==1){ divPrevScr.scrollTop  = 5000  ;   divPrevScr.scrollLeft = 5000 }

    if(eleOnId=="divPrevMov" && evento=='mouseup' && downPla==1) { 
        eleTa.focus()
        el('divPrevScr').style.pointerEvents="auto"
        downPla = 0
    }
    if(eleOnId=="divPrevMov" && evento=='mousedown' && eleTaId=="divPrevMov") {
        el('divPrevScr').style.pointerEvents="none"
        downPla = 1
    }    
    // .....[prevent default de scroll]
    //el('ConsAma-Txt').innerHTML = ' delY: '+delY
    if( (evento=='wheel' || evento=='touchmove') && eleFoClass.includes("-Pla") && scrollHab==1){
        nomeSheet   = eleFoClass
        nLinPla     = Cells[divSheetId][0][0]['nLinPla'] ; nColPla   = Cells[divSheetId][0][0]['nColPla']
        lplan0      = Cells[divSheetId][0][0]['lplan0']  ; cplan0    = Cells[divSheetId][0][0]['cplan0']
        lFrz        = Cells[divSheetId][0][0]['lFrz']    ; cFrz      = Cells[divSheetId][0][0]['cFrz']
        iElN        = Cells[divSheetId][0][0]['iElC']    ; jElN      = Cells[divSheetId][0][0]['jElC']
        
        // ..... deltas de wheel
        delLin = 0 ; delCol = 0
        if(scrTurn==1){ delX = -delX }

        if (delY!=0){ delLin = (Math.abs(delY)/delY) }
        if (delX!=0){ delCol = (Math.abs(delX)/delX) }
        if(evento=='touchmove' && Math.abs(delY)<20) { delLin = 0 }
        if(evento=='touchmove' && Math.abs(delX)<20) { delCol = 0 }
        //if(scrTurn==1){ [delLin, delCol] = [delCol, -delLin] }

        el('ConsAma-Txt').innerHTML = ' delY: '+delY+'::'+delX+' delLin:'+delLin+'  scrTurn:'+scrTurn

        // .....[deltas de wheel]
        lplan0A = lplan0          ; cplan0A = cplan0
        lplan0N = lplan0 + delLin ; cplan0N = cplan0 + delCol
        
        preencheSheet(lplanIni=lplan0N, cplanIni=cplan0N, divSheetId, prTotal=0)
        
        delL = Cells[divSheetId][0][0]['lplan0'] - lplan0A  ; delC = Cells[divSheetId][0][0]['cplan0'] - cplan0A
        iElM = Number(divSheet.getAttribute('iElM'))        ; jElM = Number(divSheet.getAttribute('jElM'))

        iElNfo  = iElN   + delL ; jElNfo  = jElN   + delC
        if(iElNfo<lFrz)   { iElNfo = lFrz}      ; if(jElNfo<cFrz)     { jElNfo = cFrz    }
        if(iElNfo>iElM){ iElNfo = iElM}   ; if(jElNfo>jElM)  { jElNfo = jElM }
        poeFoco(divSheetId, iElNfo, jElNfo)     //; el(divSheetId, iSheet, jSheet).style.zIndex = 3

        // ..... recoloca Prevent Scroll
        el('divPrevScr').scrollTop  = 5000
        el('divPrevScr').scrollLeft = 5000
    }
    // ----------[Scroll de página e panilha]

    // ---------- scroll de Painel
    if( (((evento=='wheel' || evento=='touchmove') || evento=='click') || (evento=='keydown' && (keyCode>36 && keyCode<41) ))){
        pai = 1 ; try{ divPaId  = Painel[painelNome]['divPainelId'] } catch{ pai = 0 }
        if(pai==1){
            rePrint = 0 ; roda = 0
            // setas - preventdefault
            if(evento=='keydown' && (keyCode>36 && keyCode<41)){ event.preventDefault() }

            // ... parâmetros de Painel
            divPaId  = Painel[painelNome]['divPainelId'] ; porLInha  = Painel[painelNome]['porLInha']
            altLin   = Painel[painelNome]['altLin']      ; larCol    = Painel[painelNome]['larCol']
            nLinPai  = Painel[painelNome]['nLinPai']     ; nColPai     = Painel[painelNome]['nColPai']
            nCelPai  = Painel[painelNome]['nCelPai']
            nlWind   = Painel[painelNome]['nlWind']      ; ncWind    = Painel[painelNome]['ncWind']
            nLar     = Painel[painelNome]['nLar']        ; nCar      = Painel[painelNome]['nCar']
            tOri     = Painel[painelNome]['tOri']        ; lOri      = Painel[painelNome]['lOri']
            
            // . . . atual de Painel
            iTop     = Painel[painelNome]['iTop']        ; jLef      = Painel[painelNome]['jLeft']
            cellCurr = Painel[painelNome]['cellCurr']    ; iPaiCurr  = Painel[painelNome]['iPaiCurr']    ; jPaiCurr  = Painel[painelNome]['jPaiCurr']
            iTDisp   = Painel[painelNome]['iTDisp']      ; jLDisp    = Painel[painelNome]['jLDisp']
            scrTdis  = Painel[painelNome]['scrTdis']     ; scrLdis  = Painel[painelNome]['scrLdis']
            
            scrTop   = el(divPaId).scrollTop             ; scrLef   = el(divPaId).scrollLeft

            // ------- calcula cell foco - iPaiCurr
            // . . . click - define cell foco
            kcelarray   = parseInt(eleTa.getAttribute('kcelarray'))
            iArray      = parseInt(eleTa.getAttribute('iArray'))    ; jArray     = parseInt(eleTa.getAttribute('jArray'))
            if( evento=='click')  { 
                iPaiCurr =  iArray ; jPaiCurr =  jArray
                rePrint = 0 ; roda = 0
            }
            
            // . . . flechas
            if(evento=='keydown'){
                if(keyCode==38){ iPaiCurr = iPaiCurr-1 }
                if(keyCode==40){ iPaiCurr = iPaiCurr+1 }
                if(keyCode==37){ jPaiCurr = jPaiCurr-1 }
                if(keyCode==39){ jPaiCurr = jPaiCurr+1 }

                rePrint = 1 ; roda = 1
            }
            // . . .[flechas]

            // . . . wheel
            if((evento=='wheel' || evento=='touchmove')){
                elTop = el(painelNome, iTDisp, jLDisp) ; elFim = el(painelNome, iTDisp+nlWind-1, jLDisp+ncWind-1)

                hWin = parseInt(window.getComputedStyle(el(divPaId)).height) ; wWin = parseInt(window.getComputedStyle(el(divPaId)).width)
                yTop = parseInt(window.getComputedStyle(elTop).top)          ; yfim = parseInt(window.getComputedStyle(elFim).top)  + altLin
                xTop = parseInt(window.getComputedStyle(elTop).left)         ; xfim = parseInt(window.getComputedStyle(elFim).left) + larCol

                if ( (scrTop-scrTdis) > (altLin*0.75) && delY>0 && porLInha==1) { iPaiCurr = iTDisp+nlWind   ; rePrint = 1 ; roda = 1 }
                if (-(scrTop-scrTdis) > (altLin*0.75) && delY<0 && porLInha==1) { iPaiCurr = iTDisp-1        ; rePrint = 1 ; roda = 1 }
                
                if ( (scrLef-scrLdis) > (larCol*0.75) && delX>0 && porLInha==0) { jPaiCurr = jLDisp+ncWind   ; rePrint = 1 ; roda = 1 }
                if (-(scrLef-scrLdis) > (larCol*0.75) && delX<0 && porLInha==0) { jPaiCurr = jLDisp-1        ; rePrint = 1 ; roda = 1 }
            }
            // . . .[wheel]

            // ...............
            if(rePrint==1){

                if(porLInha==1){
                    if(jPaiCurr<1)    { jPaiCurr = nCar ; iPaiCurr = iPaiCurr - 1 ; if(iPaiCurr<1) { jPaiCurr = 1 } }
                    if(jPaiCurr>nCar) { jPaiCurr = 1    ; iPaiCurr = iPaiCurr + 1 }

                    if(iPaiCurr<1) { iPaiCurr = 1 }     ; if(jPaiCurr>nColPai) { jPaiCurr = nColPai }
                    if(jPaiCurr<1) { jPaiCurr = 1 }

                    cellCurr = (iPaiCurr-1)*nCar + jPaiCurr
                    if(cellCurr>nCelPai){ cellCurr = nCelPai ;  iPaiCurr = Math.ceil(cellCurr/nCar)    ; jPaiCurr = cellCurr-(iPaiCurr-1)*nCar }
                }
                if(porLInha==0){
                    if(iPaiCurr<1)    { iPaiCurr = nLar ; jPaiCurr = jPaiCurr - 1 ; if(jPaiCurr<1) { iPaiCurr = 1 } }
                    if(iPaiCurr>nLar) { iPaiCurr = 1    ; jPaiCurr = jPaiCurr + 1 }

                    if(jPaiCurr<1) { jPaiCurr = 1 }     ; if(jPaiCurr>nColPai) { jPaiCurr = nColPai }
                    if(iPaiCurr<1) { iPaiCurr = 1 }

                    cellCurr = (jPaiCurr-1)*nLar + iPaiCurr
                    if(cellCurr>nCelPai){ cellCurr = nCelPai ;  jPaiCurr = Math.ceil(cellCurr/nLar)    ; iPaiCurr = cellCurr-(jPaiCurr-1)*nLar }
                }
                Painel[painelNome]['cellCurr'] = cellCurr    ; Painel[painelNome]['iPaiCurr'] = iPaiCurr   ; Painel[painelNome]['jPaiCurr'] = jPaiCurr
                preenchePainel( cellCurr=cellCurr, painelNome=painelNome, centra=roda, nCelPai=1) 
            }
            Painel[painelNome]['cellCurr'] = cellCurr    ; Painel[painelNome]['iPaiCurr'] = iPaiCurr   ; Painel[painelNome]['jPaiCurr'] = jPaiCurr
            
            // .... põe e tira cursor
            celEl = el(painelNome, iPaiCurr, jPaiCurr) ; celEl .style.outlineWidth = "1px"
            if(celElA!='' && celElA!=celEl)            { celElA.style.outlineWidth = "0px" }
            celElA = celEl    
        }

        // . . .
    }
    // ----------[scroll de Painel]


    // -------- CONSOLE
    //if (evento=='keydown') { print(' keyCode: '+keyCode) }
    // ... copia valor de input                 - ctr"insert" ctr"c" 
    if (evento=='keydown' && (keyCode==67 || keyCode==45) && ctrK && eleTaid!='console') {
        navigator.clipboard.writeText(eleTa.value)
    }

    // ... copia id                 - ctr"i"
    if (evento=='keydown' && (keyCode==73 || keyCode==89) && ctrK) {
        if (keyCode==73) { nome = eleOnId                   ; navigator.clipboard.writeText(nome) }
        if (keyCode==89) { nome = eleOn.parentElement.id    ; navigator.clipboard.writeText(nome) }

        // . . . CSS properties
        try{
            print('')
            print('...............')
            LproprShow = ['top', 'left', 'width','height', 'color', 'backgroundColor', 'textOverflow', 'fontStyle', 'fontSize', 'fontFamily', 'fontWeight', 'white-space', 'text-align', 'padding', 'padding-left']
            comp = getComputedStyle(el(nome))
            for (i=LproprShow.length-1; i>=0 ; i--) {
                kPropr  = LproprShow[i]
                pro     = comp[kPropr]
                po      = '                   '. slice(0, 15-kPropr.length) +  kPropr
                print(' - '+po+' : '+pro)
            }
            print(' CSS Properties: ')
            print('')
            // . . . HTML atrributes
            print('')
            attrs = el(nome).getAttributeNames()
            for (i=attrs.length-1; i>=0; i--) {
                nomeAtt = attrs[i] ; no = '                   '. slice(0, 20-nomeAtt.length) +  nomeAtt
                if (nomeAtt!='id' && nomeAtt!='style') { print(no+': '+el(nome).getAttribute(nomeAtt)) }
            }
            print(' HTML atrributes: ')
            print('')
            print(' Parent : '+ el(nome).parentElement.id)
            print(' Tipo   : '+ el(nome).tagName)
            print(' Id     : '+ nome)
            print('...............')
            print('')
        } catch{}
    }

    // ... toggle CONSOLE visível   - ctr"c"
    if (evento=='keydown' && keyCode==90 && ctrK) { 
        leftAtu = el('console').style.left
        habLoc=1
        if (leftAtu!='-3000px'  && habLoc==1){ el('console').style.left = '-3000px' ; el('proprBox').style.left = '-3000px' ;  habLoc=0 ; proprHab = 0}
        if (leftAtu=='-3000px'  && habLoc==1){ el('console').style.left = leftCon+'px'    ; el('proprBox').style.left = (leftCon+4)+'px'     ;  habLoc=0 ; proprHab = 1}
    }
    // ... toggle CONSOLE front/bottom com Propriedades   - click
    if (evento=='click' && (eleTaId=='proprBox' || eleOn.parentElement.id=='proprBox')) { 
        frontAtu = el('console').style.zIndex
        habLoc=1
        if (frontAtu<101  && habLoc==1){ el('console').style.zIndex = '101' ; habLoc=0; proprHab = 0}
        if (frontAtu==101 && habLoc==1){ el('console').style.zIndex = '99'  ; habLoc=0; proprHab = 1}
    }

    // . . . preenche quadro
    if (eleOnId!='O' && eleOnId!=undefined && proprHab==1) {
        onBg = onStyle.backgroundColor

        el("eleon2").innerHTML                  = eleOnId                
        el("eleon2").style.backgroundColor      = onBg
        el("tipoDiv2").innerHTML                = eleOnTy                        //+" "+eleOn.scrollTop
        el("parentDiv2").innerHTML              = eleOn.parentElement.id         //+" "+window.pageYOffset
        el("classDiv2").innerHTML               = eleOn.className                + ' : '+ eleOn.getAttribute('class')                   //+" "+window.screenY
        el("eleTarDiv2").innerHTML              = eleTaId
        el("focusDiv2").innerHTML               = eleFoId
        el("eventoDiv2").innerHTML              = evento
        el("mouseEleDiv2").innerHTML            = " Z:"+eleTa.style.zIndex //"xMe:"+xMe+" yMe:"+yMe
        el("mouseWinDiv2").innerHTML            =  "xMw:"+xMw+" yMw:"+yMw+" Ih:"+Ih+" Iw:"+Iw
        el("mouseCscDiv2").innerHTML            = "xMs:"+xMs+" yMs:"+yMs+" Sh:"+Sh+" Sw:"+Sw
        el("mousePageDiv2").innerHTML           = "xMp:"+xMp+" yMp:"+yMp+" Ph:"+Ph+" Pw:"+Pw

        el("scrollPageDiv2").innerHTML          = "xPs:"+xPs+" yPs:"+yPs+" Ah:"+Ah+" Aw:"+Aw
    }
    // . . .[preenche quadro]

    // --------[CONSOLE]

    
    // ----------- INPUT
    entraInp = ''
    if (eleTaTy=='INPUT' || eleTaTyA=='INPUT') {
        if ( (evento=="keyup" || evento=="keydown" || evento=='focusin' || evento=="focusout" || evento=="click") && (eleFo.type!='range' || eleFoA.type=='range') ) {    
            
            readO = eleFo.getAttribute('readonly')
            
            // ... apaga em primeiro toque
            if(evento=="keydown" && readO!='true' && (keyCode>40 || keyCode==8 || keyCode==46) && toques==0) { eleFo.value = '' ; print('######')}
            // ... escape
            if(evento=="keydown" && readO!='true' && keyCode==27) { eleFo.value = iniValueInput }

            // . . . Got Focus
            if (eleTaTy=='INPUT' && evento=='focusin' && inpCur!=eleTa)  { 
                inpCur = eleTa  ; inpCurId = eleTaId
                entraInp = 'Got'
                iniValueInput  = eleTa.value
                // . . . toques    
                toques = 0 ; habClic = 0 ; setTimeout( fhabClic, 300)
                SelectionChange()
            } 
            // . . .[Got Focus]
            
            // . . . Lost Focus            
            if (eleTaTy=='INPUT' && evento=="focusout") {
                inpAnt = inpCur ; inpAntId = inpCurId
                inpCur = ''  ; inpCurId = 'lost'
                
                entraInp = 'Lost'
                iniValueInputA  = iniValueInput
                finValueInputA  = inpAnt.value

                SelectionChange()
            }
            // . . .[Lost Focus]
            
            // . . . incrementa toques
            if( readO!='true' && ((evento=="keydown" && keyCode>40) || (evento=="click" && habClic==1))  )   { toques++ }
        
            // .....
        }
    }
    // -----------[INPUT]

    // --------- ZOOM    
    // ... sai de Zoom
    if((evento=='mousemove' || evento=='touchmove') && menuZoom==1){
        if(eleTa.parentElement!=divOri && eleTaId!="ZoomDivCur" && eleTa.parentElement.id!="ZoomDivCur"){
            if(divZooId=='Zoom'){ el(divZooId).style.top        = '-3000px'  ; el(divZooId).style.left      = '-3000px' }
            el('Corpo').appendChild( el('ZoomImg') )
        
            divCursor.style.top   = '-3000px'  ; divCursor.style.left  = '-3000px'
            imgZoom.style.top     = '-3000px'  ; imgZoom.style.left     = '-3000px'
            imgZoom.style.width   = '10px '    ; imgZoom.style.height   = '10px'
            imgZoom.style.zIndex  = '0'
            menuZoom = 0
        }
    }
    // ...[sai de Zoom]

    // ... inicia Zoom
    if((evento=='mousemove' || evento=='touchmove' || evento=='click') && eleTa.getAttribute('zoommenudiv')!=null && eleTa.getAttribute('zoommenudiv')!='' && eleTaTy=='IMG' && menuZoom==0){
        imgZooOri   = eleTa                     ;   imgZooOriId   = eleTaId
        divOri      = imgZooOri.parentElement   ;   divOriId      = divOri.id
        divCursor   = el("ZoomDivCur")
        imgZoom     = el('ZoomImg')
        // params de zoom em Img
        zooParam    = imgZooOri.getAttribute('zoommenudiv')
        cursLadoS   = imgZooOri.getAttribute("curlado")       ; cursLado = Number(cursLadoS)
        fatZoomS    = imgZooOri.getAttribute("zoomfat")       ; fatZoom  = Number(fatZoomS)

        divZooId = zooParam
        divCursor.style.opacity         = 0.4
        divCursor.style.borderRadius    = '0%'
        divCursor.style.borderWidth     = '0px' 
        divCursor.style.zIndex          = '60' 
        divOri.style.cursor             = 'none'

        if(zooParam=='local')   { divZooId = 'Zoom' }
        if(zooParam=='self')    { divZooId = divOriId }
        if(zooParam=='cursor')  { divZooId = 'ZoomDivCur' ; divCursor.style.opacity = 1 }
        if(zooParam=='lupa')    { 
            divZooId = 'ZoomDivCur' 
            divCursor.style.opacity      = 1
            divCursor.style.borderRadius = '50%'
            divCursor.style.borderWidth  = '2px' 
        }

        divCursor.style.width   = cursLado+'px'
        divCursor.style.height  = cursLado+'px'
        
        // posição absoluta de divOri em relação à Window
        leftAbs = divOri.getBoundingClientRect().left   ;   topAbs = divOri.getBoundingClientRect().top

        // scale de divOri (transform)
        scaleOriX = divOri.getBoundingClientRect().width  / divOri.offsetWidth
        scaleOriY = divOri.getBoundingClientRect().height / divOri.offsetHeight
        if(scrTurn==1){ 
            scaleOriX = divOri.getBoundingClientRect().width  / divOri.offsetHeight
            scaleOriY = divOri.getBoundingClientRect().height / divOri.offsetWidth
        }

        // posição relativa de Img em divOri
        leftImO = imgZooOri.offsetLeft                  ;   topImO = imgZooOri.offsetTop
        
        // ...
        divOri.appendChild(divCursor)
        imgZoom.src = imgZooOri.src
        el(divZooId).appendChild( imgZoom )
        imgZoom.style.zIndex = '59'

        // . . . posiciona div de zoom para 'local'
        if(divZooId=='Zoom') {
            divZooId = 'Zoom'
            parentZoomId = el(divOriId).parentElement.id
            el(parentZoomId).appendChild(el('Zoom'))       // div de zoom é irmã de divZoom
            el(divZooId).style.height           = '150px'
            el(divZooId).style.zIndex           = '2'
            el(divZooId).style.overflow         = 'hidden'

            tDivOri  = cssUnitToNr(window.getComputedStyle(el(divOriId)).top)
            lDivOri  = divOri.offsetLeft
            wZoom    = divOri.offsetWidth
            hZoom    = divOri.offsetHeight
            el(divZooId).style.top  = (tDivOri-wZoom)+'px'
            el(divZooId).style.left = (lDivOri-hZoom)+'px'
        }

        // . . . formata imagem de zoom
        hImgOri    = imgZooOri.offsetHeight 
        wImgOri    = imgZooOri.offsetWidth
        imgZoom.style.height  = (hImgOri*fatZoom)+'px'
        imgZoom.style.width   = (wImgOri*fatZoom)+'px'

        menuZoom = 1
    }
    // ...[inicia Zoom]

    // ... move cursor Zoom
    if((evento=='mousemove' || evento=='touchmove' || evento=='click') && menuZoom==1){
        // . . . posicão do mouse sobre div Ori
        xMdivZ = xMw - leftAbs      ; yMdivZ = yMw - topAbs
        if(scrTurn==1){ xMdivZ = divOri.getBoundingClientRect().height - (yMw - topAbs) ; yMdivZ = xMw - leftAbs }

        // . . . posicão do mouse sobre imagem original
        xMdivZ = xMdivZ/scaleOriX   ; yMdivZ = yMdivZ/scaleOriY // corrige pos com escala de Ori (transform)
        xMimg  = xMdivZ - leftImO   ; yMimg = yMdivZ - topImO
        
        // . . . posiciona cursor
        if (zooParam!='self2') { 
            xCur = xMdivZ - cursLado/2 ; yCur = yMdivZ - cursLado/2 
            divCursor.style.top  = yCur+'px'
            divCursor.style.left = xCur+'px'
        }
        
        // . . . posiciona Img em div de zoom (divOri, divLocal, divPort ou cursor)
        // . zomm em cursor
        if (zooParam=='cursor' || zooParam=='lupa'){
            imgZoom.style.top  = (-yMimg*fatZoom + cursLado/2)+'px'
            imgZoom.style.left = (-xMimg*fatZoom + cursLado/2)+'px'
        }
        // . zomm em self, local ou divPort
        if (zooParam!='cursor' && zooParam!='lupa'){
            imgZoom.style.top  = (-yMimg*fatZoom + yMdivZ)+'px'
            imgZoom.style.left = (-xMimg*fatZoom + xMdivZ)+'px'
        }
        // . . .[posiciona Img em div de zoom (divOri, divLocal, divPort ou cursor)]
    }
    // ...[move cursor Zoom]
    // ---------[ZOOM]


    // ---------------- Menus    
    // ....... abre e recolhe    
    // . . .  recolhe Menu Pop  e Horizontal com click
    if (divMenuAti!='') {
        if ( (evento=='click') && ! eleTaClass.includes(divMenuAti.id) ){ 
            showMenu('menuDiv', recolhe=1) 
        }
    }
    // . . . abre Pop com rightclick
    if(popMenuAti!='') {        
        if(evento=='rightclick'){
            menuDiv = popMenuAti.id
            WidtRel = Menus[menuDiv][0][23] ; LeftRel   = Menus[menuDiv][0][24]
            wTot    = Menus[menuDiv][0][18] ; hTot      = Menus[menuDiv][0][19]

            y = yMp ; x = xMp

            if (yMw+hTot>Dh){ y = Dh-hTot-yMw + y }
            if (xMw+wTot>Dw){ x = Dw-wTot-xMw + x }
            popMenuAti.style.top  = y+'px' ; popMenuAti.style.left  = x+'px'

            IsubOpen   = [1, 1, -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]
            Menus[menuDiv][0][20] = IsubOpen
            showMenu(menuDiv)
            divMenuAti = popMenuAti
        }
    }
    // .....[abre Pop com rightclick]

    // .......[abre e recolhe]

    // ---------------- eventos em Menu
    if (eleTaClass.includes("-Menu")){
        if ( (evento=="keydown" || evento=='mousemove' || evento=='click') ){        
            // ... identifica divMenu e linMenu
            linMen = eleTa  ; linMenId = eleTaId    
            linMenId  = linMenId.replace("D", "")   ; linMen  = el(linMenId)
            linMenIdD = linMenId+'D'                ; linMenD = el(linMenIdD)
            inMe =  linMenId.indexOf("-Menu") ; abrePar = linMenId.indexOf("(") ; fechaPar  = eleTaId.indexOf(")") ; virg = linMenId.indexOf(",")
            if (inMe>0)    { divMenuId  = linMenId.slice(0, inMe)  ; divMenu = el(divMenuId)}
            if (abrePar>0) { iLinMenu   = parseInt(linMenId.slice(abrePar+1, virg)) }
        
            // ... altera foco
            if(eleTaIdAnt!=linMenId || evento=='click') {
                eleTaIdAnt = linMenId

                // . . . índices de menu
                formMenu = Menus[divMenuId]         ; formD         = formMenu[iLinMenu][5]    ; bc         = formD[0]
                formT    = formMenu[iLinMenu][4]    ; cF            = formT[0]
                nivel    = formMenu[iLinMenu][1]    ; opLin         = formMenu[iLinMenu][0] 
                iSub     = formMenu[iLinMenu][9]    ; subF          = formMenu[iLinMenu][10]
                nLins    = formMenu[0][0]           ; orient        = Menus[divMenuId][0][2]   ; hoverClic  = formMenu[0][3]
                paiIsub  = formMenu[iLinMenu][11]
                iMenAnt  = Menus[divMenuId][0][21]  ; linMeIdAnt    = Menus[divMenuId][0][22]
                
                // .... lê cabeçalho inicial
                lin1        = el(divMenuId+ "-Menu(" +1+ ")")
                top1        = parseInt(lin1.style.top)           ; h1 = parseInt(lin1.style.height)
                topAcu      = top1 + h1

                // ..... põe highlight - retira da lin anterior
                if (iMenAnt>0 && iMenAnt!=iLinMenu){
                    bcA     = formMenu[iMenAnt][5][0] ;  cFa     = formMenu[iMenAnt][4][0]
                    el(linMeIdAnt).style.backgroundColor    = bcA
                    el(linMeIdAnt).style.color              = cFa   ; el(linMeIdAnt+'D').style.color              = cFa
                }
                if (iMenAnt!=iLinMenu) {
                    bcH     = formMenu[iLinMenu][7][0] ;  cFh     = formMenu[iLinMenu][6][0]
                    linMen.style.backgroundColor            = bcH   
                    linMen.style.color                      = cFh   ;   linMenD.style.color                      = cFh   
                    
                    iMenAnt = iLinMenu ; linMeIdAnt = linMenId 
                    Menus[divMenuId][0][21] = iMenAnt ; Menus[divMenuId][0][22] = linMeIdAnt
                }        
                // .....[põe highlight]
            
                // ..... Click - Abre SubMenu
                if ( (evento=='click' || (evento=='mousemove' && hoverClic=='hover'))) {            

                    // ... marca submenu Open / Closed
                    habShow = 0 ; IsubOpen = Menus[divMenuId][0][20]
                    if (orient=='vertical'   && subF>0) { IsubOpen[subF] = -IsubOpen[subF] ; habShow = 1 }            
                    if (orient=='horizontal' || orient=='pop' || orient=='aba') { 
                        habShow = 1
                        jaAber = IsubOpen[subF]
                        // ... fecha sub 
                        for (iS = 0; iS <= IsubOpen.length-1; iS++){
                            if (IsubOpen[iS]>iSub){ IsubOpen[iS] = -1 }
                        }
                        // ... abre sub
                        if (subF>0 && (jaAber<0 || hoverClic=='hover'))  { IsubOpen[subF] = subF }
                    }
                
                    if (habShow>0) { Menus[divMenuId][0][20] = IsubOpen ; showMenu(divMenuId) ; divMenuAti = divMenu }
                }
                // .....[Click - Abre SubMenu]
            }
            // ... altera foco

            // ...
        }
    }
    // ----------------[eventos em Menu]
    
    // ---------------- Movimento em Sheet
    if (eleTaClass.includes("-Pla")) {
        if (blocTrap==0 && (evento=="keydown" || (evento=="click" && Number(eleTa.getAttribute('iElN'))>0) || evento=='input')  ){

            scro = 0 ; keyCodeF = 100
            if (evento=="click") { keyCodeF = 1 }

            iElN    = Number(eleTa.getAttribute('iElN'))    ; jElN   = Number(eleTa.getAttribute('jElN'))
            elAnt   = eleTa

            // .... inibe scroll com flechas
            readO = eleFo.getAttribute('readonly')
            if(evento=='keydown' && (keyCode<41 && keyCode!=8 && keyCode!=32) || (readO=='true' && keyCode==32)){ event.preventDefault() }

            // .... parâmetros de Sheet
            nomeSheet   = eleFoClass ; divSheet      = eleFo.parentElement ; divSheetId = divSheet.id ; divSheetAtiId = divSheetId
            divPlanCurr = divSheet   ; divPlanCurrId = divSheetId

            nLinSh      = Number(divSheet.getAttribute('nLinSh'))   ; nColSh    = Number(divSheet.getAttribute('nColSh'))
            nLinInps    = Number(divSheet.getAttribute('nLinInps')) ; nColInps  = Number(divSheet.getAttribute('nColInps'))
            hWindow     = Number(divSheet.getAttribute('hWindow'))  ; wWindow   = Number(divSheet.getAttribute('wWindow'))

            // .... parâmetros de Planilha
            nLinPla     = Cells[divSheetId][0][0]['nLinPla']        ; nColPla   = Cells[divSheetId][0][0]['nColPla']
            lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
            lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
            iElC        = Cells[divSheetId][0][0]['iElC']           ; jElC       = Cells[divSheetId][0][0]['jElC']             // foco atual em Pla
            cursor      = Cells[divSheetId][0][0]['cursor']         ; enterMove = Cells[divSheetId][0][0]['enterMove']   
            scrollBars  = Cells[divSheetId][0][0]['scrollBars'] 
            header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']


            // ------ mudança de foco em Sheet

            // ------ linha entrada em cabeçalho
            if (evento=="keydown" && keyCode==13 && eleFoId.includes("-headerLinL")){
                linEntr = parseInt( el(nomeSheet+"-headerLinL").value ) + lFrz - 1
                if (linEntr>nLinPla) { linEntr=nLinPla }
                iElN = linEntr ; jElN = jElC ; keyCode = 0
            }
            // ------[linha entrada em cabeçalho]

            // ------ movimento em Planilha - Scroll Bar
            if (evento=="focusin" && eleTaId.includes("Scroll")){ keyCodeF = 2}
            if (scrollBars==1 && eleTaId.includes("Scroll") && evento=='input'){
                valBarV = el(nomeSheet+"-vertScroll") .value
                valBarH = el(nomeSheet+"-horizScroll").value
                if (eleTaId.includes("vert"))  { iElN = parseInt( (valBarV/100)*(nLinPla-0) + 0 ) ; jElN = jElC }
                if (eleTaId.includes("horiz")) { jElN = parseInt( (valBarH/100)*(nColPla-0) + 0 ) ; iElN = iElC } 
                if (iElN<2){ iElN = 1 }
                if (jElN<2){ jElN = 1 }
                keyCodeF = 1
            }        
            // ------[movimento em Planilha - Scroll Bar]

            // ------ movimento em Planilha - Tecla - NOVO FOCO
            if ( (evento=="keydown" && keyCode<41 && keyCode!=8  && keyCode!=32 ) || keyCodeF<5){
                // ------ movimento em Planilha - Tecla
                iElNf = iElN
                if (LinMod>0 && iElN>lFrz) { iElNf = lFrz }    
                mCols = Cells[divSheetId][iElNf][jElN]['mergedCols']
                mLins = Cells[divSheetId][iElNf][jElN]['mergedLins']

                if (keyCode==13 && enterMove=='')     { jElN = jElN+1                       }       // Enter Dir
                if (keyCode==13 && enterMove=='baixo'){ iElN = iElN+1                       }       // Enter Baixo
                if (keyCode==39 && toques==0)         { jElN = jElN+mCols                   }       // Dir
                if (keyCode==37 && toques==0)         { jElN = jElN-1                       }       // Esq
                if (keyCode==38)                      { iElN = iElN-1                       }       // Cima
                if (keyCode==40)                      { iElN = iElN+mLins                   }       // Baixo
                
                if (keyCode==36 && ctrK && toques==0) { iElN = lFrz     ; jElN = cFrz       }       // Home
                if (keyCode==35 && ctrK && toques==0) { iElN = nLinPla  ; jElN = nColPla    }       // End
                if (keyCode==39 && ctrK && toques==0) { jElN = nColPla                      }       // Dir
                if (keyCode==37 && ctrK && toques==0) { jElN = cFrz                         }       // Esq                    
                if (keyCode==38 && ctrK)              { iElN = lFrz                         }       // Cima
                if (keyCode==40 && ctrK)              { iElN = nLinPla                      }       // Baixo
                // ------[movimento em Planilha - Tecla]

                if (iElN==lFrz && iElC<lFrz)  { iElN = lplan0 }    // transição cabeçalho-corpo
                if (jElN==cFrz && jElC<cFrz)  { jElN = cplan0 }    // transição cabeçalho-corpo
                            
                // ----------------------------------------------------------------
                focoCell(divSheetId)   // iElN, jElN
                // ----------------------------------------------------------------

            }    
        }
        // ------[movimento em Planilha - Tecla - NOVO FOCO]
    
        // ......        
    }
    // ----------------[Movimento em Sheet]

    // ------------------
    // ... eventTrap()
    trapEvent() // executa traps específicos do projeto em .html
}    
//  --------------[Trap de Eventos] -------------

            
// ************************* FUNCTIONS DE SISTEMA EM .js

// ..... cssUnitToNr
function cssUnitToNr(cssUnS){
    unS = cssUnS.replace('px', '')
    unS = unS.replace('%' , '')

    numUnit = Number(unS)
    // ....
    return numUnit
}



// ..... Converte data
function DataConv(dataS, brAm=''){                       
    // Data brasileira ou americana para JDN
    //   Brasileira: brAm='' ; Americana: brAm='A'
    //   Julian Day Number  -  https://en.wikipedia.org/wiki/Julian_day
    //   JDN = 0 assigned to the day starting at noon on Monday, January 1, 4713 BC
    //   JDN = 367 × Y - (7 × (Y + 5001 + (M - 9)/7))/4 + (275 × M)/9 + D + 1729777
    // Fomratos de entrada:
    //   "02/02/1961" ; "02-02-1961" ; "02/fev/61"; "2 de fevereiro de 1961"; "fevereiro-1961"
    //   "february 02 of 1961"
    // Saída:
    //   DataConv = { 'JDN':nnnnn , 'dd':d , 'mm':m , 'yy':y}
    DataDict = {}
    msPerDay = 1000 * 60 * 60 * 24
    Months   = ['', 'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

    dataSa = dataS
    dataS  = dataS.toLowerCase() ; dataS = dataS.trim()
    
    dataS  = dataS.replaceAll( "dia", " ")
    dataS  = dataS.replaceAll( ",",   " ")
    dataS  = dataS.replaceAll( "-",   " ")
    dataS  = dataS.replaceAll( "/",   " ")
    dataS  = dataS.replaceAll( "de ", " ")
    dataS  = dataS.replace( "  ",  " ")
    dataS  = dataS.replace( "  ",  " ")
    dataS  = dataS.replace( "  ",  " ")
    dataS  = dataS.replace( "  ",  " ")
    dataS  = dataS.replace( "  ",  " ")

    // ------------
    dataS  = dataS.replace("janeiro",   "jan")
    dataS  = dataS.replace("fevereiro", "feb")
    dataS  = dataS.replace("março",     "mar")
    dataS  = dataS.replace("abril",     "apr")
    dataS  = dataS.replace("maio",      "may")
    dataS  = dataS.replace("junho",     "jun")
    dataS  = dataS.replace("julho",     "jul")
    dataS  = dataS.replace("agosto",    "aug")
    dataS  = dataS.replace("setembro",  "sep")
    dataS  = dataS.replace("outubro",   "oct")
    dataS  = dataS.replace("novembro",  "nov")
    dataS  = dataS.replace("dezembro",  "dec")

    dataS  = dataS.replace("fev",     "feb")
    dataS  = dataS.replace("abr",     "apr")
    dataS  = dataS.replace("mai",     "may")
    dataS  = dataS.replace("ago",     "aug")
    dataS  = dataS.replace("set",     "sep")
    dataS  = dataS.replace("out",     "oct")
    dataS  = dataS.replace("dez",     "dec")

    // ------------
        
    // . . . data bras. tipo "25 12 1968"
    inMes = dataS.indexOf(" ") ; fiMes  = dataS.indexOf(" ", inMes+1)
    // . . . data sem dia ("fevereiro-61")
    if(inMes>-1 && fiMes==-1){ dataS = "1 "+dataS ; inMes = dataS.indexOf(" ") ; fiMes  = dataS.indexOf(" ", inMes+1) }
    
    dia = 77 ; mes = 88 ; ano = 0
    if(inMes>0)                 { diaS = dataS.substring(0, inMes+1).trim()      }
    if(inMes>0 && fiMes>inMes)  { mesS = dataS.substring(inMes+1, fiMes).trim()  }
    if(fiMes>inMes)             { anoS = dataS.substring(fiMes+1).trim()         }

    diaN = parseInt(diaS)

    // . . . corrige ano de dois dígitos
    if(anoS.length<4){ 
        if(anoS.length==1)   { anoS = '0'+anoS}
        if(parseInt(anoS)>29){ anoS = "19"+anoS }
        if(parseInt(anoS)<30){ anoS = "20"+anoS }
    }
    anoN = parseInt(anoS)

    // . . . corrige mês
    iM = Months.indexOf(mesS)
    if(iM>0) { mesN = iM}
    if(iM<0) { mesN = parseInt(mesS) ; mesS = Months[mesN] }
    
    // . . .[data bras. tipo "25/12/1968"]
    
    JDN = 367 * anoN - parseInt(parseInt(7 * (anoN + 5001 + (mesN - 9)/7))/4) + parseInt((275 * mesN)/9) + diaN + 1729777

    DataDict['dd']  = diaN ; DataDict['mm']  = mesN ; DataDict['yy']  = anoN
    DataDict['mS']  = mesS ; DataDict['JDN'] = JDN

    //dataS  = 'Mon Jan 01 1968'
    //dateJS = new Date(dataS)

    // ....
    return DataDict 
}
// .....[Converte data]


// ..... scale div
function scaleDiv(divId='', fitS='', faEscX = '', faEscY = '', desl=''){
    divE = el(divId)
    if(faEscX==''){ faEscX = faEscY }
    if(faEscY==''){ faEscY = faEscX }

    // . . . fit
    if(fitS!=''){
        divPar = divE.parentElement
        wPar = cssUnitToNr(window.getComputedStyle(divPar).width) ; hPar = cssUnitToNr(window.getComputedStyle(divPar).height)
        wDiv = cssUnitToNr(window.getComputedStyle(divE)  .width) ; hDiv = cssUnitToNr(window.getComputedStyle(divE)  .height)
        
        if(fitS=='fitW'){ faEscX = wPar/wDiv ; faEscY = faEscX }
        if(fitS=='fitH'){ faEscY = hPar/hDiv ; faEscX = faEscY }
    }


    x0 = cssUnitToNr(window.getComputedStyle(divE).left)  ; y0 = cssUnitToNr(window.getComputedStyle(divE).top)
    w0 = cssUnitToNr(window.getComputedStyle(divE).width) ; h0 = cssUnitToNr(window.getComputedStyle(divE).height)

    // . . . aplica escs
    divE.style.transform = 'scaleX('+faEscX+') scaleY('+faEscY+')'
    xDesl = (w0/2)*(faEscX-1) ; yDesl = (h0/2)*(faEscY-1)
    if(desl!=''){ divE.style.left =  (x0 + xDesl)+'px' ; divE.style.top = (y0 + yDesl)+'px' }
}
// .....[scale div]

// ..... Fomata Image
function fomatImg(imgId='', imgSrc='', x='', y='', w='', h='', fitS='', faEscX = '', faEscY = '', adjDiv=''){
    eleImg = el(imgId)
    
    if(x=='')      { x         = eleImg.getAttribute('x')       ; if(x==null)       { x = '0'} }
    if(y=='')      { y         = eleImg.getAttribute('y')       ; if(y==null)       { y = '0'}  }        
    if(w=='')      { w         = eleImg.getAttribute('w')       ; if(w==null)       { w = '0'}  }        
    if(h=='')      { h         = eleImg.getAttribute('h')       ; if(h==null)       { h = '0'}  }        
    if(fitS=='')   { fitS      = eleImg.getAttribute('fitS')    ; if(fitS==null)    { fitS = 'fit'} }
    if(faEscX=='') { faEscX    = eleImg.getAttribute('faEscX')  ; if(faEscX==null)  { faEscX = 1}  }   
    if(faEscY=='') { faEscY    = eleImg.getAttribute('faEscY')  ; if(faEscY==null)  { faEscY = 1}  }   
    if(adjDiv=='') { adjDiv    = eleImg.getAttribute('adjDiv')  ; if(adjDiv==null)  { adjDiv = 0}  }   
    eleImg.setAttribute('x' , x)
    eleImg.setAttribute('y' , y)
    eleImg.setAttribute('w' , w)
    eleImg.setAttribute('h' , h)
    eleImg.setAttribute('fitS' , fitS)
    eleImg.setAttribute('faEscX' , faEscX)
    eleImg.setAttribute('faEscY' , faEscY)
    eleImg.setAttribute('adjDiv' , adjDiv)

    srsAtu = eleImg.getAttribute('src')
    // . . . carrega nova imagem
    if (imgSrc!='' && imgSrc!=srsAtu) { el(imgId).src = imgSrc }

    //eleImg.style.left  = '-10000px'  ; eleImg.style.top = '-10000px'

    // ------- listener
    eleImg.addEventListener("load", loaImg) ;  function loaImg() { formIm(imgId,  x, y, w, h, fitS, faEscX, faEscY, adjDiv) } 
// . . . 
}
// . . . formata e não carrega
function formIm(imgId,  x='0', y='0', w='0', h='0', fitS, faEscX, faEscY, adjDiv){
    // ***** depois de carregar Img
    
    if(el(imgId).complete==true){
        eleImg = el(imgId) ; divImg = eleImg.parentElement ; divImgId = divImg.id
        wDS = (window.getComputedStyle(divImg).width)
        hDS = (window.getComputedStyle(divImg).height)

        wD = Math.round(cssUnitToNr(wDS))
        hD = Math.round(cssUnitToNr(hDS))

        wDC = divImg.getBoundingClientRect().width
        hDC = divImg.getBoundingClientRect().height

        oriW = eleImg.naturalWidth ; oriH = eleImg.naturalHeight

        // ----- Tamanho de Img
        imgWn = oriW ; imgHn = oriH

        // . . . W e H (prioridade 2)
        if (w!='0' || h!='0'){         
            w0 = cssUnitToNr(w) ; h0 = cssUnitToNr(h)
            if (w.includes('%')){ w0 = (w0/100) * wD }
            if (h.includes('%')){ h0 = (h0/100) * hD }        
            if (w0!=0){ imgWn = w0 ; faEscX = 1 ; if (h0==0){ imgHn = (imgWn/oriW)*oriH ; faEscY = 1 } }
            if (h0!=0){ imgHn = h0 ; faEscY = 1 ; if (w0==0){ imgWn = (imgHn/oriH)*oriW ; faEscX = 1 } }
        }
        // . . . fit (prioridade 1)        
        if (fitS=='fitW' || fitS=='fitH' || fitS=='fit'){
            if (fitS=='fitW'){ faEscX = wD / oriW ; faEscY = faEscX      }
            if (fitS=='fitH'){ faEscY = hD / oriH ; faEscX = faEscY      }
            if (fitS=='fit') { faEscX = wD / oriW ; faEscY = hD / oriH   }    
        }       
        // . . . scales (prioridade 3)

        imgWn = Math.round(oriW*faEscX)   ; imgHn = Math.round(oriH*faEscY)
        eleImg.style.width  = imgWn+'px'  ; eleImg.style.height = imgHn+'px'
        // -----[Tamanho de Img]

        // . . . ajusta tamanho de container
        // document.getElementById('div_register').setAttribute("style","width:500px")
        if(adjDiv==1) { divImg.style.width  = imgWn+'px'  ; divImg.style.height = imgHn+'px'}

        // . . . ajusta posição
        left0 = cssUnitToNr(x) ; top0 = cssUnitToNr(y)
        if (x.includes('%')){ left0 = (left0/100) * oriW }
        if (y.includes('%')){ top0  = (top0/100)  * oriH }
        eleImg.style.left  = left0+'px'  ; eleImg.style.top = top0+'px'
        // . . .[ajusta posição]
    }
    // *****[depois de carregar Img]
}
// .....[Fomata Image]

// ..... Clona Elemento
function clonaEle(oriId, idApp){
    clonaUniE(oriId, idApp, cloneChild=0)

    // . . . clone Children
    //Chil = el(oriId).children
    Descen = el(oriId).querySelectorAll('*')
    nCh = Descen.length
    parCloneId = oriId+idApp
    for (iCh = 0; iCh<=nCh-1; iCh++){
        oriIdC = Descen[iCh].id
        clonaUniE(oriIdC, idApp, cloneChild=1)
    }
    // . . .[clone Children]
}
function clonaUniE(oriId, idApp, cloneChild){
    // ----- function clonaUniE(oriIdU, idApp, parCloneId)
    // ... clona attributes                
    eleOri = el(oriId)  ; parOri = eleOri.parentElement   ; parOriId = parOri.id    ;   tipoEl = eleOri.tagName
    styles = window.getComputedStyle(eleOri) ; Att = eleOri.attributes

    // . . . separa i e j de idApp prese
    // . . . índices de input focus em Sheet
    iArray = '' ; jArray = ''
    
    idAppI = '*'+idApp
    try{
        abrePar = idAppI.indexOf("_") ; virg  = idAppI.indexOf(":") ; fim = idAppI.length
        if (abrePar>0){ iArray  = idAppI.slice(abrePar+1, virg) ; jArray   = idAppI.slice(virg+1, fim) }
    }catch{}
           
    if(cloneChild==0){ parCloneId = parOriId }    
    if(cloneChild==1){ parCloneId = parOriId+ idApp}    
    
    para = document.createElement(tipoEl)
    para.id = oriId + idApp
    el(parCloneId).appendChild(para)

    // ... clona Attr
    for (i=0; i<Att.length; i++){ if(Att[i].name!='id'){ para.setAttribute(Att[i].name, Att[i].value) } }
    // ... clona CSS
    for (iSt in styles){ stN = styles[iSt] ; if (stN!=undefined && stN!=''){ para.style[iSt] = stN } }
    // -----[function clonaUniE(oriIdU, idApp, parCloneId)]

    // ... attrs de identificação
    if (iArray!=''){ 
        para.setAttribute('iArray' , iArray) ; para.setAttribute('jArray' , jArray)
        para.setAttribute('niArray', nLar)   ; para.setAttribute('njArray', nCar)
    }

}
// .....[Clona Elemento]


function atuJib(){
    // ... escreve para Jiboia
    // . . . 
    ctx.fillStyle = 'rgb(127,127,127)'  ; ctx.fillRect(0, 0, 1, 1)  // topJib

    x1 = parseInt(xPs/255) ; x2 = parseInt(xPs - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
    ctx.fillStyle = corC                ; ctx.fillRect(1, 0, 1, 1)  // xPs
    y1 = parseInt(yPs/255) ; y2 = parseInt(yPs - y1*255)    ; corC='rgb('+y1+','+y2+','+0+')'
    ctx.fillStyle = corC                ; ctx.fillRect(2, 0, 1, 1)  // yPs

    // ....... Inicial abas
    allEl = document.getElementsByTagName("div")  ;  nDivs = allEl.length
    Ldivs = []  ; for (i = 0; i <= nDivs-1; i++){ Ldivs.push(allEl[i].id) }
    
    // . . . desl de Scroll e parâmetros de Aba
    nD = 0 ; nS = 0 ; verScr = verScr + 1
    corC='rgb('+verScr+','+0+','+0+')' ; ctx.fillStyle = corC ; ctx.fillRect(49       , 0, 1, 1)  // verScr
    corC='rgb('+0     +','+0+','+0+')' ; ctx.fillStyle = corC ; ctx.fillRect(50       , 0, 1, 1)  // nD
    corC='rgb('+0     +','+0+','+0+')' ; ctx.fillStyle = corC ; ctx.fillRect(150      , 0, 1, 1)  // nS
    
    for (iD=0; iD<=nDivs-1; iD++){
        
        divId = Ldivs[iD] ; div = el(divId)
        try{ iw = div.getAttribute("iw") }catch{ iw = 0 }
        if (iw==null){ iw = 0 }
        if (iw>0) { 

            // . . . iWAba; iAbaAti
            if (divId.includes('-Aba-1')){ 
                divAbaId = divId.slice(0, iA)
                try{ naba = paraAba[divAbaId][4]  }catch{ naba = 0 }
                if(naba>0){
                    iwAba = el(divId).getAttribute('iw')
                    x1 = parseInt(iwAba/255) ; x2 = parseInt(iwAba - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(3, 0, 1, 1)  // iwAba
                    x1 = parseInt(iAbaAti)                                  ; corC='rgb('+x1+','+0+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(4, 0, 1, 1)  // aba ativa
                }
            }
            // . . .[iWAba; iAbaAti]

            // . . . deslocamentos de top e left de 50 a 111 ( 20 elementos )
            tOri = dictCoods[iw][0]     ; lOri = dictCoods[iw][1]
            novoX = parseInt(window.getComputedStyle(div).left) ; novoY = parseInt(window.getComputedStyle(div).top)
            delX = lOri - novoX ; delY = tOri - novoY
            if (delX!=0 || delY!=0){
                nD = nD + 1
                corC='rgb('+nD+','+0+','+0+')' ; ctx.fillStyle = corC ; ctx.fillRect(50       , 0, 1, 1)  // nD
                x1 = parseInt(iw/255)   ; x2 = parseInt(iw   - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nD*3+48  , 0, 1, 1)  // iw            
                x1 = parseInt(delX/255) ; x2 = parseInt(delX - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nD*3+48+1, 0, 1, 1)  // delX
                x1 = parseInt(delY/255) ; x2 = parseInt(delY - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nD*3+48+2, 0, 1, 1)  // delY
            }
            // . . .[deslocamentos de top e left de 50 a 111 ( 20 elementos )]

            // . . . deslocamentos de scroll de ele            
            scr = div.getAttribute("scroll")
            scrX = div.scrollLeft    ; scrY = div.scrollTop
            if (scrX!=0 || scrY!=0){
                nS = nS + 1
                corC='rgb('+nS+','+0+','+0+')' ; ctx.fillStyle = corC ; ctx.fillRect(150       , 0, 1, 1)  // nS
                x1 = parseInt(iw/255)   ; x2 = parseInt(iw   - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nS*3+148  , 0, 1, 1)  // iw            
                x1 = parseInt(scrX/255) ; x2 = parseInt(scrX - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nS*3+148+1, 0, 1, 1)  // delX
                x1 = parseInt(scrY/255) ; x2 = parseInt(scrY - x1*255)    ; corC='rgb('+x1+','+x2+','+0+')'
                    ctx.fillStyle = corC                ; ctx.fillRect(nS*3+148+2, 0, 1, 1)  // delY
            }
            // . . .[deslocamentos de scroll de ele]
            
        }
    }
    // . . .[desl de Scroll e parâmetros de Aba]
    
        // ....
        return
}
// ...[escreve para Jiboia]

// ---- print em console
function print(strPrint){ linPrint ++; eventoAc = '<br> '+(linPrint)+' '+strPrint +eventoAc ; el("console").innerHTML = eventoAc ; return }
// ---

// ---- Get El by Id
function el(idS, iArr=0, jArr=0){ 
    if(typeof(idS)=='object') { idS = idS.id }
    if(iArr!=0 && jArr!=0)    { idS = idS+'_'+iArr+':'+jArr }
    return document.getElementById(idS) 
}
// ---

// ---- substitui caracteres especiais de strO
function charSpeci(strO){
    iEsp = 0
    while (iEsp>-1){
        if (iEsp==0){ iEsp = -1 }
        iEsp = strO.indexOf("&#", iEsp+1) ; iVir = strO.indexOf(";" , iEsp+1)
        if (iVir-iEsp<7) { nCars = iVir-iEsp-2 }
        if (iEsp>-1){
            carOri      = strO.substr(iEsp+2, nCars)
            carSubCod   = parseInt( carOri ) 
            carSub      = String.fromCharCode(carSubCod)
            strO        = strO.replace('&#'+carOri+';', carSub)
        }
    }    
    // ....
    return strO
}
// ----[substitui caracteres especiais de strO]

// ---- mede width de texto em px
function lenPxTxt (text, fontFam, fontSize){
    // ... 
    paraP = document.createElement("div")
    el('Corpo').appendChild(paraP)

    paraP.style.position    = "absolute"
    paraP.style.width       = 'auto'
    paraP.style.padding     = '0px'
    paraP.style.top         = '-10000px'
    paraP.style.margin      = '0px'
    paraP.style.fontFamily  = fontFam
    paraP.style.fontSize    = fontSize+'px'
    paraP.innerHTML         = text

    wD = parseInt(window.getComputedStyle(paraP).width)
    cW = paraP.clientWidth
    paraP.remove()
    // ....
    return cW
}
// ----[mede width de texto em px]

// .... INPUT
function fhabClic(){ habClic = 1 ; eleFo.setSelectionRange( 0, 0)}   //forward, backward and None

// ... show Aba
function showAba(divId, iAba=1){ 
    //iA = divId.indexOf("-Aba-") ; iAba  = parseInt(divId.slice(iA+5)) ; divAbaId = divId.slice(0, iA)
    
    divM = divId+'-Aba-'+iAba
    el(divM).style.zIndex  = 10
    
    nAbas = paraAba[divId][4]
    for (iAb = 1; iAb <= nAbas; iAb++){
        divF = divId+'-Aba-'+iAb
        if (iAb!=iAba) { el(divF).style.zIndex  = 0 }
    }
    iAbaAti = iAba
    // ...
}
// ...[show Aba]

// ... show/vanish Menu Completo
function showMenu(divMenuId, recolhe=0){ 
    topAcu   = 0 ; retorna = 1

    // .... RECOLHE se houver ativo
    if(divMenuAti.id!=divMenuId && divMenuAti!='' && recolhe==0){
        divMenuIdR = divMenuId ; divMenuR = divMenu ; IsubOpenR = IsubOpen
        showMenu('menuDiv', recolhe=1)
        divMenuId = divMenuIdR ; divMenu = divMenuR ; IsubOpen = IsubOpenR ; divMenuAti = ''
        recolhe = 0
    }
    // ....[RECOLHE se houver ativo]

    if(recolhe==0) {
        formMenu = Menus[divMenuId] ; orient   = formMenu[0][2]   ; IsubOpen = formMenu[0][20] ; nLins    = formMenu[0][0]
        retorna = 0
    }
    // .... recolhe menu
    if(recolhe==1 && divMenuAti!='') {
        retorna   = 0
        divMenuId = divMenuAti.id
        formMenu  = Menus[divMenuId] ; orient   = formMenu[0][2]   ; IsubOpen = formMenu[0][20] ; nLins    = formMenu[0][0]
        if (orient=='horizontal' || orient=='aba')  { IsubOpen   = [1, 1, -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]}
        if (orient=='pop')                          { IsubOpen   = [1,-1, -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]}
        formMenu[0][20] = IsubOpen ; divMenuAti=''
        if (orient=='vertical')                     { retorna   = 1}
    }

    if(retorna==1){ return }

    // ... vanish Pop
    if (orient=='pop' && IsubOpen[1]==-1) { el(divMenuId).style.top = '-10000px'}
    
    // ... vanish/show linhas
    for (i = 1; i <= nLins; i++){
        lin    = el(divMenuId+ "-Menu(" +i+ ")")    ; linD = el(divMenuId+ "-Menu(" +i+ ")D")
        topAtu = parseInt(lin.style.top)            ; htu  = parseInt(lin.style.height)
        
        iSubSV  = formMenu[i][9]                                    // subMenu de linha
        iPai    = formMenu[i][11]     ; iSubP   = formMenu[iPai][9] // linha pai de linha   // submenu de pai de linha
        iAvo    = formMenu[iPai][11]  ; iSubV   = formMenu[iAvo][9] // linha avó de linha   // submenu de avó de linha                    
        subOpen = IsubOpen[iSubSV]      // subMenu de linha aberto
        paiOpen = IsubOpen[iSubP]       // subMenu de pai aberto
        avoOpen = IsubOpen[iSubV]       // subMenu de avó aberto
        showVan = 0
        if ( (subOpen>0 && paiOpen>0 && avoOpen>0)  ) { showVan =  1 }
        if ( (subOpen<0 || paiOpen<0 || avoOpen<0)  ) { showVan = -1 }
        
        if (orient=='horizontal' || orient=='pop' || orient=='aba') {
            // . . vanish
            if (showVan==-1 && topAtu>=0) { lin.style.top  = (topAtu-10000)+'px' ; linD.style.top  = (topAtu-10000)+'px'}
            // . . show
            if (showVan== 1 && topAtu<0)  { lin.style.top  = (topAtu+10000)+'px' ; linD.style.top  = (topAtu+10000)+'px'}
        }
        if (orient=='vertical') {
            // . . . show
            if (showVan==1)     { lin.style.top = topAcu+'px'; linD.style.top = topAcu+'px'  ; topAcu = topAcu + htu }
            // . . . vanish
            if (showVan==-1)    { lin.style.top = -10000+'px'; linD.style.top = -10000+'px' }
        }
    }
    // ...

}
// ...[show/vanish Menu Completo]

// ---- Cria Menu
function criaMenu(divMenuId) {
    // -------- MONTA MENU - (divMenuId)
    var iSubA    = [0,0,0,0,0,0,0,0,0,0,0,0,0]
    var paiIsub  = [0,0,0,0,0,0,0,0,0,0,0,0,0]
    var TopRel   = [0,0,0,0,0,0,0,0,0,0,0,0,0]
    var LeftRel  = [0,0,0,0,0,0,0,0,0,0,0,0,0]
    var WidtRel  = [0,0,0,0,0,0,0,0,0,0,0,0,0]
    var HeigRel  = [0,0,0,0,0,0,0,0,0,0,0,0,0]

    widthL1 = 250

    formMenu        = Menus[divMenuId]
    formGer         = formMenu[0]
    nLins = formGer[0] ; orient = formGer[2] ; autosize = formGer[5] ; fitDiv = formGer[4]
  
    // ... cria div para popMenu
    if (orient=='pop') {
        para = document.createElement("div")
        el('Corpo').appendChild(para);
        para.id     = divMenuId
        para.className  = divMenuId+'-Menu'
        para.style  = "position: absolute; box-sizing: border-box; border-width: 0px; border-color: black; outline-width: 0; background-color: red "
        popMenuAti      = para
    }
    // ...[cria div para popMenu]

    divMenu         = el(divMenuId)
    if (divMenu==null && orient!='pop' ) { return }

    // --------- prévio
    
    // --------------- cria subMenus
    nSubMenus = 0 ; subF = 0
    nivel = 0 ; nivelA = -1 ; iSub = -1 ; iSubLa = -1 ; topRel = 0 ; iSubN = 0 ; topAcu = 0 ; popMenu = 1
    formMenu[0][9]  = 0   // iSub da linha 0
    formMenu[0][11] = 0   // pai da linha 0

    
    if ((orient=='horizontal'|| orient=='aba') && fitDiv=='fitDiv') { formMenu[1][5][2] = el(divMenuId).offsetHeight }

    for (i = 1; i <= nLins; i++){
        nivel   = formMenu[i][1] ;   opLin  = formMenu[i][0] ;   opLinA  = formMenu[i-1][0]
        if(nivel==1){ popMenu = 0 }

        // . . . 
        texto  = formMenu[i][2]             ;   textoD = formMenu[i][3]
        formT           = formMenu[i][4]    ;   sF = formT[1]  ; nF = formT[5]
        formD           = formMenu[i][5]    ;   hl = formD[2]  ; wl = formD[3] ; brd = formD[4]

        text = texto ; fontFam = nF ; fontSize = sF
        wL = lenPxTxt (text, fontFam, fontSize)

        subF = 0

        // .... cria novo submenu
        if (nivel>nivelA) {  
            nSubMenus       = nSubMenus + 1
            iSubA[nivelA]   = iSub
            iSub            = nSubMenus
            subF            = iSub
            paiIsub[iSub]   = i-1            
            iSubN           = iSub
        }
        if (nivel<nivelA) {  iSub = iSubA[nivel] }

        formMenu[i][9]   = iSub                         // número do submenu a que pertence a linha
        if (i>1)        { formMenu[i-1][10] = subF }    // número do submenu filho da linha
        if (i==nLins)   { formMenu[i][10]   = 0    }    // número do submenu filho da última linha
        formMenu[i][11]  = paiIsub[iSub]                // número da linha pai do submenu a que pertence a linha

        // ...... Geometria
    
        // . . horizontal
        if (orient=='horizontal' || orient=='pop' || orient=='aba' ) {
            
            heighLi = hl
            if (nivel==1 && i>1) { heighLi = formMenu[1][16] }  // height de cabeçalho horizontal = primeiro
            formMenu[i][16]  = heighLi

            if (nivel==1) {TopRel[iSub] = 0}
            
            // . . novo subMenu
            deslHab = 0
            // . . . top inicial de cada subMenu
            if (iSub>1 && iSubN==iSub) {
                iSubN = 0 ; iLp = paiIsub[iSub]
                tP = formMenu[iLp][12] ; hP = formMenu[iLp][16]
                if (iLp==0)   { tP=0; hP=0 }
                if (tP<0)     { tP = tP+10000}
                if (nivel==2) { TopRel[iSub] = tP+hP }
                if (nivel==3) { TopRel[iSub] = tP+2  }
                if (nivel==4) { TopRel[iSub] = tP+2  }
                deslHab = -10000
            }
            // . .[novo subMenu]
            
            formMenu[i][12]  = TopRel[iSub] + deslHab       // topL            
            formMenu[i][15]  = 10                           // padL
            TopRel[iSub]     = TopRel[iSub] + heighLi
        
            // ... bordas
            BorderL = brd
            if ( (nivelA>=nivel && nivel>1)  || (nivel==2 && popMenu==0) )   { BorderL[0]  = 0 }
            if ( (nivel==1 && i>1) )                        { BorderL[3]  = 0 }
            formMenu[i][17]  = BorderL
        }

        // . . vertical
        if (orient=='vertical') {
            formMenu[i][12]  = topAcu
            formMenu[i][13]  = 0
            formMenu[i][14]  = widthL1
            formMenu[i][15]  = 10+nivel*20
            formMenu[i][16]  = hl      

            topAcu = topAcu + hl
        }

        HeigRel[iSub] = HeigRel[iSub] + hl

        // ......[Geometria]                
        
        // ...
        nivelA = nivel
    }
    formMenu[0][25] = HeigRel
    formMenu[0][26] = TopRel
    // ---------------[cria subMenus]

    // ............................ width e left em 'horizontal'
    if(orient=='horizontal' || orient=='pop' || orient=='aba'){

       // .... width máxima de cada Sub
        WidtRel = [0,0,0,0,0,0,0,0,0,0,0,0,0]
        iSub   = -1
        for (i = 1; i <= nLins; i++){
            iSubA    = iSub
            iSub     = formMenu[i][9]   ; wl         = formMenu[i][5][3]
            texto    = formMenu[i][2]   ; fontFam    = formMenu[i][4][5] ;  fontSiz    = formMenu[i][4][1]
            textoDir = formMenu[i][3]   ; fontFamDir = formMenu[i][6][5] ;  fontSizDir = formMenu[i][6][1]
            if(autosize=='autosize'){                                       // width máxima de cada sub
                lenPxT   = lenPxTxt (texto   , fontFam   , fontSiz   )
                lenPxD   = lenPxTxt (textoDir, fontFamDir, fontSizDir)
                widLin   = lenPxT+lenPxD + 50
                if (widLin>WidtRel[iSub])   { WidtRel[iSub] = widLin }
            }
            if(autosize==''){                                               // width cabeça de cada sub
                if(iSubA!=iSub)             { WidtRel[iSub] = wl }
            }
        }
        formMenu[0][23] = WidtRel
        // ....[width máxima de cada Sub]        

        // ...... calcula left e aplica left e width em cada linha
        // .  .  .  aplica width de sub em cada linha
        for (i = 1; i <= nLins; i++){ iSub    = formMenu[i][9] ; formMenu[i][14] = WidtRel[iSub] } 

        lAcu = 0
        for (i = 1; i <= nLins; i++){
            iSub    = formMenu[i][9]    ;   nivel   = formMenu[i][1]
            iLp = paiIsub[iSub]         ;   lP = formMenu[iLp][13]      ; wP = formMenu[iLp][14]
            if(iLp==0){ lP=0; wP=0 }
            if (nivel==1) { LeftRel[iSub] = lAcu    ; lAcu = lAcu + formMenu[i][14]  }
            if (nivel==2) { LeftRel[iSub] = lP       }
            if (nivel==3) { LeftRel[iSub] = lP+wP-3  }
            if (nivel==4) { LeftRel[iSub] = lP+wP-3  }
            formMenu[i][13]  = LeftRel[iSub]
        } 
        formMenu[0][24] = LeftRel
        // .... máximos de Div 'horizontal'
        wTot = lAcu
        hTot = formMenu[1][16]
        formMenu[0][18] = wTot
        formMenu[0][19] = hTot
    }
    // ......[calcula e aplica left em cada linha]
    // ............................[width e left em 'horizontal']

    // .... define IsubOpen
    IsubOpen   = [1,1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]
    if(popMenu==1){ IsubOpen[1] = -1 }      // .... define POP
    formMenu[0][20] = IsubOpen


    // .... máximos de Div 'pop'
    if(orient=='pop'){  
        hTot = 0 ; wTot = 0
        for (i = 1; i <= nLins; i++){
            topL     = formMenu[i][12]
            leftL    = formMenu[i][13]
            widthL   = formMenu[i][14]
            hl       = formMenu[i][16]
            if (topL+hl>hTot)      { hTot = topL+hl }
            if (leftL+widthL>wTot) { wTot = leftL+widthL }
        }
        formMenu[0][18] = wTot
        formMenu[0][19] = hTot
    }
    // ....[máximos de Div]

    // .... máximos de Div 'vertical'
    if(orient=='vertical'){  
        hTot = formMenu[nLins][12] + formMenu[nLins][16]
        wTot = formMenu[1][14]
        formMenu[0][18] = wTot
        formMenu[0][19] = hTot
    }
    // ....[máximos de Div]

    // ... formata Div de Menu
    wTot    = formMenu[0][18]
    hTot    = formMenu[0][19]

    //divMenu.style.height   = (hTot)+"px"
    //divMenu.style.width    = (wTot)+"px"
    divMenu.style.zIndex   = 90

    // ---------[prévio]

    // ---------------------------------
    // .... cria DIVs de Menu
    nivel = 0
    for (i = 1; i <= nLins; i++){
        opLin   = formMenu[i][0]
        nivel   = formMenu[i][1]
        texto   = formMenu[i][2]
        textoD  = formMenu[i][3]

        iSub     = formMenu[i][9]        // número do submenu a que pertence a linha
        subF     = formMenu[i][10]       // número do submenu filho  da linha
        iLinP    = formMenu[i][11]       // número do pai            da linha
        topL     = formMenu[i][12]
        leftL    = formMenu[i][13]
        widthL   = formMenu[i][14]
        padL     = formMenu[i][15]
        hl       = formMenu[i][16]
        brd      = formMenu[i][17]
        
        wTot        = formMenu[0][18]
        hTot        = formMenu[0][19]
        IsubOpen    = formMenu[0][20]
        iMenAnt     = formMenu[0][21]  
        linMeIdAnt  = formMenu[0][22]
        WidtRel     = formMenu[0][23]
        LeftRel     = formMenu[0][24]
        HeigRel     = formMenu[0][25]
        TopRel      = formMenu[0][26]

        // ... form lin
        formT   = formMenu[i][4]    ; cF = formT[0] ; sF     = formT[1] ; nF = formT[5]
        formD   = formMenu[i][5]    ; bc = formD[0] ; anchor = formD[1] 
        padT    = (hl/2-(sF*1.2)/2)

        // . . . formata divs de menu
        // .. Esq
        para = document.createElement("div")
        divMenu.appendChild(para)
        para.id         = divMenuId+ "-Menu(" +i+ ")"
        para.className  = divMenuId+'-Menu'
        para.innerHTML  = texto

        para.style  = "position: absolute; box-sizing: border-box; border-style: solid; border-color: black; outline-width: 0;"
        para.style.color                    = cF
        para.style.fontFamily               = nF
        para.style.fontSize                 = sF+'px'
        para.style.textAlign                = anchor

        para.style.top                      = (topL)+"px"
        para.style.left                     = (leftL)+"px"
        para.style.width                    = (widthL)+"px"
        para.style.height                   = (hl)+"px"

        para.style.zIndex                   = (90+nivel)
        para.style.padding                  = "0px"
        para.style.paddingTop               = padT+"px"
        para.style.paddingLeft              = padL+"px"
        para.style.margin                   = "0px"
        para.style.textIndent               = '0px'
        para.style.textDecorationThickness  = '1px'
        para.style.backgroundColor          = bc
        
        desBordas(para, brd)
        // . . menu tipo "aba"    border-bottom-left-radius
        if(orient=='aba'){ 
            para.style.borderTopLeftRadius      = (hl/1.5)+"px"
            para.style.borderTopRightRadius     = (hl/1.5)+"px"
        }


        // .. Dir
        paraD = document.createElement("div");
        divMenu.appendChild(paraD)
        paraD.id        = divMenuId+ "-Menu(" +i+ ")D"
        paraD.className = divMenuId+'-Menu'
        paraD.innerHTML = textoD

        rD = parseInt(window.getComputedStyle(para).right)
        paraD.style = "position: absolute; box-sizing: border-box; border-style: solid; border-color: black; outline-width: 0;"

        paraD.style.color                   = cF
        paraD.style.fontFamily              = nF
        paraD.style.fontSize                = sF+'px'
        paraD.style.textAlign               = 'right'

        paraD.style.top                     = (topL)+"px"
        paraD.style.right                   = (rD)+"px"
        paraD.style.width                   = 'auto'
        paraD.style.height                  = (hl)+"px"

        paraD.style.zIndex                  = (90+nivel)
        paraD.style.padding                 = "0px"
        paraD.style.paddingTop              = padT+"px"
        paraD.style.paddingRight            = "10px"

        paraD.style.borderWidth             = "0px"
        paraD.style.textIndent              = '0px'
        paraD.style.backgroundColor         = 'transparent'
        // ..

    }
    // ....[cria DIVs de Menu]

    showMenu(divMenuId)

    // --------[MONTA MENU]    
// ... 
}           
// ----[Cria Menu]

// ------ Inclui texto em div Word
function icluiTxtemDivWord(divSheetId) {
    divSheet = el(divSheetId)
    var yIpar = []

    // ... format geral 
    textForm    = Texts[divSheetId]
    formGeral   = textForm[0]
    nParas      = formGeral.nParas  ; nShapes = formGeral.nShapes
    escaTot     = formGeral.escDiv

    margEsq     = formGeral.margEsq     ; margDir       = formGeral.margDir
    margSup     = formGeral.margSup     ; margInf       = formGeral.margInf

    margIntSup  = formGeral.margIntS    ; margIntInf    = formGeral.margIntI
    margIntEsq  = formGeral.margIntE    ; margIntDir    = formGeral.margIntD
    
    espeBor     = formGeral.espBord ; corBor  = formGeral.corBord
    
    lPar0       = 0

    // ... cria div de base '-texto '
    divTextId       = divSheetId+"-texto"
    divText         = document.createElement("div")
    divSheet.appendChild(divText)
    divText.id              = divTextId
    divText.style.zIndex    = 30

    divText.style = 'position: absolute; top: 0px; left: 0px; width: 50px; height: auto; box-sizing: content-box; margin: 0.0px; padding: 0px; padding-left: 0.0px; cursor: text; border-color: red;  border-width: 0.0px; border-style: solid; background-color: light blue; opacity: 1; z-index: 32; font-family: Courier New, sans-serif; font-weight: bold; font-size: 10.0px; text-indent: 0.0px; text-shadow: 0.0px 0.0px; text-align: center; overflow: hidden; white-space: pre-wrap; '
    
    divText.style.marginTop     = margSup+'px'
    divText.style.marginLeft    = margEsq+'px'

    //divText.style.paddingTop    = margIntSup+'px'
    divText.style.paddingBottom = margIntInf+'px'

    divText.style.borderWidth = espeBor+'px'
    divText.style.borderColor = corBor
    // ...[cria div de base '-texto ']

    // ... cria divs de cima para cada parágrafo
    yAcu0 = margIntSup ; yAcu = yAcu0 ; iDpar = 1
    for(ipar=1; ipar<nParas+nShapes ; ipar++){
        // ... lê texto e formatação de parágrafo
        textPar = textForm[ipar*2]
        parVaz = 0  ;   if (textPar=='') { textPar = '.' ; parVaz = 1}

        // ------- cria div de parágrafo
        para = document.createElement("div")
        divText.appendChild(para)
        para.id     = "divCima"+ipar
        para.style = 'position: absolute; top: 0px; left: 0px; height: auto; width: 200px; box-sizing: border-box; margin: 0.0px; padding: 0px; padding-left: 0.0px; cursor: text; border-color: black;  border-width: 0.0px; border-style: solid; background-color: transparent; opacity: 1; z-index: 32; font-family: Courier New, sans-serif; font-weight: bold; font-size: 10.0px; text-indent: 0.0px; text-shadow: 0.0px 0.0px; text-align: center; overflow: hidden; white-space: pre-wrap; word-break: break-word; '
        para.style.zIndex = 35
        para.style.borderWidth = '0px'
        //para.style.lineHeight = '40px'
        
        
        // ... format Para
        formtPara = textForm[ipar*2-1]
        tipoPara  = formtPara.tipoPar
        alimPara  = formtPara.alignP            ;   para.style.textAlign    = alimPara
        foSiPara  = formtPara.parFoS            ;   para.style.fontSize     = foSiPara+'px'
        indePara  = formtPara.indentP           ;   //para.style.paddingLeft  = indePara+'px'
        spacPara  = formtPara.lineSpaceP        ;   para.style.lineHeight   = spacPara*1.1
        // -------[cria div de parágrafo]

        // ... inclui texto em div de parágrafo
        para.innerHTML = textPar
        hTx = para.clientHeight

        // . . . linha de parametrização
        if (ipar==1){   
            para.style.whiteSpace  = ''         ; para.style.wordBreak  = ''
            para.style.height  = '100px'
            para.style.width  = 'auto'
            para.style.paddingLeft     = (margIntEsq + indePara)+'px'
            para.style.paddingRight    = margIntDir+'px'

            wDivTex = para.clientWidth
            para.style.width    = wDivTex+'px'
            yAcu = yAcu0
            para.remove()
        }   
        
        if  (i>1){

            // ...... div de textBox
            if (tipoPara=='txbox' || tipoPara=='pict') {

                tRel    = formtPara.t     ; lRel  = formtPara.l ; wTx  = formtPara.w ; hTx  = formtPara.h 
                lay     = formtPara.lay   ;  bc = formtPara.bc  ; op = formtPara.op  ; cB = formtPara.cB
                wB      = formtPara.wB
                iParO   = formtPara.iPar ; yPar = yIpar[iParO]
                
                para.style.borderColor      = cB
                para.style.borderWidth      = wB+'px'
                para.style.paddingLeft     = (indePara)+'px'

                // . . . cria div de fundo de textBox (para permitir opacity)
                if (tipoPara=='txbox'){
                    parFu = document.createElement("div")
                    divText.appendChild(parFu)
                    parFu.id     = "divFundo"+ipar
                    parFu.style = 'position: absolute; top: 0px; left: 0px; height: auto; width: 200px; box-sizing: border-box; margin: 0.0px; padding: 0px; padding-left: 0.0px; cursor: text; border-color: black;  border-width: 0.0px; border-style: solid; background-color: tranparent; opacity: 1; z-index: 32; font-family: Courier New, sans-serif; font-weight: bold; font-size: 10.0px; text-indent: 0.0px; text-shadow: 0.0px 0.0px; text-align: center; overflow: visible; white-space: pre-wrap; word-break: break-all; '
                    parFu.style.zIndex = 33 // na frente de divText e atrás de para de texBox                
                    parFu.style.backgroundColor  = bc
                    parFu.style.opacity          = 1-op

                    parFu.style.top      = (yPar+tRel)+'px'
                    parFu.style.left     = (lPar0+lRel+margIntEsq)+'px'

                    parFu.style.width    = wTx+'px'
                    parFu.style.height   = hTx+'px'
                    // .... layOut                
                    if (lay==5) { para.style.zIndex   = 34 ; parFu.style.zIndex = 33 }  // caixa de texto atrás     de parágrafo base
                    if (lay==3) { para.style.zIndex   = 37 ; parFu.style.zIndex = 36 }  // caixa de texto na frente de parágrafo base
                }
                // . . .[cria div de fundo de textBox (para permitir opacity)]

                // . . . cria img dentro de shape
                fN = formtPara.fN
                if (tipoPara=='pict' && fN!=''){
                    fR ='C:/Users/Paulo/Documents/CHAKUR/TÉCNICOS/Cianotipia/luis vaz de camoes.jpg'
                    fN = charSpeci(fN)

                    parIm = document.createElement("img")
                    para.appendChild(parIm)
                    parIm.id        = "imgShape"+ipar
                    parIm.style = 'position: absolute; left: 0px; top: 0px;  box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: text; border-color: black; border-radius: 0.0%; border-width: 0.0px; border-style: solid; background-color: gray; opacity: 1'
                    parIm.src = fN
                    
                    parIm.style.width    = wTx+'px'
                    parIm.style.height   = hTx+'px'

                    // .... layOut                
                    if (lay==5) { parIm.style.zIndex   = 34 ; para.style.zIndex = 34}  // imagem atrás     de parágrafo base
                    if (lay==3) { parIm.style.zIndex   = 37 ; para.style.zIndex = 37}  // imagem na frente de parágrafo base
                }
                // . . .[cria img dentro de shape]

                // . . .
                para.style.width    = wTx+'px'
                para.style.height   = hTx+'px'
                para.style.top      = (yPar+tRel)+'px'
                para.style.left     = (lPar0+lRel+margIntEsq)+'px'

            }
            // ......[div de textBox]

            // ...... div de para
            if (tipoPara=='para') { 
                espP = formtPara.espPar

                // . . .
                para.style.width    = wDivTex+'px'
                para.style.paddingLeft     = (margIntEsq + indePara)+'px'
                para.style.paddingRight    = margIntDir+'px'

                hTx = para.clientHeight
                para.style.height   = hTx+'px'
                para.style.top      = (yAcu)+'px'
                para.style.left     = (lPar0)+'px'
                
                                    
                // ...
                yIpar[iDpar] = yAcu ; iDpar = iDpar+1
                cH = para.clientHeight ; yAcu = yAcu + hTx + espP + 0
            }
            // ......[div de para]
        }
        // ...[format Para]  
    
        if (parVaz==1 && tipoPara!='pict') {  para.innerHTML = '' }
    
    
    }
    // ...[cria divs de cima para cada parágrafo]
    
    // .... ajusta tamanhos de Text e Fundo
    divText.style.height = yAcu+'px'
    divText.style.width  = wDivTex+'px'

    // .... ajusta scale
    wOri = divText.getBoundingClientRect().width  ;  hOri = divText.getBoundingClientRect().height // textBox Externo
    wD   = divSheet.clientWidth                   ;  hD   = divSheet.clientHeight    // Div Int. desconta Scrolls

    escX = escaTot  ; escY = escaTot    
    if (escaTot=='fitw') { escX = (wD-margDir-margEsq)/(wOri) ; escY = escX}
    if (escaTot=='fith') { escY = (hD-margSup-margInf)/(hOri) ; escX = escY }
    if (escaTot=='fit')  { 
        escY = (hD-margSup-margInf)/(hOri) 
        // . . . aplica escY e mede novamente wD
        xDesl = (wOri/2)*(escX-1)                       ;   yDesl = (hOri/2)*(escY-1)
        divText.style.left      =  xDesl+'px'           ;   divText.style.top      =  yDesl+'px'    
        divText.style.transform = 'scaleX('+1+') scaleY('+escY+')'
        wD = divSheet.clientWidth

        escX = (wD-margDir-margEsq)/(wOri) 
    }

    // . . . aplica Scale
    xDesl = (wOri/2)*(escX-1)                       ;   yDesl = (hOri/2)*(escY-1)
    divText.style.left      =  xDesl+'px'           ;   divText.style.top      =  yDesl+'px'
    divText.style.transform = 'scaleX('+escX+') scaleY('+escY+')'
    // ....[ajusta tamanhos de Text e Fundo]

    // . . . ajusta div de fundo (deve ser filha única de div de texto)
    eleFc   = divSheet.firstElementChild  ;   tyFc    = eleFc.tagName
    if (tyFc=='DIV') {
        wD = divSheet.clientWidth           ;   hD = divSheet.clientHeight
        wT = margEsq+wOri*escX+margDir   ;   hT = margSup+hOri*escY+margInf
        wF = Math.max(wD, wT)               ;   hF = Math.max(hD, hT)
        eleFc.style.width   = wF+'px'       ;   eleFc.style.height = hF+'px'
        eleFc.style.left    = 0+'px'        ;   eleFc.style.top = 0+'px'
        eleFc.zIndex        = 29 // atrás da div base divText
    }
    // . . .[ajusta div de fundo (deve ser filha única de div de texto)]

    // . . . divs de margens
    if (margDir>0){
        marg = document.createElement("div")
        divSheet.appendChild(marg)
        marg.id     = "divMargS"+ipar
        marg.style = 'position: absolute; top: 0px; left: 0px; height: auto; width: 200px; box-sizing: border-box; margin: 0.0px; padding: 0px; padding-left: 0.0px; cursor: ; border-color: black;  border-width: 0.0px; border-style: solid; background-color: transparent; opacity: 1; z-index: 32; font-family: Courier New, sans-serif; font-weight: bold; font-size: 10.0px; text-indent: 0.0px; text-shadow: 0.0px 0.0px; '
        marg.style.zIndex = 35

        marg.style.width    = margDir+'px'
        marg.style.height   = margInf+'px'
        marg.style.top      = 0+'px'
        marg.style.left     = (margEsq+wOri*escX)+'px'
    }
    if (margInf>0){
        marg = document.createElement("div")
        divSheet.appendChild(marg)
        marg.id     = "divMargI"+ipar
        marg.style = 'position: absolute; top: 0px; left: 0px; height: auto; width: 200px; box-sizing: border-box; margin: 0.0px; padding: 0px; padding-left: 0.0px; cursor: ; border-color: black;  border-width: 0.0px; border-style: solid; background-color: transparent; opacity: 1; z-index: 32; font-family: Courier New, sans-serif; font-weight: bold; font-size: 10.0px; text-indent: 0.0px; text-shadow: 0.0px 0.0px; '
        marg.style.zIndex = 35

        marg.style.width    = margDir+'px'
        marg.style.height   = margInf+'px'
        marg.style.top      = (margSup+hOri*escY)+'px'
        marg.style.left     = 0+'px'
    }

    // ... 
}          
// ------[Inclui texto em div Word]

// ---- Calcula larguras de Scrolls
function largSroll(ele){
    // ********************* dimensões de ele
    bE  = parseInt(window.getComputedStyle(ele).borderLeftWidth)
    bD  = parseInt(window.getComputedStyle(ele).borderRightWidth)
    bT  = parseInt(window.getComputedStyle(ele).borderTopWidth)
    bB  = parseInt(window.getComputedStyle(ele).borderBottomWidth)
    
    wCss  = parseInt(window.getComputedStyle(ele).width)  // width  nominal
    hCss  = parseInt(window.getComputedStyle(ele).height) // height nominal

    hOver = ele.scrollHeight    //  includes padding, excludes borders, scrollbars and margins.
    wOver = ele.scrollWidth     //  includes what is not viewable (because of overflow).
                                //   This property will round the value to an integer. 

    wExt  = Math.round(ele.getBoundingClientRect().width)   //  width  externa total, excludes margins - taking into account the transformation
    hExt  = Math.round(ele.getBoundingClientRect().height)  //  height externa total, excludes margins - taking into account the transformation 
                                                    //  ele.getBoundingClientRect() = { x : 113.5 , y : 50 , width : 440 , height : 240 , top : 50 , right : 553.5 , bottom : 290 , left : 113.5 }
                                                    //  https://developer.mozilla.org/en-US/docs/Web/API/Element/getBoundingClientRect 
    
    wUtil = Math.round(ele.clientWidth)  // width  interna útil -includes padding, excludes borders, scrollbars and margins.
    hUtil = Math.round(ele.clientHeight) // height interna útil -includes padding, excludes borders, scrollbars and margins.
    
    // ... Sroll é retirado da área útil, nunca aumenta retângulo externo
    wScr = Math.round(wExt - wUtil - bE - bD)   // larg do scroll vert
    hScr = Math.round(hExt - hUtil - bT - bB)   // larg do scroll hori

    print('wCss: '+wCss+' wUtil: '+wUtil+' wExt: '+wExt+' wOver: '+wOver)
    print('hCss: '+hCss+' hUtil: '+hUtil+' hExt: '+hExt+' hOver: '+hOver)
    print(` scrVert: ${wScr}   scrHor: ${hScr}`)
    // *********************[dimensões de ele]

    // ............
    var Scrll = [0,0]
    Scrll[0] = wScr ; Scrll[1] = hScr
    return Scrll
}
// ----[Calcula larguras de Scrolls]

// ------ Carrega arquivo de dados
function loadDadosJS(divSheetId, arqDados=''){
    
    nomeProjS = el("Corpo").getAttribute('nomeProj')
    nLoad = nLoad + 1
    if (arqDados=='') { arqDados = Cells[divSheetId][0][0]['arqDados'] }
    
    // ........ Load arquivo de dados JS
    scriptEle1 = document.createElement("script")        ; el('Corpo').appendChild(scriptEle1)
    scriptEle1.setAttribute("type", "text/javascript")   ; scriptEle1.setAttribute("async", true)
    scriptEle1.setAttribute("src", nomeProjS+"/ArquvosJs/"+arqDados)
    LoadedJS[divSheetId] = 0

    if (nLoad==1) { scriptEle1.addEventListener("load", loa1) ;  function loa1() { divLoadId = divSheetId; loadListener(divLoadId) } }
    if (nLoad==2) { scriptEle1.addEventListener("load", loa2) ;  function loa2() { divLoadId = divSheetId; loadListener(divLoadId) } }
    scriptEle1.remove()
}
function loadListener(divLoadId){
    divLoad = el(divLoadId)
    listS = charSpeci( JSON.stringify(ListArq) )
    Clone = JSON.parse(listS)
    ListPlan[divLoadId] = Clone
    listS = '' ; Clone = [] ; ListArq = ''
    nLinPla = ListPlan[divLoadId][0][0]
    LoadedJS[divLoadId] = 1
    
    // .... somente para arquivos vinculados a Sheet - lplan0Max , nLinInps, nColInps
    try{
        Cells[divLoadId][0][0]['nLinPla'] = nLinPla
        Cells[divLoadId][0][1]['nLinPla'] = 1       ;  ListPlan[divLoadId][0][1] = 1 // flag de arquivo carregado

        nLinInps    = Number(divSheet.getAttribute('nLinInps'))  ; nColInps  = Number(divSheet.getAttribute('nColInps'))
        

        // . . . recalcula lplan0Max
        lFrz        = Cells[divLoadId][0][0]['lFrz']           ; hei       = Cells[divLoadId][lFrz][1]['height']
        hWindow     = Number(divLoad.getAttribute('hWindow'))  ; LinMod    = Cells[divLoadId][0][0]['LinMod']
        if(LinMod>0) { 
            lplan0Max = nLinPla - parseInt(hWindow/hei) + 1
            if(lplan0Max<lFrz){ lplan0Max = lFrz }
            divLoad.setAttribute('lplan0Max', lplan0Max)
        }
        // ...

        preencheSheet(lplanIni=0, cplanIni=0, divLoadId, prTotal=1)
    }catch{}
    // .... somente para arquivos vinculados a Sheet

    trapLoad(divLoadId)

    // ...
    }
// ------[Carrega arquivo de dados]

// ---- Cria Sheet baseado em dictFormPla e MatrCellsPla
function criaSheet(divSheetId){
    nLinPla     = Cells[divSheetId][0][0]['nLinPla']        ; nColPla   = Cells[divSheetId][0][0]['nColPla']
    lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    iElC        = Cells[divSheetId][0][0]['iElC']           ; jElC      = Cells[divSheetId][0][0]['jElC']      // foco atual em Pla
    cursor      = Cells[divSheetId][0][0]['cursor']         ; enterMove = Cells[divSheetId][0][0]['enterMove']
    header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']
    scrollBars  = Cells[divSheetId][0][0]['scrollBars']     ; autoSize  = Cells[divSheetId][0][0]['autosize']
    arqDados    = Cells[divSheetId][0][0]['arqDados']

    Cells[divSheetId][0][0]['lplan0']   = lFrz   ; Cells[divSheetId][0][0]['cplan0'] = cFrz
    Cells[divSheetId][0][0]['iElC']     = lFrz   ; Cells[divSheetId][0][0]['jElC']   = cFrz
    
    ListPlan[divSheetId] = [ [nLinPla] ]
    
    lplan0  = lFrz  ; cplan0    = cFrz
    
    // ... substitui caracteres especiais e Carrega ListPlan
    nLin = nLinPla
    if(LinMod>0) {nLin = lFrz}
    for (i = 1; i <= nLin; i++){
        linList = [""]
        for (j = 1; j <= nColPla; j++){
            val     = Cells[divSheetId][i][j]['valor']
            strO    = charSpeci(val)
            Cells[divSheetId][i][j]['valor']    = strO
            linList.push(strO)
        }
        ListPlan[divSheetId].push(linList)
    }
    // ... [substitui caracteres especiais e Carrega ListPlan]
   
    // ... Sheet de Inputs
    divSheet = el(divSheetId)   ;   nomeSheet = divSheetId+'-Pla'

    // . . . margens
    margT = 0 ; margL = 0 ; margD = 0 ; margB = 0
    if (scrollBars==1)      { margD = 18 ; margB = 18 }
    if (header=='header')   { margT = 20 ; margL = 0 }

    // ... ajusta parâmetros de div 
    el(divSheetId).style.boxSizing = 'content-box'
    el(divSheetId).style.overflow  = 'hidden'
    
    // ... parâmetros de Sheets, definidos em Jiboia
    // calculados
    yTbord    = parseInt(window.getComputedStyle(divSheet).borderTopWidth)  ; yBbord    = parseInt(window.getComputedStyle(divSheet).borderBottomWidth)
    xEbord    = parseInt(window.getComputedStyle(divSheet).borderLeftWidth) ; xDbEord   = parseInt(window.getComputedStyle(divSheet).borderRightWidth)
    hSheet    = divSheet.offsetHeight                                       ; wSheet    = divSheet.offsetWidth
    hUtil     = hSheet - yTbord - yBbord - margT                            ; wUtil     = wSheet - xEbord - xDbEord - margL       
    hWindow   = hUtil - margB - Cells[divSheetId][lFrz][1]['top'] - 1       ; wWindow   = wUtil - margD - Cells[divSheetId][lFrz][cFrz]['left'] - 1

    nLinSh  = 0 ; nColSh  = 0
    
    // ---------- calcula lplan0Max cplan0Max    
    lplan0Max = 0 ; cplan0Max = 0
    
    // ... cplan0Max
    wPint = 0
    for (j = nColPla; j >= cFrz; j--){
        wPint = wPint + Cells[divSheetId][lFrz][j]['width']
        if(wPint>wWindow){ cplan0Max = j + 1; j = 0}
    }
    if(cplan0Max<cFrz){ cplan0Max = cFrz }
    // ... lplan0Max
    if(LinMod==0){ 
        hPint = 0
        for (i = nLinPla; i >= lFrz; i--){
            hPint = hPint + Cells[divSheetId][i][1]['height']
            if(hPint>hWindow){ lplan0Max = i + 1; break}
        }
    }
    lFrz        = Cells[divSheetId][0][0]['lFrz'] ; hei = Cells[divSheetId][lFrz][1]['height']
    if(LinMod>0)   { lplan0Max = nLinPla - parseInt(hWindow/hei) + 1}
    if(lplan0Max<lFrz){ lplan0Max = lFrz }
    // ---------- [calcula lplan0Max cplan0Max]

    // ...... calcula nLinInps e nColInps
    nColInps = 0
    for (j = cFrz; j <= nColPla; j++){
        nC = 0 ; leftJ = Cells[divSheetId][lFrz][j]['left']
        for (jj = j+1; jj<= nColPla; jj++){
            wWin = Cells[divSheetId][lFrz][jj]['left'] - leftJ + Cells[divSheetId][lFrz][jj]['width']
            if(wWin>=wWindow){ break}
        }
        nC = jj-j+cFrz+1 ; if(nC>nColInps){ nColInps = nC }
    }
    
    nLinInps = 0
    if(LinMod==0){
        for (i = lFrz; i <= nLinPla; i++){
            nL = 0 ; topI = Cells[divSheetId][i][1]['top']
            for (ii = i+1; ii<= nLinPla; ii++){
                hWin = Cells[divSheetId][ii][1]['top'] - topI + Cells[divSheetId][ii][1]['height']
                if(hWin>=hWindow){ break}
            }
            nL = ii-i+lFrz+1 ; if(nL>nLinInps){ nLinInps = nL }
        }
    }
    if(LinMod>0){ nLinInps = parseInt(hWindow / Cells[divSheetId][lFrz][1]['height']) + lFrz }
    // ......[calcula nLinInps e nColInps]

    print('   nLinInps:'+nLinInps+' nColInps:'+nColInps)

    // guarda
    divSheet.setAttribute('nomeSheet', nomeSheet)
    divSheet.setAttribute('nLinSh', nLinSh)             ; divSheet.setAttribute('nColSh'   , nColSh)
    divSheet.setAttribute('nLinInps', nLinInps)         ; divSheet.setAttribute('nColInps'   , nColInps)
    divSheet.setAttribute('hWindow', hWindow)           ; divSheet.setAttribute('wWindow'   , wWindow)
    divSheet.setAttribute('margL', margL)               ; divSheet.setAttribute('margT'   , margT)
    divSheet.setAttribute('lplan0Max', lplan0Max)       ; divSheet.setAttribute('cplan0Max', cplan0Max)
    divSheet.setAttribute('wheelPrev', true)
    
    // . . . para leitura
    nLinSh      = Number(divSheet.getAttribute('nLinSh'))    ; nColSh    = Number(divSheet.getAttribute('nColSh'))
    nLinInps    = Number(divSheet.getAttribute('nLinInps'))  ; nColInps  = Number(divSheet.getAttribute('nColInps'))
    hWindow     = Number(divSheet.getAttribute('hWindow'))   ; wWindow   = Number(divSheet.getAttribute('wWindow'))
    margL       = Number(divSheet.getAttribute('margL'))     ; margT     = Number(divSheet.getAttribute('margT'))
    lplan0Max   = Number(divSheet.getAttribute('lplan0Max')) ; cplan0Max = Number(divSheet.getAttribute('cplan0Max'))
    wheelPrev   = Number(divSheet.getAttribute('wheelPrev'))
    
    // ... cria Inputs
    for (i = 1; i <= nLinInps; i++){
        for (j = 1; j <= nColInps; j++){
            para = document.createElement("INPUT")
            divSheet.appendChild(para);
            para.id         = divSheetId+'_'+i+':'+j
            para.className  = nomeSheet
            para.style      = "position: absolute; left:"+(j*100-90)+"px; top:"+(i*30+10)+"px; width: 99px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; outline-width: 0;"
            para.style.borderWidth              = "0px"                
            para.style.textDecorationThickness  = '1px'
            para.style.zIndex                   = 2
            para.setAttribute('spellcheck', false)
            para.setAttribute('iSheet', i) ; para.setAttribute('jSheet', j)
            para.setAttribute('divSheetId', divSheetId)
        }
    }
    // ...[cria Inputs]
    
    // ... cria header
    if (header=='header'){        
        para = document.createElement("DIV")
        divSheet.appendChild(para)
        para.id     = nomeSheet+"-fundoHeader"
        para.className  = nomeSheet
        para.style  = 'position: absolute; left: 0px; top: 0px; height: 20px; width:'+(wUtil)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: ridge; border-color: black; outline-width: 1;'
        para.style.borderWidth      = "0px" ;   para.style.borderBottomWidth      = "1px"   
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        // 
        para = document.createElement("INPUT") ; para.setAttribute('readonly', 'readonly') ; para.value = 'Lin:'
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinS"
        para.className  = nomeSheet
        para.type="text"
        para.style  = 'position: absolute; left: 1px; top: 0px; height: 19px; width: 25px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"   
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        
        para = document.createElement("INPUT")
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinL"
        para.className  = nomeSheet
        para.type="number"
        para.style  = 'position: absolute; left: 31px; top: 0px; height: 19px; width: 100px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"  
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        para.value = 1
        
        para = document.createElement("INPUT") ; para.setAttribute('readonly', 'readonly')
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinN"
        para.className  = nomeSheet
        para.type="text"
        para.style  = 'position: absolute; left: 140px; top: 0px; height: 19px; width: 80px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"    ;   para.style.borderRightWidth      = "1px"    
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
    }
    // ...[cria header]

    
    // ... cria barras de scroll
    if (scrollBars==1){
        // . . . horizontal
        // fundo horizontal
        para = document.createElement("DIV");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-fundoHori"
        para.className  = nomeSheet
        para.style  = 'position: absolute; left: '+(margL)+'px; bottom: 0px; height: 17px; width:'+(wUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
        fundoHor = para
        // 
        para = document.createElement("INPUT");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-horizScroll"
        para.className  = nomeSheet
        para.type="range"
        para.style  = 'position: absolute; left: 0px; bottom: 1px; height: 17px; width:'+(wUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"   
        para.style.zIndex = '4'

        // . . . vertical
        // fundo vertical
        para = document.createElement("DIV");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-fundoVert"
        para.className  = nomeSheet
        para.style  = 'position: absolute; right: 0px; top:'+(margT)+'px; height: '+(hUtil-17)+'px; width: 17px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black ; '
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
        fundoVer = para
        // 
        para = document.createElement("INPUT");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-vertScroll"
        para.className  = nomeSheet
        para.type="range"
        para.style  = 'position: absolute; top:'+( (hUtil-17)/2 - 16/2 + margT) +'px; right:'+( -(hUtil-17)/2 + 16/2 )+'px; height: 16px; width:'+(hUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; transform: rotateZ(90deg);'
        para.style.borderWidth      = "2px"   
        para.style.zIndex = '4'

        // canto
        para = document.createElement("DIV");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-fundoCanto"
        para.className  = nomeSheet
        para.style  = 'position: absolute; bottom: 0px; right: 0px; height: 17px; width: 17px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; transform: rotateZ(90deg);'
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
    }
    // ...[cria barras de scroll]
    
    
    // ...
}           
// ----[Cria Sheet baseado em dictFormPla e MatrCellsPla]

// ----- Põe foco em Cell de Plan
function focoCell(divSheetId){    // iElN, jElN
    scro = 0
    lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    nLinPla     = Cells[divSheetId][0][0]['nLinPla']        ; nColPla   = Cells[divSheetId][0][0]['nColPla']
    lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    scrollBars  = Cells[divSheetId][0][0]['scrollBars'] 
    header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']
    cursor      = Cells[divSheetId][0][0]['cursor']

    iElC        = Cells[divSheetId][0][0]['iElC']           ; jElC      = Cells[divSheetId][0][0]['jElC']

    // . . . limites
    if (iElN<=1)       { iElN = 1       }
    if (jElN<=1)       { jElN = 1       }
    if (iElN>nLinPla)  { iElN = nLinPla }
    if (jElN>nColPla)  { jElN = nColPla }

    iElNf = iElN
    if (LinMod>0 && iElN>lFrz) { iElNf = lFrz }
    mergedCols          = Cells[divSheetId][iElNf][jElN]['mergedCols']
    mergedLins          = Cells[divSheetId][iElNf][jElN]['mergedLins']
    if (mergedCols<0) { jElN = jElN + mergedCols }
    if (mergedLins<0) { iElN = iElN + mergedLins }

    // ------------- analisa posição de novo foco - iElN , jElN   lplan0Max
    hei = Cells[divSheetId][lFrz][1]['height']
    if(iElN<lFrz)  {bF = 0}        
    if(iElN>=lFrz && LinMod==0)  { bF = Cells[divSheetId][iElN][jElN]['top'] - Cells[divSheetId][lplan0][jElN]['top'] + Cells[divSheetId][iElN][jElN]['height'] }
    if(iElN>=lFrz && LinMod>0)   { bF = (iElN - lplan0 + 1)*hei}

    xF = Cells[divSheetId][iElNf][jElN]['left'] - Cells[divSheetId][lFrz][cplan0]['left']    
    wF = Cells[divSheetId][iElNf][jElN]['width']
    dF = xF + wF
    lplanM = lplan0 + nLinSh-lFrz ; cplanM = lplan0 + nLinSh-cFrz
    // -------------[analisa posição de novo foco - iElN , jElN]

    // ---------- define lplan0 e cplan0
    // . . . estouro ACIMA
    if (iElN<lplan0 && iElN>=lFrz) { lplan0 = iElN  ;  Cells[divSheetId][0][0]['lplan0'] = lplan0 ;     scro = 1 }
    // . . . estouro ESQUERDA
    if (jElN<cplan0 && jElN>=cFrz) { cplan0 = jElN  ;  Cells[divSheetId][0][0]['cplan0'] = cplan0 ;     scro = 2 }
    // . . . estouro ABAIXO
    if (iElN>lplanM)         {                hPint = 0;                                                 scro = 3
        if(LinMod==0){
            for (i = iElN; i >= lFrz; i--){
                jj = cFrz
                if (i==iElN) { jj = jElN }    
                hPint = hPint + Cells[divSheetId][i][jj]['height']
                if(hPint>hWindow){ lplan0 = i + 1; break}
            }
        }
        if(LinMod>0){lplan0 = iElN - nLinSh + lFrz}
        Cells[divSheetId][0][0]['lplan0'] = lplan0
    }

    // . . . estouro DIREITA
    if (dF>wWindow)         {                wPint = 0;                                                 scro = 4
        for (j = jElN; j >= cFrz; j--){
            ii = lFrz
            if (j==jElN) { ii = iElNf }
            wPint = wPint + Cells[divSheetId][ii][j]['width']
            if(wPint>wWindow){ cplan0 = j + 1; break}
        }
        Cells[divSheetId][0][0]['cplan0'] = cplan0
    }
    // -----------[define lplan0 e cplan0]

    // ------------  Scroll de plan
    if (scro>0){ preencheSheet(lplanIni=0, cplanIni=0, divSheetId, prTotal=0) ; toques = 0 }
        
        // .......[SelectionChange de Scroll]
        //inpAnt = elAnt; inpAntId = inpAnt.id; inpCur = ''; inpCurId = 'lost';   finValueInputA = inpAnt.value
        //entraInp = 'Lost'; SelectionChange()
        
        // .... preenche de Scrooll Plan
        
        // ....[preenche de Scrooll Plan]

        //inpCur = prox; inpCurId = inpCur.id;                                    iniValueInput  = inpCur.value
        //entraInp = 'Got' ; SelectionChange()
    // ------------ [Scroll de plan]

    poeFoco(divSheetId, iElN, jElN)
    
    // ...
    return
}
// -----[Põe foco em Cell de Plan]

// ----- foco em próximo
function poeFoco(divSheetId, iElN, jElN){
    // ----- foco em próximo divSheetId, iElN, jElN
    lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']
    iElNf = iElN ; if (LinMod>0 && iElN>lFrz) { iElNf = lFrz }

    // ... prox
    iSheetP = iElN - lplan0 + lFrz       ;   jSheetP = jElN - cplan0 + cFrz
    if (iElN<lFrz){ iSheetP = iElN }     ;   if (jElN<cFrz){ jSheetP = jElN }
    prox = el(divSheetId, iSheetP, jSheetP)
    Cells[divSheetId][0][0]['iElC'] = iElN ; Cells[divSheetId][0][0]['jElC'] = jElN

    // .... põe e tira cursor de Pla
    // . . . retira Cursor anterior
    if (curAnt!=''){
        zAnt = 2
        if(iElN<lFrz || jElN<cFrz) { zAnt = zAnt + 2}
        mergedColsA  = Cells[divSheetId][iElNf][jElN]['mergedCols']
        mergedLinsA  = Cells[divSheetId][iElNf][jElN]['mergedLins']
        if(mergedColsA!=1 || mergedLinsA!=1){ zAnt = zAnt - 1 }

        el(curAnt).style.zIndex = zAnt
        if (cursor!='outline'){ desBordas(el(curAnt)) }
    }
    // . . .[retira Cursor anterior]

    // . . . põe cursor
    prox.style.zIndex = '5'
    if (cursor!='outline')  { prox.style.borderColor = cursor ; prox.style.borderWidth = '2px'; prox.style.outlineWidth      = "0px" }
    if (cursor=='outline')  { prox.style.outlineWidth      = "1px" }
    curAnt = prox
    // . . .[põe cursor]

    // ....[põe e tira cursor de Pla]

    // . . . foco em novo
    // .  .
    if(header=='header'){ 
        iLinH = iElN - lFrz + 1
        if (iLinH<0){ iLinH=0 }
        head = el(nomeSheet+"-headerLinL")   ; head.value = iLinH
        head = el(nomeSheet+"-headerLinN")   ; head.value = ' de '+(nLinPla-lFrz+1)
    }
    // .  .

    // .  . scrollBars de Plan
    if (scrollBars==1){
        el(nomeSheet+"-vertScroll").value  = parseInt( ((iElN-0)/(nLinPla-0) )*100 )
        el(nomeSheet+"-horizScroll").value = parseInt( ((jElN-0)/(nColPla-0) )*100 )
    }

    blocTrap = 1
    prox.focus()
    blocTrap = 0
    // .  .
    
    // . . .[foco em novo]

    // -----[foco em próximo]


    // ...
    return
}
// -----[foco em próximo]


// ---- Preenche Sheet a partir de Cells
function preencheSheet(lplanIni=0, cplanIni=0, divSheetId, prTotal=0){
    start = Date.now()
    //  lplanIni, cplanIni - da célula de Cells a aparecer no Top-Left Most
    divSheet = el(divSheetId)

    // ... cria Sheet em Div, se inexistente
    if (divSheet.getAttribute('nLinSh')==null){ criaSheet(divSheetId) }

    // . . . parâmetros de Planilha
    nLinPla     = Cells[divSheetId][0][0]['nLinPla']       ; nColPla   = Cells[divSheetId][0][0]['nColPla']
    lplan0      = Cells[divSheetId][0][0]['lplan0']        ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    lFrz        = Cells[divSheetId][0][0]['lFrz']          ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    iElC        = Cells[divSheetId][0][0]['iElC']          ; jElC      = Cells[divSheetId][0][0]['jElC']        // foco atual em Pla
    cursor      = Cells[divSheetId][0][0]['cursor']        ; enterMove = Cells[divSheetId][0][0]['enterMove']
    scrollBars  = Cells[divSheetId][0][0]['scrollBars'] 
    header      = Cells[divSheetId][0][0]['header']        ; LinMod    = Cells[divSheetId][0][0]['LinMod']
    autoSize    = Cells[divSheetId][0][0]['autosize']

    // ....  parâmetros de Sheet
    nLinSh      = Number(divSheet.getAttribute('nLinSh'))    ; nColSh    = Number(divSheet.getAttribute('nColSh'))
    nLinInps    = Number(divSheet.getAttribute('nLinInps'))  ; nColInps  = Number(divSheet.getAttribute('nColInps'))
    hWindow     = Number(divSheet.getAttribute('hWindow'))   ; wWindow   = Number(divSheet.getAttribute('wWindow'))
    margT       = Number(divSheet.getAttribute('margT'))     ; margL     = Number(divSheet.getAttribute('margL'))
    lplan0Max   = Number(divSheet.getAttribute('lplan0Max')) ; cplan0Max = Number(divSheet.getAttribute('cplan0Max'))
    
    // . . . redefine lplan0, cplan0
    if (lplanIni>=lFrz){ lplan0 = lplanIni}
    if (cplanIni>=cFrz){ cplan0 = cplanIni}
    // .  .  .   limites lplan0 e cplan0
    if (lplan0>lplan0Max && lplan0Max>0){ lplan0 = lplan0Max}
    if (cplan0>cplan0Max && cplan0Max>0){ cplan0 = cplan0Max}


    Cells[divSheetId][0][0]['lplan0'] = lplan0  ; Cells[divSheetId][0][0]['cplan0'] = cplan0
    // . . .[redefine lplan0, cplan0]

    // . . .  xTop ; yTop
    xTop = Cells[divSheetId][lFrz][cFrz]['left']  - Cells[divSheetId][lFrz][cplan0]['left'] + 1
    if(LinMod==1){ yTop = 0  }
    if(LinMod==0){ yTop = Cells[divSheetId][lFrz][1]['top']   - Cells[divSheetId][lplan0][1]['top']  + 1 }
    // . . . [xTop ; yTop] 

    // . . . define nLinSh (total de linhas, incluindo cabeçalho) e nLinShC (linhas de corpo)
    hPint = 0 ; nLinShC = 0
    for (i = lplan0; i <= nLinPla; i++){
        ii = i  ;   if(LinMod>0){ ii=lFrz }
        
        hPint = hPint + Cells[divSheetId][ii][1]['height'] ; nLinShC++
        if(hPint>hWindow){nLinShC-- ; break}
    }
    nLinSh = nLinShC + lFrz - 1
    divSheet.setAttribute('nLinSh', nLinSh)
    divSheet.setAttribute('iElM', lplan0+nLinShC-1)
    // . . .[define nLinSh]
            
    // . . . define nColSh de corpo
    wPint = 0 ; nColShC = 0
    for (j = cplan0; j <= nColPla; j++){
        wPint = wPint + Cells[divSheetId][lFrz][j]['width'] ; nColShC++
        if(wPint>wWindow){ nColShC-- ; break}
    }
    nColSh = nColShC + cFrz - 1
    divSheet.setAttribute('nColSh', nColSh)
    divSheet.setAttribute('jElM', cplan0+nColShC-1)
    // . . .[define nColSh]

    // . . . ajusta autosize - dimensões de div window (divSheetId)
    if(autoSize=='autosize'){
        xUlt = Cells[divSheetId][lFrz][nColPla]['left'] + Cells[divSheetId][lFrz][nColPla]['width']  +1         + margD
        if(LinMod==0) { yUlt = Cells[divSheetId][nLinPla][1]['top']  + Cells[divSheetId][nLinPla][1]['height']         +1 + margT + margB }
        if(LinMod!=0) { yUlt = Cells[divSheetId][lFrz][1]['top']     + Cells[divSheetId][lFrz][1]['height']*(nLinSh-1) +1 + margT + margB }
        if (xUlt<wSheet) { el(divSheetId).style.width = xUlt+'px' }
        if (yUlt<hSheet) { el(divSheetId).style.height= yUlt+'px' }
    }
    // . . .[ajusta autosize - dimensões de div window (divSheetId)]
    // ------------------

    // .... renomeia para provisório
    for(i=1; i<=nLinInps; i++){for (j=1; j<=nColInps; j++){ el(divSheetId, i, j).id = 'provCell_'+i+':'+j }}
    // .... renomeia para provisório
    iSub = 0
    for (i = 1; i <= nLinInps; i++){  for (j = 1; j <= nColInps; j++){
        cell = el('provCell', i, j)
        iC = Number(cell.getAttribute('iElN')) ; jC = Number(cell.getAttribute('jElN'))

        // ----------- dentro de window
        dentro = 0
        if(iC>0 && jC>0){
            iMax = lplan0+nLinShC-1 ; jMax = cplan0+nColShC-1
            // corpo
            if( (iC>=lplan0 && iC<=iMax) && (jC>=cplan0 && jC<=jMax) )      { dentro = 1 }
            // cabeçalho
            if(iC<lFrz && ( jC<cFrz || (jC>=cplan0 && jC<=jMax) ) )                   { dentro = 1 }
            // lateral
            if(jC<cFrz && ( iC<lFrz || (iC>=lplan0 && iC<=iMax) ) )                   { dentro = 1 }
        }
        // -----------[dentro de window]
        
        // .... print TOTAL - sem reaproveitameto
        if(prTotal==1) { dentro = 0 }

        // . . . renomeia existentes
        if(dentro==1){ 
            iSheet=iC; jSheet=jC; if(iC>=lFrz){iSheet = iC - lplan0 + lFrz}; if(jC>=cFrz){jSheet = jC - cplan0 + cFrz}
            cell.id = divSheetId+'_'+iSheet+':'+jSheet
        }
        // . . .[renomeia existentes]

        // . . . marca disponíveis para sub
        if(dentro==0){ 
            cell.setAttribute('iElN',   -1) ; cell.setAttribute('jElN',   -1)
            iSub++; cell.id = 'sub_'+iSub 
        }
    }}
    nSub = iSub
    // ....[renomeia para provisório]

    // ----
    //  i, j - Sheet               iC, jC - Planilha
    // ... printa Cells  -  geom Cells
    iSub = 0
    nLinP = nLinSh ; nColP = nColSh
    if((nLinP+lplan0-lFrz)<nLinPla){ nLinP++ } ; if((nColP+cplan0-cFrz)<nColPla){ nColP++ }
    for (i = 1; i <= nLinP; i++){  for (j = 1; j <= nColP; j++){
        iC = -2
        // . . . célula existente
        try{
            cell = el(divSheetId, i, j) ; iC = Number(cell.getAttribute('iElN')) ; jC = Number(cell.getAttribute('jElN'))
        }catch{}
        // . . .[célula existente]

        if(iC<1){ iSub++ ; el('sub_'+iSub).id = divSheetId+'_'+i+':'+j ; printCell(i, j, divSheetId) }
        
        // geometria de Cell
        //  ... células mescladas
        if(LinMod==0 || iC<lFrz){
            meL = 0 ; iCy = iC ; jCy = jC
            try{
                mergedCols          = Cells[divSheetId][iC][jC]['mergedCols']
                mergedLins          = Cells[divSheetId][iC][jC]['mergedLins']
                if (mergedCols<0) { jCy = jC + mergedCols ; meL = 1 }
                if (mergedLins<0) { iCy = iC + mergedLins ; meL = 1 }
            } catch{}
            if(meL==1){ iC = iCy ; jC = jCy }
        }
        //  ...[células mescladas]
        
        iCf = iC    ; if (LinMod>0 && iCf>lFrz) { iCf = lFrz }
        
        yTopP = yTop ; xTopP = xTop
        if(iC<lFrz)  { yTopP = 0}   ;   if(jC<cFrz)  { xTopP = 0} // lateral e cabeçalho

        cellLef = Cells[divSheetId][lFrz][jC]['left'] + xTopP + margL

        hei = Cells[divSheetId][lFrz][1]['height']*(i-lFrz)
        if(iC<lFrz  || LinMod==0)  {                cellTop = (Cells[divSheetId][iC][jC] ['top']  + yTopP + margT)  }
        if(iC>=lFrz && LinMod>0)   { iCf = lFrz ;   cellTop =  Cells[divSheetId][lFrz][1]['top']  + hei   + margT + 0 }
        
        cell.style.top     = cellTop + 'px'
        cell.style.left    = cellLef + 'px'
        //[geometria de Cell]

    }}
    // ...[printa Cells  -  geom Cells]

    // . . . . devolve nomes para Subs não utilizados - define 'iSheet'
    for (i = 1; i <= nLinInps; i++){  for (j = 1; j <= nColInps; j++){

        cell = el(divSheetId, i, j)
        if(cell==null){ iSub++ ; cell = el('sub_'+iSub) ; el('sub_'+iSub).id = divSheetId+'_'+i+':'+j }
        cell.setAttribute('iSheet', i)  ; cell.setAttribute('jSheet', j)

        // ... expulsa fora dos limites

        iC = Number(cell.getAttribute('iElN'))

        // ...
        if(iC==-1){ cell.style.top  = '-3000px' }        
    }}
    // . . . .[devolve nomes para Subs não utilizados - define 'iSheet']

    // ******* FINAL
    // . . . foco inicial
    if(iElC==0){
        iElC = lplan0 ; jElC = cplan0 
        Cells[divSheetId][0][0]['iElC'] = iElC ; Cells[divSheetId][0][0]['jElC'] = jElC
        iSheet = iElC - lplan0 + lFrz          ; jSheet = jElC - cplan0 + cFrz
        prox = el(divSheetId, lFrz, cFrz)
        prox.focus()    ;   prox.style.zIndex = '5'
        iElN = lFrz ;  jElN = cFrz
    }
    // *******[FINAL]

    // ....
}    
// ----[Preenche Sheet a partir de Cells]

function printCell(i, j, divSheetId){   // i, j  em Sheet
    // ---
    iC = i ; jC = j ; if(i>=lFrz) { iC  = i + lplan0 - lFrz } ; if(j>=cFrz) { jC  = j + cplan0 - cFrz }
    // ---

    zOrd = 2 ; prCell = 1
    if(i<lFrz || j<cFrz) { zOrd = zOrd + 2}

    cell = el(divSheetId, i, j)
    iE = Number(cell.getAttribute('iElN')) ; jE = Number(cell.getAttribute('jElN'))
        
    //  ... células mescladas
    me = 0 ; jCz = jC ; iCz = iC
    if(LinMod==0 || iC<lFrz){
        try{
            mergedCols          = Cells[divSheetId][iC][jC]['mergedCols']
            mergedLins          = Cells[divSheetId][iC][jC]['mergedLins']
            if (mergedCols<0) { jCz = jC + mergedCols ; zOrd = zOrd - 1 ; me = 1 }
            if (mergedLins<0) { iCz = iC + mergedLins ; zOrd = zOrd - 1 ; me = 1 }
        } catch{}
    }
    //  ...[células mescladas]
        
    // ... printa célula de Sheet
    cell.setAttribute('iElN',   iC) ; cell.setAttribute('jElN',   jC)

    // ... parâmetros de Cell
    if(me==1){ iC = iCz ; jC = jCz }
    iCf = iC    ; if (LinMod>0 && iCf>lFrz) { iCf = lFrz }
    backgroundColor     = Cells[divSheetId][iCf][jC]['backgroundColor']
    color               = Cells[divSheetId][iCf][jC]['color']
    textAlign           = Cells[divSheetId][iCf][jC]['textAlign']
    fontFamily          = Cells[divSheetId][iCf][jC]['fontFamily']
    fontSize            = Cells[divSheetId][iCf][jC]['fontSize']
    fontWeight          = Cells[divSheetId][iCf][jC]['fontWeight']
    fontStyle           = Cells[divSheetId][iCf][jC]['fontStyle']
    textDecorationLine  = Cells[divSheetId][iCf][jC]['textDecorationLine']
    readO               = Cells[divSheetId][iCf][jC]['readonly']
    width               = Cells[divSheetId][iCf][jC]['width']
    height              = Cells[divSheetId][iCf][jC]['height']
    // ...[parâmetros de Cell]
    
    // ---- TRAP  DE VALOR
    valor = ''
    // ...  arquivo JS
    try{ valor   = ListPlan[divSheetId][iC][jC] } catch{}
    planValorTrap(divSheetId, iC, jC)
    cell.value                      = valor
    // ----[TRAP  DE VALOR]

    // format Cell
    cell.style.zIndex               = zOrd
    cell.style.backgroundColor      = backgroundColor
    cell.style.color                = color
    cell.style.textAlign            = textAlign
    cell.style.fontFamily           = fontFamily
    cell.style.fontSize             = fontSize  + 'px'
    cell.style.fontWeight           = fontWeight
    cell.style.fontStyle            = fontStyle
    cell.style.textDecorationLine   = textDecorationLine
    cell.style.width                = width  + 'px'
    cell.style.height               = height + 'px'

    cell.removeAttribute('readonly')
    if (readO=='true')  { cell.setAttribute('readonly', true)  }

    desBordas(cell)
    // ...[printa célula de Sheet]

    // ...
}
// ------- function printCell()

// ----- Desenha Bordas
function desBordas(eleBord, fomBord=''){
    divSheet  = eleBord.parentElement ; divSheetId = divSheet.id
    iCf       = Number(eleBord.getAttribute('iElN'))    ; jC   = Number(eleBord.getAttribute('jElN'))

    // bordas
    if (iCf>0)  {
        lFrz        = Cells[divSheetId][0][0]['lFrz'] ; LinMod    = Cells[divSheetId][0][0]['LinMod']
        if(iCf>=lFrz && LinMod>0)   { iCf = lFrz }        
        bordas      = Cells[divSheetId][iCf][jC]['bordas'] 
    }
    if (iCf==0) { bordas      = fomBord }

    // .... corrige bordas

    eleBord.style.borderColor = 'black'

    eleBord.style.borderBottomWidth    = '0px' ; eleBord.style.borderBottomStyle  = "solid"
    eleBord.style.borderRightWidth     = '0px' ; eleBord.style.borderRightStyle   = "solid"
    eleBord.style.borderTopWidth       = '0px' ; eleBord.style.borderTopStyle     = "solid"
    eleBord.style.borderLeftWidth      = '0px' ; eleBord.style.borderLeftStyle    = "solid"

    if (bordas[2]!=0 && bordas[2]!=4)   { eleBord.style.borderBottomWidth = bordas[2] + 'px' }
    if (bordas[2]==4)                   { eleBord.style.borderBottomWidth = '3px' ; eleBord.style.borderBottomStyle = "double" }
    
    if (bordas[1]!=0 && bordas[1]!=4)   { eleBord.style.borderRightWidth  = bordas[1] + 'px' }
    if (bordas[1]==4)                   { eleBord.style.borderRightWidth  = '3px' ; eleBord.style.borderRightStyle  = "double" }
    
    if (bordas[0]!=0 && bordas[0]!=4){ eleBord.style.borderTopWidth = bordas[0] + 'px' }
    if (bordas[0]==4){ eleBord.style.borderTopWidth = '3px' ; eleBord.style.borderTopStyle = "double" }

    if (bordas[3]!=0 && bordas[3]!=4){ eleBord.style.borderLeftWidth = bordas[3] + 'px' }
    if (bordas[3]==4){ eleBord.style.borderLeftWidth = '3px' ; eleBord.style.borderLeftStyle = "double" }

    // ...
}
// -----[Desenha Bordas]

// ----- Preenche Pag de Painel
function preenchePainel(cellCurr=0, painelNome='', centra=0, nCelPai=0, porLInha=1){
    // iTopPai  - linha do alto do painel
    //      preenche todas Cell definidas em divArrId
    // VER // ---------- scroll de Painel
    
    // . . . parâmetrod de divArr
    divArrId = painelNome  ;  dictArray = Array[divArrId]
    parDivId    = dictArray['parDivId']
    nLar        = dictArray['nLar']     ; nCar      = dictArray['nCar']     ; hPar = dictArray['hPar'] ; vPar = dictArray['vPar']
    lOri        = dictArray['lOri']     ; tOri      = dictArray['tOri']     ; wOri = dictArray['wOri'] ; hOri = dictArray['hOri']    
    iTpArr      = dictArray['topLin']   ; jLfArr    = dictArray['lefCol']    
    altLin      =  hOri+vPar            ; larCol    = wOri+hPar
    // . . .[parâmetrod de divArr]

    // ..... cria Painel, se inexistente painelNome = arrayClass = divArrId
    try{ a = Painel[painelNome]['novoPainel'] ; novoPainel = a } catch{ novoPainel = 1}    
    if(novoPainel==1){        
        porLInha = 1    ;   if(dictArray['liCo']=='Coluna'){ porLInha = 0}
        parDiv  = el(parDivId) ; parDiv.className = painelNome
        altWind = parseInt(window.getComputedStyle(parDiv).height)  ; larWind = parseInt(window.getComputedStyle(parDiv).width)
        nlWind  = Math.floor(altWind/altLin)                        ; ncWind  = Math.floor(larWind/larCol)
        if(nLar<nlWind){ nlWind = nLar }                            ; if(nCar<ncWind){ ncWind = nCar }        
        slackL  = Math.floor( (nLar-nlWind)/2 )                     ; slackC  = Math.floor( (nCar-ncWind)/2 )
        
        iTDisp  = 1  ;  jLDisp = 1 ; iPaiCurr = 0 ; jPaiCurr = 0

        if(nCelPai==0)  { nCelPai = nLar*nCar }
        if(porLInha==0) { nLinPai = nLar                       ;   nColPai = Math.ceil(nCelPai/nLar) }
        if(porLInha==1) { nLinPai = Math.ceil(nCelPai/nCar)    ;   nColPai = nCar }
 
        Painel[painelNome] = {  'divPainelId': parDivId , 'porLInha': porLInha, 
                                'nCelPai' : nCelPai,      'nLinPai' : nLinPai,   'nColPai' : nColPai,
                                'nlWind'  : nlWind,       'ncWind'  : ncWind,    'altLin'  : altLin,   'larCol': larCol,
                                'slackL'  : slackL,       'slackC'  : slackC,
                                'nLar'    : nLar,         'nCar'    : nCar,      'lOri'   : lOri,     'tOri'  : tOri,
                                'scrTdis' : tOri,         'scrLdis' : lOri,
                                'iTop'    : 0,            'jLeft'   : 0,
                                'iTDisp'  : iTDisp,       'jLDisp'  : jLDisp,
                                'cellCurr': cellCurr,     'iPaiCurr': jPaiCurr,  'jPaiCurr': jPaiCurr,
                                'novoPainel': 0
                             }
    }
    // .....[cria Painel, se inexistente painelNome = arrayClass = divArrId]
    
    // .......... lê parâmetros correntes de Painel
    porLInha    = Painel[painelNome]['porLInha']    
    nCelPai     = Painel[painelNome]['nCelPai']
    nLinPai     = Painel[painelNome]['nLinPai']  ; nColPai  = Painel[painelNome]['nColPai']
    nlWind      = Painel[painelNome]['nlWind']   ; ncWind   = Painel[painelNome]['ncWind']
    slackL      = Painel[painelNome]['slackL']   ; slackC   = Painel[painelNome]['slackC']
    
    cellCurrAtu = Painel[painelNome]['cellCurr']
    iTopAtu     = Painel[painelNome]['iTop']     ; jTopAtu  = Painel[painelNome]['jLeft']
    iTDisp      = Painel[painelNome]['iTDisp']   ; jLDisp   = Painel[painelNome]['jLDisp']

    if(porLInha==1){ iPaiCurr = Math.ceil(cellCurr/nCar)    ; jPaiCurr = cellCurr-(iPaiCurr-1)*nCar }
    if(porLInha==0){ jPaiCurr = Math.ceil(cellCurr/nLar)    ; iPaiCurr = cellCurr-(jPaiCurr-1)*nLar }

    // . . . verifica se iPaiCurr dentro de window atual
    dentro = 0                                                                          // dentro
    if(iPaiCurr<iTDisp)             { dentro = -1 ; iTDisp = iPaiCurr }                 // acima
    if(iPaiCurr>iTDisp+nlWind-1)    { dentro =  1 ; iTDisp = iPaiCurr - nlWind + 1 }    // abaixo
    if(jPaiCurr<jLDisp)             { dentro = -1 ; jLDisp = jPaiCurr }                 // esquerda
    if(jPaiCurr>jLDisp+ncWind-1)    { dentro =  1 ; jLDisp = jPaiCurr - ncWind + 1 }    // direita

    iTopPai = iTDisp - slackL ; jTopPai = jLDisp-slackC

    // . . . 
    if(iTopPai<1) { iTopPai = 1 } ;   if(jTopPai<1) { jTopPai = 1 }

    if(iTopPai> nLinPai - nLar){ iTopPai = nLinPai - nLar + 1 }
    if(jTopPai> nColPai - nCar){ jTopPai = nColPai - nCar + 1 }
    
    // ------ Preenche painel  iTopPai, jTopPai
    if(iTopAtu!=iTopPai || jTopPai!=jTopAtu){
        iPaiIni = iTopPai ; iPaiFin = iTopPai  + nLar - 1
        if(iPaiFin>nLinPai){ iPaiFin = nLinPai }  // menor que nLinPai

        jPaiIni = jTopPai ; jPaiFin =  jTopPai  + nCar - 1
        if(jPaiFin>nColPai){ jPaiFin = nColPai }  // menor que nColPai
        
        // ... de Array, marca cells disponíveis para Subst
        iCsub = 0
        for(iArr=iTpArr ; iArr<iTpArr+nLar ; iArr++){    for(jArr=jLfArr ; jArr<jLfArr+nCar ; jArr++){
            disp = 0

            // . . . fora do range de novo print
            if( (iArr<iPaiIni || iArr>iPaiFin) || (jArr<jPaiIni  || jArr>jPaiFin) || novoPainel==1 ){ disp = 1 }
            if(disp==1){ 
                iCsub = iCsub + 1 ; idSub = 'Sub'+iCsub ; el(divArrId,iArr,jArr).id = idSub
                
                Descen = el(idSub).querySelectorAll('*')    ;   nCh = Descen.length
                for(iCh = 0; iCh<=nCh-1; iCh++){
                    descen = Descen[iCh] ; descId = descen.id
                    posCoo = descId.indexOf("_") ; oriId  = descId.slice(0, posCoo)
                    descen.id = oriId+'_'+iCh
                }
            }
        }}
        // ...[de Array, marca cells disponíveis para Sub]
        
        iCell = 0
        for(iPai=iPaiIni ; iPai<=iPaiFin ; iPai++){
            for(jPai=jPaiIni ; jPai<=jPaiFin ; jPai++){
                
                // *****       
                if(porLInha==1)  { kCell    = (iPai-1)*nCar + jPai }
                if(porLInha==0)  { kCell    = (jPai-1)*nLar + iPai }
                // *****
                
                // ... célula de Painel já preenchida
                cellPres = 0
                try{
                    elCell  = el(divArrId,iPai,jPai)
                    num     = parseInt(elCell.getAttribute('kcelarray'))
                    if (num==kCell){ cellPres = 1 }
                }catch{}
                // ...[célula de Painel já preenchida]
                
                // ...... célula nova reciclada
                if(cellPres==0){
                    iCell = iCell + 1

                    // . . . printa Cell iPai, jPai sobre Array(iSub, jSub)
                    coordsN  = '_'+iPai+':'+jPai
                    novoNome = divArrId + coordsN
                    
                    // ------- célula Sub
                    elCell   = el('Sub'+iCell)
                    
                    // . . . renomeia Cell substituta reciclada
                    elCell.id        = novoNome
                    elCell.setAttribute('iArray' , iPai) ; elCell.setAttribute('jArray' , jPai)
                    elCell.setAttribute('kcelarray' , kCell)

                    // . . . renomeia descendentes
                    Descen = el(novoNome).querySelectorAll('*')
                    nCh = Descen.length
                    for(iCh = 0; iCh<=nCh-1; iCh++){
                        descen = Descen[iCh] ; descId = descen.id
                        
                        posCoo = descId.indexOf("_")
                        oriId  = descId.slice(0, posCoo)
                        descen.id = oriId+coordsN
                        descen.setAttribute('iArray' , iPai) ; descen.setAttribute('jArray' , jPai) ; descen.setAttribute('kcelarray' , kCell)
                    }
                    // . . .[renomeia descendentes]

                    // . . . trap de compõe cell de Arr
                    if(kCell<=nCelPai){ ContentCellArr(divArrId, iPai, jPai, kCell) }
                    // -------
                    // . . .[printa Cell iPai, jPai]

                }
                // ......[célula nova reciclada]

            }
            // . . .[preenche linha]
        }

        // .... Reposiciona Painel
        for(iPai=iPaiIni ; iPai<=iPaiFin ; iPai++){ ;   for(jPai=jPaiIni ; jPai<=jPaiFin ; jPai++){
            oriCell  = el(divArrId,iPai,jPai)
            oriKcell = parseInt(oriCell.getAttribute('kcelarray'))
            topCell  = tOri + (iPai-iPaiIni) * altLin
            leftCell = lOri + (jPai-jPaiIni) * larCol
            // . . . expulsa maiores que máximo
            if(oriKcell>nCelPai){ leftCell = -1000 }

            oriCell.style.top  = topCell +'px'
            oriCell.style.left = leftCell+'px'
        }}
        // ....[Reposiciona Painel]

        // ------[Preenche painel]

    }
    // ------[Preenche painel  iTopPai, jTopPai]

    // scroll div parent
    scrTa = el(parDivId).scrollTop      ; scrLa = el(parDivId).scrollLeft
    scrTc = (iTDisp-iPaiIni)*altLin     ; scrLc = (jLDisp-jPaiIni)*larCol
    desvV =  Math.abs(scrTa-scrTc)      ; desvH =  Math.abs(scrLa-scrLc)

    if(centra!=0 &&  (desvV>altLin*0.75 && porLInha==1)){ el(parDivId).scrollTop  = scrTc ; scrTdis = scrTc }
    if(centra!=0 &&  (desvH>larCol*0.75 && porLInha==0)){ el(parDivId).scrollLeft = scrLc ; scrLdis = scrLc  }

    // . . . atualiza Array e Painel
    Array[divArrId]['topLin']       = iTopPai   ;   Array[divArrId]['lefCol']       = jTopPai
  
    Painel[painelNome]['scrTdis']   = scrTdis   ;   Painel[painelNome]['scrLdis']   = scrLdis
    Painel[painelNome]['iTop']      = iTopPai   ;   Painel[painelNome]['jLeft']     = jTopPai
    Painel[painelNome]['iTDisp']    = iTDisp    ;   Painel[painelNome]['jLDisp']    = jLDisp
    Painel[painelNome]['iPaiCurr']  = iPaiCurr  ;   Painel[painelNome]['jPaiCurr']  = jPaiCurr
    Painel[painelNome]['cellCurr']  = cellCurr
    // . . .[atualiza Array e Painel]
    
    // .... põe e tira cursor
    celEl = el(painelNome, iPaiCurr, jPaiCurr) ; celEl .style.outlineWidth = "1px"
    if(celElA!='' && celElA!=celEl)            { celElA.style.outlineWidth = "0px" }
    celElA = celEl    
    // ....
}    
// -----[Preenche Pag de Painel]


// *************************[FUNCTIONS DE SISTEMA EM .js]
