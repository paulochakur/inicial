/*
    REFERÊNCIAS
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

*/


//  -------------- Variáveis Globais -------------
var fixTop = '0px'
var habClic = 0 ; var iniValueInput
var toques=0
var iCa=0
var List = [ 4 ]
var eventoAc = '.......'

var blocTrap = 0

// ... Zoom
var divOriId = ''       // div da img original
var divZooId = ''       // div da amostra aumentada
var parentZoomId = ''   // parent de divOriId
var imgZoomId = ''      // img de divOriId a ser amostrada
var fatZoom = 1         // fator de aumento
var menuZoom = 0        // flag de zoom ativo

var nIcis=0, proprHab = 0

// ... evento
var evento  = '' ; var Leventos  = []

var eleTa = 1   ; var eleTaId  = 'T' ; var eleTaTy     = 1 ; var taStyle    = 1 ; var eleTaClass = 1
var eleOn = 'O' ; var eleOnId  = 'O' ; var eleOnTy     = 1 ; var onStyle    = 1 ; var eleOnClass = 1
var eleFo = '@' ; var eleFoId  = 'F' ; var eleFoTy     = 1 ; var foStyle    = 1 ; var eleFoClass = 1
var eleTaIdAnt = 0
var eleScrTop = 1

// ... coordenadas
var xMe  = 1    , yMe   = 1 , yMe   = 1 , yMe   = 1                             // relativas ao elemento
var xMw  = 1    , yMw   = 1 , Wh    = 1 , Ww    = 1                             // relativas a  window
var xMs  = 1    , yMs   = 1 , Sh    = 1 , Sw    = 1  , Ah    = 1 , Aw    = 1    // relativas a  screen
var xMp  = 1    , yMp   = 1 , Ph    = 1 , Pw    = 1                             // relativas a  page
var Dh   = 1    , Dw    = 1                                                     // relativas a  document
var xPs  = 1    , xPs   = 1                                                     // scroll de page


var keyCode                         // keydown event
var ctrK                            // keydown event
// ...[evento]

// ... menu
var IsubOpen   = [1,1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]
var divMenuAti = '' ; var popMenuAti = ''

// ... planilha
var multH = { '1':[0] }


var elAnt = ''   ; var zAnt = 0 ; var iElA = 0 ; var jElA = 0
var iMenAnt = 0  ; var linMeIdAnt = ''


// ************ TESTES

/*
chave1 = "nome"  ;     chave2 = "nLin"  ; chave3 = "nCol"  ; chT = "nTot"

dictJs = {chave1:"Plan" , "nLin":30 , chave3:60}
dictJs.nomePlan = "planilha"
dictJs[chT] = 456

dictMt["primeira"] = dictJs

eventoAc = '<br> RODA INI  '     +eventoAc ;  el("geralDiv").innerHTML = eventoAc
*/

//  --------------[Variáveis Globais] -------------

//  --------------Trap de Eventos -------------
Leventos = ["abort", "afterprint", "animationend", "animationiteration", "animationstart", "beforeprint", "beforeunload", "blur", "canplay", "canplaythrough", "change", "click", "contextmenu", "copy", "cut", "dblclick", "drag", "dragend", "dragenter", "dragleave", "dragover", "dragstart", "drop", "durationchange", "ended", "error", "focus", "focusin", "focusout", "fullscreenchange", "fullscreenerror", "hashchange", "input", "invalid", "keydown", "keypress", "keyup", "load", "loadeddata", "loadedmetadata", "loadstart", "message", "mousedown", "mouseenter", "mouseleave", "mousemove", "mouseover", "mouseout", "mouseup", "mousewheel", "offline", "online", "open", "pagehide", "pageshow", "paste", "pause", "play", "playing", "popstate", "progress", "ratechange", "resize", "reset", "scroll", "search", "seeked", "seeking", "select", "show", "stalled", "storage", "submit", "suspend", "timeupdate", "toggle", "touchcancel", "touchend", "touchmove", "touchstart", "transitionend", "unload", "volumechange", "waiting", "wheel"]

// ----- funções de addEventListener
document.addEventListener("DOMContentLoaded", evIni)        ;  function evIni(){ evento = 'inicio'    ; iniSys() }

document.addEventListener(Leventos[00], ev00) ;  function ev00() { evento = Leventos[00]; eventTrap() }
document.addEventListener(Leventos[01], ev01) ;  function ev01() { evento = Leventos[01]; eventTrap() }
document.addEventListener(Leventos[02], ev02) ;  function ev02() { evento = Leventos[02]; eventTrap() }
document.addEventListener(Leventos[03], ev03) ;  function ev03() { evento = Leventos[03]; eventTrap() }
document.addEventListener(Leventos[04], ev04) ;  function ev04() { evento = Leventos[04]; eventTrap() }
document.addEventListener(Leventos[05], ev05) ;  function ev05() { evento = Leventos[05]; eventTrap() }
document.addEventListener(Leventos[06], ev06) ;  function ev06() { evento = Leventos[06]; eventTrap() }
document.addEventListener(Leventos[07], ev07) ;  function ev07() { evento = Leventos[07]; eventTrap() }
document.addEventListener(Leventos[08], ev08) ;  function ev08() { evento = Leventos[08]; eventTrap() }
document.addEventListener(Leventos[09], ev09) ;  function ev09() { evento = Leventos[09]; eventTrap() }
document.addEventListener(Leventos[10], ev10) ;  function ev10() { evento = Leventos[10]; eventTrap() }
document.addEventListener(Leventos[11], ev11) ;  function ev11() { evento = Leventos[11]; eventTrap() }
document.addEventListener(Leventos[12], ev12) ;  function ev12() { evento = Leventos[12]; eventTrap() }
document.addEventListener(Leventos[13], ev13) ;  function ev13() { evento = Leventos[13]; eventTrap() }
document.addEventListener(Leventos[14], ev14) ;  function ev14() { evento = Leventos[14]; eventTrap() }
document.addEventListener(Leventos[15], ev15) ;  function ev15() { evento = Leventos[15]; eventTrap() }
document.addEventListener(Leventos[16], ev16) ;  function ev16() { evento = Leventos[16]; eventTrap() }
document.addEventListener(Leventos[17], ev17) ;  function ev17() { evento = Leventos[17]; eventTrap() }
document.addEventListener(Leventos[18], ev18) ;  function ev18() { evento = Leventos[18]; eventTrap() }
document.addEventListener(Leventos[19], ev19) ;  function ev19() { evento = Leventos[19]; eventTrap() }
document.addEventListener(Leventos[20], ev20) ;  function ev20() { evento = Leventos[20]; eventTrap() }
document.addEventListener(Leventos[21], ev21) ;  function ev21() { evento = Leventos[21]; eventTrap() }
document.addEventListener(Leventos[22], ev22) ;  function ev22() { evento = Leventos[22]; eventTrap() }
document.addEventListener(Leventos[23], ev23) ;  function ev23() { evento = Leventos[23]; eventTrap() }
document.addEventListener(Leventos[24], ev24) ;  function ev24() { evento = Leventos[24]; eventTrap() }
document.addEventListener(Leventos[25], ev25) ;  function ev25() { evento = Leventos[25]; eventTrap() }
document.addEventListener(Leventos[26], ev26) ;  function ev26() { evento = Leventos[26]; eventTrap() }
document.addEventListener(Leventos[27], ev27) ;  function ev27() { evento = Leventos[27]; eventTrap() }
document.addEventListener(Leventos[28], ev28) ;  function ev28() { evento = Leventos[28]; eventTrap() }
document.addEventListener(Leventos[29], ev29) ;  function ev29() { evento = Leventos[29]; eventTrap() }
document.addEventListener(Leventos[30], ev30) ;  function ev30() { evento = Leventos[30]; eventTrap() }
document.addEventListener(Leventos[31], ev31) ;  function ev31() { evento = Leventos[31]; eventTrap() }
document.addEventListener(Leventos[32], ev32) ;  function ev32() { evento = Leventos[32]; eventTrap() }
document.addEventListener(Leventos[33], ev33) ;  function ev33() { evento = Leventos[33]; eventTrap() }
document.addEventListener(Leventos[34], ev34) ;  function ev34() { evento = Leventos[34]; eventTrap() }
document.addEventListener(Leventos[35], ev35) ;  function ev35() { evento = Leventos[35]; eventTrap() }
document.addEventListener(Leventos[36], ev36) ;  function ev36() { evento = Leventos[36]; eventTrap() }
document.addEventListener(Leventos[37], ev37) ;  function ev37() { evento = Leventos[37]; eventTrap() }
document.addEventListener(Leventos[38], ev38) ;  function ev38() { evento = Leventos[38]; eventTrap() }
document.addEventListener(Leventos[39], ev39) ;  function ev39() { evento = Leventos[39]; eventTrap() }
document.addEventListener(Leventos[40], ev40) ;  function ev40() { evento = Leventos[40]; eventTrap() }
document.addEventListener(Leventos[41], ev41) ;  function ev41() { evento = Leventos[41]; eventTrap() }
document.addEventListener(Leventos[42], ev42) ;  function ev42() { evento = Leventos[42]; eventTrap() }
document.addEventListener(Leventos[43], ev43) ;  function ev43() { evento = Leventos[43]; eventTrap() }
document.addEventListener(Leventos[44], ev44) ;  function ev44() { evento = Leventos[44]; eventTrap() }
document.addEventListener(Leventos[45], ev45) ;  function ev45() { evento = Leventos[45]; eventTrap() }
document.addEventListener(Leventos[46], ev46) ;  function ev46() { evento = Leventos[46]; eventTrap() }
document.addEventListener(Leventos[47], ev47) ;  function ev47() { evento = Leventos[47]; eventTrap() }
document.addEventListener(Leventos[48], ev48) ;  function ev48() { evento = Leventos[48]; eventTrap() }
document.addEventListener(Leventos[49], ev49) ;  function ev49() { evento = Leventos[49]; eventTrap() }
document.addEventListener(Leventos[50], ev50) ;  function ev50() { evento = Leventos[50]; eventTrap() }
document.addEventListener(Leventos[51], ev51) ;  function ev51() { evento = Leventos[51]; eventTrap() }
document.addEventListener(Leventos[52], ev52) ;  function ev52() { evento = Leventos[52]; eventTrap() }
document.addEventListener(Leventos[53], ev53) ;  function ev53() { evento = Leventos[53]; eventTrap() }
document.addEventListener(Leventos[54], ev54) ;  function ev54() { evento = Leventos[54]; eventTrap() }
document.addEventListener(Leventos[55], ev55) ;  function ev55() { evento = Leventos[55]; eventTrap() }
document.addEventListener(Leventos[56], ev56) ;  function ev56() { evento = Leventos[56]; eventTrap() }
document.addEventListener(Leventos[57], ev57) ;  function ev57() { evento = Leventos[57]; eventTrap() }
document.addEventListener(Leventos[58], ev58) ;  function ev58() { evento = Leventos[58]; eventTrap() }
document.addEventListener(Leventos[59], ev59) ;  function ev59() { evento = Leventos[59]; eventTrap() }
document.addEventListener(Leventos[60], ev60) ;  function ev60() { evento = Leventos[60]; eventTrap() }
document.addEventListener(Leventos[61], ev61) ;  function ev61() { evento = Leventos[61]; eventTrap() }
document.addEventListener(Leventos[62], ev62) ;  function ev62() { evento = Leventos[62]; eventTrap() }
document.addEventListener(Leventos[63], ev63) ;  function ev63() { evento = Leventos[63]; eventTrap() }
document.addEventListener(Leventos[64], ev64) ;  function ev64() { evento = Leventos[64]; eventTrap() }
document.addEventListener(Leventos[65], ev65) ;  function ev65() { evento = Leventos[65]; eventTrap() }
document.addEventListener(Leventos[66], ev66) ;  function ev66() { evento = Leventos[66]; eventTrap() }
document.addEventListener(Leventos[67], ev67) ;  function ev67() { evento = Leventos[67]; eventTrap() }
document.addEventListener(Leventos[68], ev68) ;  function ev68() { evento = Leventos[68]; eventTrap() }
document.addEventListener(Leventos[69], ev69) ;  function ev69() { evento = Leventos[69]; eventTrap() }
document.addEventListener(Leventos[70], ev70) ;  function ev70() { evento = Leventos[70]; eventTrap() }
document.addEventListener(Leventos[71], ev71) ;  function ev71() { evento = Leventos[71]; eventTrap() }
document.addEventListener(Leventos[72], ev72) ;  function ev72() { evento = Leventos[72]; eventTrap() }
document.addEventListener(Leventos[73], ev73) ;  function ev73() { evento = Leventos[73]; eventTrap() }
document.addEventListener(Leventos[74], ev74) ;  function ev74() { evento = Leventos[74]; eventTrap() }
document.addEventListener(Leventos[75], ev75) ;  function ev75() { evento = Leventos[75]; eventTrap() }
document.addEventListener(Leventos[76], ev76) ;  function ev76() { evento = Leventos[76]; eventTrap() }
document.addEventListener(Leventos[77], ev77) ;  function ev77() { evento = Leventos[77]; eventTrap() }
document.addEventListener(Leventos[78], ev78) ;  function ev78() { evento = Leventos[78]; eventTrap() }
document.addEventListener(Leventos[79], ev79) ;  function ev79() { evento = Leventos[79]; eventTrap() }
document.addEventListener(Leventos[80], ev80) ;  function ev80() { evento = Leventos[80]; eventTrap() }
document.addEventListener(Leventos[81], ev81) ;  function ev81() { evento = Leventos[81]; eventTrap() }
document.addEventListener(Leventos[82], ev82) ;  function ev82() { evento = Leventos[82]; eventTrap() }
document.addEventListener(Leventos[83], ev83) ;  function ev83() { evento = Leventos[83]; eventTrap() }
document.addEventListener(Leventos[84], ev84) ;  function ev84() { evento = Leventos[84]; eventTrap() }

document.addEventListener('DOMMouseScroll', ev84) ; function ev85() { evento = 'DOMMouseScroll'; eventTrap() }
// -----[funções de addEventListener]

// ----- inicio Sys
function iniSys(){

    evento = ''

    
    // ---------- cria elementos de sistema
    // ... cria div - "console"
    para = document.createElement("div")
    el('Corpo').appendChild(para)
    para.id     = "console"
    el("console").style = 'position: fixed; left: 30px; top: 40px; width: 300px; height: 550px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: text; border-color: black; border-radius: 1.0%; border-width: 1.0px; border-style: solid; background-color: #d1cb47; opacity: 1; overflow: auto; z-index: 32 '
    // ...[cria div - "console"]

    // ... cria div - "proprBox"
    para = document.createElement("div") ; el('Corpo').appendChild(para) ; para.id     = "proprBox"
    el("proprBox").style = 'position: fixed; left: 10px; top: 135px; width: 320px; height: 420px; box-sizing: content-box; margin: 0.0px; padding: 20px; cursor: text; border-color: black; border-radius: 1.0%; border-width: 1.0px; border-style: solid; background-color: #fae8af; opacity: 1; overflow: hidden; z-index: 31 '

    padY = 11
    Llabls = ['Elem On :', 'Tipo EleOn   :', 'Parent ElOn :', 'Class EleOn :', 'Focus El       :', 'Elem Target  :', 'Evento :', 'Mouse Ele On  :', 'Mouse Window:', 'Mouse Screen :', 'Mouse Page    :', 'Scroll Page :', '', ]
    LidPar = ['eleon2', 'tipoDiv2', 'parentDiv2', 'classDiv2', 'eleTarDiv2', 'focusDiv2', 'eventoDiv2', 'mouseEleDiv2', 'mouseWinDiv2', 'mouseCscDiv2', 'mousePageDiv2', 'scrollPageDiv2', '', '',]
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

    // ... cria div cursor - "ZoomDivCur"
    para = document.createElement("div")
    el('Corpo').appendChild(para)
    para.id = "ZoomDivCur"
    strSty  = 'position: absolute; left: -3000px; top: -3000px; width: 40.0px; height: 40px; box-sizing: content-box; margin: 0.0px; padding: 0px; cursor: none; border-color: black; border-radius: 0.0%; border-width: 0.0px; border-style: solid; background-color: gray; opacity: 0.4; z-index: 40'
    el("ZoomDivCur").style          = strSty
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





    // --------- Atributos- ida: div-429 ; zoomMenuDiv: ZoomAma ; curLado: 50 ; zoomFat: 3 ;
    
    // ... set atributos definidos em TXT
    allEles = document.getElementsByTagName("*") ; nEles = allEles.length
    Lels = []  ; for (i = 0; i <= nEles-1; i++){ Lels.push(allEles[i].id) }
    
    nIcis = nIcis+1
    print('WWWWWWWWWW nEles: '+nEles+'  nIcis: '+nIcis)

    for (i=1; i < nEles-1; i++) {
        innTex = '**' 
        try{ 
            ele     = el(Lels[i])
            innTex  = ele.innerText
            
        } catch{}
        try {
            teste = innTex.includes('Atributos-')
            if (innTex.includes('Atributos-')) {
                
                innTex = innTex.replace('Atributos-','')  ; innTex = innTex.trim()+';'
                iniAtt = 0 ; iniVal = 1 ; valAtt = '*'; j=0 ; nomAtt = '' ; valAtt = '' ; idAtt = ''
                do {
                    j++
                    iniVal  = innTex.indexOf(":", iniAtt+1)   ; fimDiv  = innTex.indexOf(";", iniVal+1)
                    if (iniVal>-1)  { 
                        nomAtt  = innTex.slice(iniAtt,  iniVal)   ; nomAtt  = nomAtt.trim()
                        valAtt  = innTex.slice(iniVal+1,  fimDiv) ; valAtt  = valAtt.trim()
                        if (nomAtt=='ida') { idAtt = valAtt }
                        if (idAtt==ele.id && nomAtt!='ida') { ele.setAttribute(nomAtt, valAtt) }
                    }        
                    // ..
                    iniAtt = fimDiv+1
                } while (iniVal>0 && fimDiv>0 && j<30)
        
            }
        } catch{}
    }    
    // ...[set atributos definidos em TXT]


    
    print('YYYYY')

    // ........ Zoom
    imgZoomId = 'img-430'
    divOriId    = el(imgZoomId).parentElement.id

    eleId = divOriId 
    ele = el(divOriId)

    print('++++++ Div: '+ele.getAttribute("zoomMenuDiv"))
    print('++++++ Lado: '+ele.getAttribute("curLado"))
    print('++++++ Fat: '+ele.getAttribute("zoomFat"))

    //if (divZooId=='self') { divZooId = divOriId }
    
    // ... cria div para abrigar img de zoom, se não existente
    if (divZooId=='local' || divZooId=='') {
        parentZoomId = el(divOriId).parentElement.id
        divZooId = 'Zoom'
        el(parentZoomId).appendChild(para)       // div de zoom é irmã de divZoom
        para.id     = divZooId
        el(divZooId).style.height           = '150px'
        el(divZooId).style.zIndex           = '1'
        el(divZooId).style.overflow         = 'hidden'
    }
    // ........[Zoom]
    




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
            hSheet      = el(divSheetId).offsetHeight
            wSheet      = el(divSheetId).offsetWidth
            if (auto=='autoload'){ 
                preencheSheet(lplanIni=0, cplanIni=0, divSheetId)
                xUlt = Cells[divSheetId][1][nColPla]['left'] + Cells[divSheetId][1][nColPla]['width']  + 17 + 3
                yUlt = Cells[divSheetId][nLinPla][1]['top']  + Cells[divSheetId][nLinPla][1]['height'] + 17 + 20
                if (autoSize=='autosize'){ 
                    if (xUlt<wSheet) { el(divSheetId).style.width = xUlt+'px' }
                    if (yUlt<hSheet) { el(divSheetId).style.height= yUlt+'px' }
                }
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
        if (textForm!=undefined){ icluiTxtemDiv(divSheetId) }

    }
    // ........[Inclui TEXTOS]

    // ........ Monta Menus
    for (me in Menus){
        formMenu    = Menus[me]
        auto        = formMenu[0][1]
        pop = me.includes("pop")
        //if (auto=='autoload' && !pop){ criaMenu(me) }
        if (auto=='autoload'){ criaMenu(me) }
    }
    // ........[Monta Menus]

    // ....
    // -------- finaliza rotina inicial
    window.scrollTo(0,0)
    iniLoc()
}
// ----- inicio Sys

//  -------------- Trap de Eventos  -------------
function eventTrap() {
    
    // ---- Parâmetros de EVENTO
    // elementos de evento - On, Target, Foco
    eleTa  = event.target
    eleFo  = document.activeElement
    if (eleOn=='O')          { eleOn  = eleFo }
    if (evento=="mousemove") { eleOn  = eleTa }
    // . .

    // On     Element
    eleOnId    = eleOn.id      ;   eleOnClass  = eleOn.class   ;   eleOnTy     = eleOn.tagName ;   onStyle = getComputedStyle(eleOn)                
    // Target Element
    eleTaId    = eleTa.id      ;   eleTaClass  = eleTa.class   ;   eleTaTy     = eleTa.tagName ;   taStyle = getComputedStyle(eleTa)
    // Focus  Element
    eleFoId    = eleFo.id      ;   eleFoClass  = eleFo.class   ;   eleFoTy     = eleFo.tagName ;   foStyle = getComputedStyle(eleFo)

    if (eleOnClass==undefined){ eleOnClass = 'Und'}
    if (eleTaClass==undefined){ eleTaClass = 'Und'}
    if (eleFoClass==undefined){ eleFoClass = 'Und'}

    // --- Mouse Coords / Window  -  Não atuliza se evento=="keyup"
    if (evento!="keyup" && evento!="keydown" && evento!="keypress") {
        // relativo ao elemento On
        xMe = event.offsetX ; yMe = event.offsetY
        // relativo a  window  - inclui scroll bars
        xMw = event.clientX ; yMw = event.clientY   ; Wh = window.innerHeight         ; Ww = window.innerWidth
        // relativo a  screen - inclui task bar
        xMs = event.screenX ; yMs = event.screenY   ; Sh = window.screen.height       ; Sw = window.screen.width
        // available  screen - exculi task bar
                                                    Ah = window.screen.availHeight  ; Aw = window.screen.availWidth
        // relativo a  page / body
        xMp = event.pageX   ; yMp = event.pageY     ; Ph = document.body.clientHeight ; Pw = document.body.clientWidth
    }
    // relativo a  document  - não inclui barras de scroll
    Dh  = document.documentElement.clientHeight ; Dw = document.documentElement.clientWidth

    // Scroll 
    //  elemento
    eleOnScrTop = eleOn.scrollTop   ;   eleOnScrLeft = eleOn.scrollLeft
    eleTaScrTop = eleTa.scrollTop   ;   eleTaScrLeft = eleTa.scrollLeft
    eleFoScrTop = eleFo.scrollTop   ;   eleFoScrLeft = eleFo.scrollLeft
    //  page
    xPs = window.pageXOffset    ; yPs = window.pageYOffset

    // Wheel
    xWh = event.deltaX      ; yWh  = event.deltaY
    xWi = window.screenX    ; yWi = window.screenY

    // Keys
    keyCode = event.keyCode ; ctrK = event.ctrlKey
    
    // .... inibe menu do browser em rightClick
    if(evento=='contextmenu'){ evento = 'rightclick' ; event.preventDefault() }

    //if (evento=='dblclick')  { event.preventDefault() ; event.stopPropagation()}
    //event.stopPropagation()
    /*
    if(evento=='dblclick'){ 
        cc = event.cancelable    
        print('cc: '+cc)
        event.preventDefault()     
    }
    */

    // ----[Parâmetros de EVENTO]

    // **************************************

    // -------- CONSOLE
    //if (evento=='keydown') { print(' keyCode: '+keyCode) }

    // ... copia id                 - ctr"i"
    if (evento=='keydown' && keyCode==73) { navigator.clipboard.writeText(eleOnId) }

    // ... toggle CONSOLE visível   - ctr"c"
    if (evento=='keydown' && keyCode==67) { 
        leftAtu = el('console').style.left
        hab=1
        if (leftAtu=='30px'     && hab==1){ el('console').style.left = '-3000px' ; el('proprBox').style.left = '-3000px' ;  hab=0 ; proprHab = 0}
        if (leftAtu=='-3000px'  && hab==1){ el('console').style.left = '30px'    ; el('proprBox').style.left = '4px'     ;  hab=0 ; proprHab = 1}
        if(el('console').style.zIndex=='32') { proprHab = 0 }
    }
    // ... toggle CONSOLE front/bottom com Propriedades   - click
    if (evento=='click' && (eleTaId=='proprBox' || eleOn.parentElement.id=='proprBox')) { 
        frontAtu = el('console').style.zIndex
        hab=1
        if (frontAtu==30 && hab==1){ el('console').style.zIndex = '32' ; hab=0; proprHab = 0}
        if (frontAtu==32 && hab==1){ el('console').style.zIndex = '30' ; hab=0; proprHab = 1}
    }

    // . . . preenche quadro
    if (eleOnId!='O' && eleOnId!=undefined && proprHab==1) {
        onBg = onStyle.backgroundColor

        el("eleon2").innerHTML                  = eleOnId                
        el("eleon2").style.backgroundColor      = onBg
        el("tipoDiv2").innerHTML                = eleOnTy                        //+" "+eleOn.scrollTop
        el("parentDiv2").innerHTML              = eleOn.parentElement.id         //+" "+window.pageYOffset
        el("classDiv2").innerHTML               = eleOn.class                    //+" "+window.screenY
        el("eleTarDiv2").innerHTML              = eleTaId
        el("focusDiv2").innerHTML               = eleFoId
        
        el("eventoDiv2").innerHTML              = evento
        el("mouseEleDiv2").innerHTML            = "xMe:"+xMe+" yMe:"+yMe
        el("mouseWinDiv2").innerHTML            = "xMw:"+xMw+" yMw:"+yMw+" Wh:"+Wh+" Ww:"+Ww
        el("mouseCscDiv2").innerHTML            = "xMs:"+xMs+" yMs:"+yMs+" Sh:"+Sh+" Sw:"+Sw
        el("mousePageDiv2").innerHTML           = "xMp:"+xMp+" yMp:"+yMp+" Ph:"+Ph+" Pw:"+Pw

        el("scrollPageDiv2").innerHTML          = "xPs:"+xPs+" yPs:"+yPs+" Ah:"+Ah+" Aw:"+Aw
    }
    // . . .[preenche quadro]


    // --------[CONSOLE]

    // ----------- INPUT
    if (eleFoTy=='INPUT') {
        if ( (evento=="keyup" || evento=="keydown" || evento=="focusin" || evento=="click") && (eleFo.type=='text' || eleFo.type=='number') ) {    
            readO = eleFo.getAttribute('readonly')
            
            // ... apaga em primeiro toque
            if(evento=="keydown" && readO!='true' && (keyCode>40 || keyCode==8 || keyCode==46) && toques==0) { eleFo.value = '' }
            // ... escape
            if(evento=="keydown" && readO!='true' && keyCode==27) { eleFo.value = iniValueInput }

            // . . . toques
            if(evento=="focusin") { 
                toques = 0 ; habClic = 0 ; setTimeout( fhabClic, 300) ; iniValueInput = eleFo.value
            } 
            if( readO!='true' && ((evento=="keydown" && keyCode>40) || (evento=="click" && habClic==1))  )   { toques++ }
        
            // .....
        }
    }
    // -----------[IMPUT]

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
                    if (orient=='horizontal' || orient=='pop') { 
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

        if ( blocTrap==0 && (evento=="click" || evento=="keydown" || evento=="focusin" || evento=='input')  ){

            scro = 0 ; keyCodeF = 100
            // . . . índices de input focus em Sheet
            abrePar = eleTaId.indexOf("(") ; fechaPar  = eleTaId.indexOf(")") ; virg = eleTaId.indexOf(",")
            if (abrePar>0){
                iSheet  = parseInt(eleTaId.slice(abrePar+1, virg)) ; jSheet   = parseInt(eleTaId.slice(virg+1, fechaPar))
                if (evento=="click" || evento=="focusin") { keyCodeF = 1 }
            }

            // .... inibe scroll com flechas
            // https://stackoverflow.com/questions/10280250/getattribute-versus-element-object-properties
            readO = eleFo.getAttribute('readonly')
            if(evento=='keydown' && readO=='true' && keyCode<41){ event.preventDefault() }

            // .... parâmetros de Sheet
            nomeSheet = eleFoClass ; divSheet = eleFo.parentElement ; divSheetId = divSheet.id

            nLinSh      = Number(divSheet.getAttribute('nLinSh'))   ; nColSh    = Number(divSheet.getAttribute('nColSh'))
            nLinInps    = Number(divSheet.getAttribute('nLinInps')) ; nColInps  = Number(divSheet.getAttribute('nColInps'))
            hWindow     = Number(divSheet.getAttribute('hWindow'))  ; wWindow   = Number(divSheet.getAttribute('wWindow'))

            // .... parâmetros de Planilha
            nLinPla     = Cells[divSheetId][0][0]['nLinPla']        ; nColPla   = Cells[divSheetId][0][0]['nColPla']
            lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
            lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
            iElC        = Cells[divSheetId][0][0]['iEl']            ; jElC      = Cells[divSheetId][0][0]['jEl']             // foco atual em Pla
            cursor      = Cells[divSheetId][0][0]['cursor']         ; enterMove = Cells[divSheetId][0][0]['enterMove']   
            scrollBars  = Cells[divSheetId][0][0]['scrollBars'] 
            header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']

            // ------ mudança de foco em Sheet
            iEl     = iSheet + lplan0 - lFrz   ; jEl       = jSheet + cplan0 - cFrz
            if (iSheet<lFrz){ iEl = iSheet }
            if (jSheet<cFrz){ jEl = jSheet }
            iElN = iEl ; jElN = jEl 
            
            // ------ linha entrada em cabeçalho
            if (evento=="keydown" && keyCode==13 && eleFoId.includes("-headerLinL")){
                iElN = parseInt( el(nomeSheet+"-headerLinL").value ) ; jElN = jElC ; keyCode = 0
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
            if ( (evento=="keydown" && keyCode<41) || keyCodeF<5){
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

                if (iElN==lFrz && iEl<lFrz)  { iElN = lplan0 }    // transição cabeçalho-corpo
                if (jElN==cFrz && jEl<cFrz)  { jElN = cplan0 }    // transição cabeçalho-corpo

                // ------------- analisa posição de novo foco - iElN , jElN
                hei = Cells[divSheetId][lFrz][1]['height']
                if(iElN<lFrz)  {bF = 0}        
                if(iElN>=lFrz && LinMod==0)  { bF = Cells[divSheetId][iElN][jElN]['top'] - Cells[divSheetId][lplan0][jElN]['top'] + Cells[divSheetId][iElN][jElN]['height'] }
                if(iElN>=lFrz && LinMod>0)   { bF = (iElN - lplan0 + 1)*hei}

                xF = Cells[divSheetId][iElNf][jElN]['left'] - Cells[divSheetId][1][cplan0]['left']    
                wF = Cells[divSheetId][iElNf][jElN]['width']
                dF = xF + wF
                // -------------[analisa posição de novo foco - iElN , jElN]
                // ---- define lplan0 e cplan0
                // . . . estouro ACIMA
                if (iElN<lplan0 && iElN>=lFrz) { lplan0 = iElN  ;  Cells[divSheetId][0][0]['lplan0'] = lplan0 ;     scro = 1 }
                // . . . estouro ESQUERDA
                if (jElN<cplan0 && jElN>=cFrz) { cplan0 = jElN  ;  Cells[divSheetId][0][0]['cplan0'] = cplan0 ;     scro = 2 }
                // . . . estouro ABAIXO
                if (bF>hWindow)         {                       scro = 3    ; hPint = 0
                    if(LinMod==0){
                        for (i = iElN; i >= lFrz; i--){
                            jj = 1
                            if (i==iElN) { jj = jElN }    
                            hPint = hPint + Cells[divSheetId][i][jj]['height']
                            if(hPint>hWindow){ lplan0 = i + 1; i = 0}
                        }
                    }
                    if(LinMod>0){lplan0 = iElN - nLinSh + lFrz}
                    Cells[divSheetId][0][0]['lplan0'] = lplan0
                }
                // . . . estouro DIREITA
                if (dF>wWindow)         {                       scro = 4    ; wPint = 0
                    for (j = jElN; j >= cFrz; j--){
                        ii = 1
                        if (j==jElN) { ii = iElNf }
                        wPint = wPint + Cells[divSheetId][ii][j]['width']
                        if(wPint>wWindow){ cplan0 = j + 1; j = 0}
                    }
                    Cells[divSheetId][0][0]['cplan0'] = cplan0
                }            
                // ----[define lplan0 e cplan0]

                //
                if (scro>0){ preencheSheet(lplanIni=0, cplanIni=0, divSheetId) }
                // .....[analisa posição de novo foco]

                // ----- foco em próximo
                // ... 
                iSheet = iElN - lplan0 + lFrz       ;   jSheet = jElN - cplan0 + cFrz
                if (iElN<lFrz){ iSheet = iElN }
                if (jElN<cFrz){ jSheet = jElN }
                prox = nomeSheet+":(" + iSheet + "," + jSheet +")"
                
                if (elAnt!=''){
                    iEl     = iElA + lplan0 - lFrz   ; jEl       = jElA + cplan0 - cFrz
                    mergedColsA  = Cells[divSheetId][iElNf][jElN]['mergedCols']
                    mergedLinsA  = Cells[divSheetId][iElNf][jElN]['mergedLins']
                    el(elAnt).style.zIndex = '2'
                    if(mergedColsA!=1 || mergedLinsA!=1){ el(elAnt).style.zIndex = '1' }
            
                    if (cursor!='outline'){ desBordas(iEl, jEl, el(elAnt)) }
                }
                
                elAnt = prox  ; zAnt = el(prox).style.zIndex  ; iElA = iSheet ; jElA = jSheet
                el(prox).style.zIndex = '3'
                
                if (cursor!='outline')  { el(prox).style.borderColor = cursor ; el(prox).style.borderWidth = '2px'; el(prox).style.outlineWidth      = "0px" }
                if (cursor=='outline')  { el(prox).style.outlineWidth      = "1px" }

                blocTrap = 1
                if (prox!=''){ 
                    if(header=='header'){ 
                        head = el(nomeSheet+"-headerLinL")   ; head.value = iElN
                        head = el(nomeSheet+"-headerLinN")   ; head.value = ' de '+nLinPla
                    }
                    Cells[divSheetId][0][0]['iEl'] = iElN ; Cells[divSheetId][0][0]['jEl'] = jElN 
                    el(prox).focus()   
                    if (scro>0){ iniValueInput = el(prox).value }
                    
                    if (scrollBars==1){
                        el(nomeSheet+"-vertScroll").value  = parseInt( ((iElN-0)/(nLinPla-0) )*100 )
                        el(nomeSheet+"-horizScroll").value = parseInt( ((jElN-0)/(nColPla-0) )*100 )
                    }
                }
                blocTrap = 0
                // -----[foco em próximo]
                
                iSheetA = iSheet ;  jSheet = jSheet
            
            }    
        }
        // ------ movimento em Planilha - Tecla - NOVO FOCO

    
        // ......        
    }
    // ----------------[Movimento em Sheet]


    // ------------------
    // ... eventTrap()
    trapEvent() // executa trap específicos do projeto em .html
}    
//  --------------[Trap de Eventos] -------------

            
// ************************* FUNCTIONS DE SISTEMA EM .js

// ---- print em console
function print(strPrint){ eventoAc = '<br>'+strPrint +eventoAc ; el("console").innerHTML = eventoAc ; return }
// ---

// ---- Get El by Id
function el(idS){ return document.getElementById(idS) }
// ---

// ---- substitui caracteres especiais de strO
function charSpeci(strO){
    iEsp = 0
    while (iEsp>-1){
        if (iEsp==0){ iEsp = -1 }
        iEsp = strO.indexOf("&#", iEsp+1)   ;   iVir = strO.indexOf(";" , iEsp+1)
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
        if (orient=='horizontal')   { IsubOpen   = [1, 1, -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]}
        if (orient=='pop')          { IsubOpen   = [1,-1, -1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1]}
        formMenu[0][20] = IsubOpen ; divMenuAti=''
        if (orient=='vertical')     { retorna   = 1}
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
        
        if (orient=='horizontal' || orient=='pop') {
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
    formGer         = formMenu[0]        ; nLins = formGer[0] ; orient = formGer[2] ; autosize = formGer[4]
  
    // ... cria div para popMenu
    if (orient=='pop') {
        
        para = document.createElement("div")
        el('Corpo').appendChild(para);
        para.id     = divMenuId
        para.class  = divMenuId+'-Menu'
        para.style  = "position: absolute; box-sizing: border-box; border-width: 0px; border-color: black; outline-width: 0; background-color: red "
        popMenuAti      = para

    }
    // ...[cria div para popMenu]

    divMenu         = el(divMenuId)
    // --------- prévio
    
    // --------------- cria subMenus
    nSubMenus = 0 ; subF = 0
    nivel = 0 ; nivelA = -1 ; iSub = -1 ; iSubLa = -1 ; topRel = 0 ; iSubN = 0 ; topAcu = 0 ; popMenu = 1
    formMenu[0][9]  = 0   // iSub da linha 0
    formMenu[0][11] = 0   // pai da linha 0
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
        if (orient=='horizontal' || orient=='pop' ) {
            
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
    if(orient=='horizontal' || orient=='pop'){
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
    divMenu.style.zIndex   = (1)

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
        padT    = (hl-sF)/2

        // . . . formata divs de menu
        // .. Esq
        para = document.createElement("div")
        divMenu.appendChild(para)
        para.id         = divMenuId+ "-Menu(" +i+ ")"
        para.class      = divMenuId+'-Menu'
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

        para.style.zIndex                   = (1+nivel)
        para.style.padding                  = "0px"
        para.style.paddingTop               = padT+"px"
        para.style.paddingLeft              = padL+"px"
        para.style.margin                   = "0px"
        para.style.textIndent               = '0px'
        para.style.textDecorationThickness  = '1px'
        para.style.backgroundColor          = bc
        desBordas(0, 0, para, brd)

        // .. Dir
        paraD = document.createElement("div");
        divMenu.appendChild(paraD)
        paraD.id        = divMenuId+ "-Menu(" +i+ ")D"
        paraD.class     = divMenuId+'-Menu'
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

        paraD.style.zIndex                  = (1+nivel)
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

// ------ Inclui texto em div
function icluiTxtemDiv(divSheetId) {
    // ...... 
    textForm    = Texts[divSheetId]
    texto       = textForm[0]
    ConfGer     = textForm[1]

    divSheet = el(divSheetId)
    
    para = document.createElement("div")
    divSheet.appendChild(para)
    para.id     = divSheetId+"-texto"
    
    // ... format Geral
    margTop = ConfGer[0]    ;   margEsq = ConfGer[1]    ;   margDir = ConfGer[2]
    escaTot = ConfGer[3]
    bordWid = ConfGer[4]    ;   bordCor = ConfGer[5]    ;   paddDt  = ConfGer[6]    ;   anchorDt  = ConfGer[7]
    // ... parâmetros de DIV
    divSheet.style.margin              = '0px'
    divSheet.style.padding             = '0px'
    divSheet.style.textShadow          = '0px 0px 0px'

    hDiv    = parseInt(window.getComputedStyle(divSheet).height)
    scrDiv  = window.getComputedStyle(divSheet).overflow

    // ... parâmetros de textDiv
    para.style.position                 = 'absolute'
    para.style.boxSizing                = 'border-box'  // 'content-box'  'border-box'
    para.style.overflow                 = 'visible'

    para.style.color                    = 'black'
    para.style.backgroundColor          = 'transparent'
    para.style.borderWidth              = bordWid+'px'
    para.style.borderColor              = bordCor
    para.style.borderStyle              = 'solid'
    
    para.style.margin                   = '0px'
    para.style.marginTop                = '0px'
    para.style.marginLeft               = '0px'
    para.style.marginRight              = margDir+'px'
    para.style.padding                  = paddDt+'px'

    para.style.textIndent               = '0px'
    para.style.width                    = (divSheet.clientWidth - (margEsq + margDir))  +'px'
    para.style.height                   = 'auto'
    para.style.fontSize                 = '20px'
    para.style.textDecorationThickness  = 'from-font'
    para.style.whiteSpace               = 'pre-wrap'
    para.style.textAlign                = anchorDt 
    para.style.textJustify              = 'inter-word'
    para.style.opacity                  = 1
    
    para.innerHTML = texto
    
    // . . limites da caixa de texto
    tD = 0  ;   lD = 0
    
    wD = parseInt(window.getComputedStyle(para).width)
    hD = parseInt(window.getComputedStyle(para).height)

    sH = parseInt(para.scrollHeight)        //  Returns the Height of an element, including padding, excluding borders, scrollbars or margins.
    sW = parseInt(para.scrollWidth)         //  Both scrollWidth and scrollHeight return the entire height and width of an element, including what is not viewable (because of overflow).
    
    // ------  corrige x e y
    // . . inclui scroll
    cW = para.clientWidth +bordWid*2                        //  It includes padding but excludes borders, margins, and vertical   scrollbars (if present).
    if (hD>hDiv && scrDiv!='hidden'){ cW  = cW  - 17}
    leftN = lD - (cW*(1-escaTot))/2
    para.style.left             = (leftN+margEsq)+'px'
    para.style.width            = cW+'px'
    
    cH = parseInt(para.clientHeight)+bordWid*2              //  It includes padding but excludes borders, margins, and horizontal scrollbars (if present).
    topN  = tD - (cH*(1-escaTot))/2    
    para.style.top              = (topN+margTop)+'px'
    para.style.height           = cH+'px'
    // ------ [corrige x e y]

    // ... aplica escala
    para.style.transform        = 'scaleX('+escaTot+') scaleY('+escaTot+')'
    corrTam = 2 ;   trapEvent()

    // ...
}
// ------[Inclui texto em div]
    

// ---- Cria Sheet baseado em dictFormPla e MatrCellsPla
function criaSheet(divSheetId){

    nLinPla     = Cells[divSheetId][0][0]['nLinPla']        ; nColPla   = Cells[divSheetId][0][0]['nColPla']
    lplan0      = Cells[divSheetId][0][0]['lplan0']         ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    lFrz        = Cells[divSheetId][0][0]['lFrz']           ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    iElC        = Cells[divSheetId][0][0]['iEl']            ; jElC      = Cells[divSheetId][0][0]['jEl']             // foco atual em Pla
    cursor      = Cells[divSheetId][0][0]['cursor']         ; enterMove = Cells[divSheetId][0][0]['enterMove']
    header      = Cells[divSheetId][0][0]['header']         ; LinMod    = Cells[divSheetId][0][0]['LinMod']
    scrollBars  = Cells[divSheetId][0][0]['scrollBars']     ; autoSize  = Cells[divSheetId][0][0]['autosize']

    Cells[divSheetId][0][0]['iEl']      = 0      ; Cells[divSheetId][0][0]['jEl']    = 0
    Cells[divSheetId][0][0]['lplan0']   = lFrz   ; Cells[divSheetId][0][0]['cplan0'] = cFrz

    lplan0  = lFrz  ; cplan0    = cFrz

    // ... substitui caracteres especiais
    nLin = nLinPla
    if(LinMod> 0) {nLin = lFrz}
    for (i = 1; i <= nLin; i++){
        for (j = 1; j <= nColPla; j++){
            try {
                val     = Cells[divSheetId][i][j]['valor']
                strO    = charSpeci(val)
                Cells[divSheetId][i][j]['valor'] = strO                       
            }
            catch{}

        }
    }
    // ... [substitui caracteres especiais]

   
    // ... Sheet de Inputs
    divSheet = el(divSheetId)   ;   nomeSheet = divSheetId+'-Pla'

    // ... parâmetros de Sheets, definidos em Jiboia
    // calculados
    margT = 0 ; margL = 0 ; margD = 0 ; margB = 0
    if (scrollBars==1)      { margD = 18 ; margB = 18 }
    if (header=='header')   { margT = 20 ; margL = 0 }
    yTbord    = parseInt(window.getComputedStyle(divSheet).borderTopWidth)  ; yBbord    = parseInt(window.getComputedStyle(divSheet).borderBottomWidth)
    xEbord    = parseInt(window.getComputedStyle(divSheet).borderLeftWidth) ; xDbEord   = parseInt(window.getComputedStyle(divSheet).borderRightWidth)
    hSheet    = divSheet.offsetHeight                                       ; wSheet    = divSheet.offsetWidth
    hUtil     = hSheet - yTbord - yBbord - margT                            ; wUtil     = wSheet - xEbord - xDbEord - margL       
    hWindow   = hUtil - margB - Cells[divSheetId][lFrz][1]['top'] - 1       ; wWindow   = wUtil - margD - Cells[divSheetId][1][cFrz]['left'] - 1

    nLinSh  = 0 ; nColSh  = 0
    nLinInps = 70 ; nColInps = 30

    if(LinMod>0){ 
        nLinSh = Math.floor( hWindow / Cells[divSheetId][lFrz][1]['height'] ) 
        Lmult = []
        for (i = 0; i <= 100; i++){  Lmult.push(i*Cells[divSheetId][lFrz][1]['height']) }
        multH[divSheetId] = Lmult
    }

    // guarda
    divSheet.setAttribute('nomeSheet', nomeSheet)
    divSheet.setAttribute('nLinSh', nLinSh)             ; divSheet.setAttribute('nColSh'   , nColSh)
    divSheet.setAttribute('nLinInps', nLinInps)         ; divSheet.setAttribute('nColInps'   , nColInps)
    divSheet.setAttribute('hWindow', hWindow)           ; divSheet.setAttribute('wWindow'   , wWindow)
    divSheet.setAttribute('margL', margL)               ; divSheet.setAttribute('margT'   , margT)
    
    // . . . para leitura
    nLinSh      = Number(divSheet.getAttribute('nLinSh'))   ; nColSh    = Number(divSheet.getAttribute('nColSh'))
    nLinInps    = Number(divSheet.getAttribute('nLinInps')) ; nColInps  = Number(divSheet.getAttribute('nColInps'))
    hWindow     = Number(divSheet.getAttribute('hWindow'))  ; wWindow   = Number(divSheet.getAttribute('wWindow'))
    margL       = Number(divSheet.getAttribute('margL'))    ; margT     = Number(divSheet.getAttribute('margT'))

    // ... cria Inputs
    for (i = 1; i <= nLinInps; i++){
        for (j = 1; j <= nColInps; j++){
            para = document.createElement("INPUT")
            divSheet.appendChild(para);
            para.id     = nomeSheet+":("+i+","+j+")"
            para.class  = nomeSheet
            para.style  = "position: absolute; left:"+(j*100-90)+"px; top:"+(i*30+10)+"px; width: 99px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; outline-width: 0;"
            para.style.borderWidth      = "0px"                
            para.style.textDecorationThickness   = '1px'
        }
    }
    // ...[cria Inputs]

    // ... cria header
    if (header=='header'){        
        para = document.createElement("DIV")
        divSheet.appendChild(para)
        para.id     = nomeSheet+"-fundoHeader"
        para.class  = nomeSheet
        para.style  = 'position: absolute; left: 0px; top: 0px; height: 20px; width:'+(wUtil)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: ridge; border-color: black; outline-width: 1;'
        para.style.borderWidth      = "0px" ;   para.style.borderBottomWidth      = "1px"   
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        // 
        para = document.createElement("INPUT") ; para.setAttribute('readonly', 'readonly') ; para.value = 'Lin:'
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinS"
        para.class  = nomeSheet
        para.type="text"
        para.style  = 'position: absolute; left: 1px; top: 0px; height: 19px; width: 25px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"   
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        
        para = document.createElement("INPUT")
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinL"
        para.class  = nomeSheet
        para.type="text"
        para.style  = 'position: absolute; left: 31px; top: 0px; height: 19px; width: 50px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"  
        para.style.backgroundColor  = '#E4E4E4'
        para.style.zIndex = '4'
        
        para = document.createElement("INPUT") ; para.setAttribute('readonly', 'readonly')
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-headerLinN"
        para.class  = nomeSheet
        para.type="text"
        para.style  = 'position: absolute; left: 81px; top: 0px; height: 19px; width: 80px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
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
        para.class  = nomeSheet
        para.style  = 'position: absolute; left: '+(margL)+'px; bottom: 0px; height: 17px; width:'+(wUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
        fundoHor = para
        // 
        para = document.createElement("INPUT");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-horizScroll"
        para.class  = nomeSheet
        para.type="range"
        para.style  = 'position: absolute; left: 0px; bottom: 1px; height: 17px; width:'+(wUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black;'
        para.style.borderWidth      = "0px"   
        para.style.zIndex = '4'

        // . . . vertical
        // fundo vertical
        para = document.createElement("DIV");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-fundoVert"
        para.class  = nomeSheet
        para.style  = 'position: absolute; right: 0px; top:'+(margT)+'px; height: '+(hUtil-17)+'px; width: 17px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black ; '
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
        fundoVer = para
        // 
        para = document.createElement("INPUT");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-vertScroll"
        para.class  = nomeSheet
        para.type="range"
        para.style  = 'position: absolute; top:'+( (hUtil-17)/2 - 16/2 + margT) +'px; right:'+( -(hUtil-17)/2 + 16/2 )+'px; height: 16px; width:'+(hUtil-17)+'px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; transform: rotateZ(90deg);'
        para.style.borderWidth      = "2px"   
        para.style.zIndex = '4'

        // canto
        para = document.createElement("DIV");
        divSheet.appendChild(para);
        para.id     = nomeSheet+"-fundoCanto"
        para.class  = nomeSheet
        para.style  = 'position: absolute; bottom: 0px; right: 0px; height: 17px; width: 17px; box-sizing: border-box; margin: 0px; pading: 0px; border-style: solid; border-color: black; transform: rotateZ(90deg);'
        para.style.borderWidth      = "1px"   
        para.style.backgroundColor  = '#eaeaea'
        para.style.zIndex = '4'
    }
    // ...[cria barras de scroll]
    
    // ... ajusta autosize
    if (autoSize=='autosize'){
        // . . . 
    


    }
    // ...ajusta autosize]

    // ...
}           
// ----[Cria Sheet baseado em dictFormPla e MatrCellsPla]

// ---- Preenche Sheet a partir de Cells
function preencheSheet(lplanIni=0, cplanIni=0, divSheetId){
    //  lplanIni, cplanIni - da célula de Cells a aparecer no Top-Left Most
    // ... cria Sheet em Div, se inexistente
    divSheet = el(divSheetId)
    if (divSheet.getAttribute('nLinSh')==null){
        criaSheet(divSheetId)
    }
    // ...[cria Sheet em Div, se inexistente]
    // . . . parâmetros de Planilha
    nLinPla  = Cells[divSheetId][0][0]['nLinPla']       ; nColPla   = Cells[divSheetId][0][0]['nColPla']
    lplan0   = Cells[divSheetId][0][0]['lplan0']        ; cplan0    = Cells[divSheetId][0][0]['cplan0']
    lFrz     = Cells[divSheetId][0][0]['lFrz']          ; cFrz      = Cells[divSheetId][0][0]['cFrz']
    iElC     = Cells[divSheetId][0][0]['iEl']           ; jElC      = Cells[divSheetId][0][0]['jEl']             // foco atual em Pla
    cursor   = Cells[divSheetId][0][0]['cursor']        ; enterMove = Cells[divSheetId][0][0]['enterMove']   ;   scrollBars = Cells[divSheetId][0][0]['scrollBars'] 
    header   = Cells[divSheetId][0][0]['header']        ; LinMod    = Cells[divSheetId][0][0]['LinMod']

    if (lplanIni>=lFrz){ lplan0 = lplanIni}
    if (cplanIni>=cFrz){ cplan0 = cplanIni}
    Cells[divSheetId][0][0]['lplan0'] = lplan0  ; Cells[divSheetId][0][0]['cplan0'] = cplan0
    lplanIni = lplan0                           ; cplanIni = cplan0

    // ....  parâmetros de Sheet
    nLinSh      = Number(divSheet.getAttribute('nLinSh'))   ; nColSh    = Number(divSheet.getAttribute('nColSh'))
    nLinInps    = Number(divSheet.getAttribute('nLinInps')) ; nColInps  = Number(divSheet.getAttribute('nColInps'))
    hWindow     = Number(divSheet.getAttribute('hWindow'))  ; wWindow   = Number(divSheet.getAttribute('wWindow'))
    margT       = Number(divSheet.getAttribute('margT'))    ; margL     = Number(divSheet.getAttribute('margL'))

    // . . . define nLinSh
    yTop = 0 ; xTop = Cells[divSheetId][1][cFrz]['left']  - Cells[divSheetId][1][cplan0]['left'] + 1
    if(LinMod==0){ yTop = Cells[divSheetId][lFrz][1]['top']   - Cells[divSheetId][lplan0][1]['top']  + 1 }
    hPint = 0 ; nLinSh = lFrz-1
    for (i = lplan0; i <= nLinPla; i++){
        ii = i
        if(LinMod>0){ ii=lFrz }
        hPint = hPint + Cells[divSheetId][ii][1]['height'] ; nLinSh++
        if(hPint>hWindow){ nLinSh-- ; break}    
    }
    divSheet.setAttribute('nLinSh', nLinSh)
        
    // . . . define nColSh
    wPint = 0 ; nColSh = cFrz-1
    for (i = cplan0; i <= nColPla; i++){
        wPint = wPint + Cells[divSheetId][1][i]['width']  ; nColSh++
        if(wPint>wWindow){ nColSh-- ; break}    
    }
    divSheet.setAttribute('nColSh', nColSh)

    // ----

    //  i, j - Sheet               iC, jC - Planilha
    // ... corpo
    for (i = 1; i <= nLinSh; i++){
        for (j = 1; j <= nColSh; j++){
            printCell(i, j, divSheet)
        }
    }

    // ... rodapé
    for (i = nLinSh+1; i <= nLinInps; i++){
        for (j = 1; j <= nColInps; j++){
            cell = el(divSheetId+"-Pla:("+i+","+j+")")
            iC = i + lplanIni - lFrz ; jC = j + cplanIni - cFrz
            if(iC<=nLinPla) { printCell(i, j, divSheet) }
            if(iC> nLinPla) { cell.style.top = '-3000px'   }
        }
    }

    // ... lateral DIREITA
    for (i = 1; i <= nLinInps; i++){
        for (j = nColSh+1; j <= nColInps; j++){
            cell = el(divSheetId+"-Pla:("+i+","+j+")")
            iC = i + lplanIni - lFrz ; jC = j + cplanIni - cFrz
            if(jC<=nColPla) { printCell(i, j, divSheet) }
            if(jC> nColPla) { cell.style.left = '-3000px' }
        }
    }

    //  ... células mescladas
    for (i = 1; i <= nLinSh+1; i++){
        for (j = 1; j <= nColSh+1; j++){
            iC = i + lplanIni - lFrz ; jC = j + cplanIni - cFrz
            if(i<lFrz)  { iC = i}
            if(j<cFrz)  { jC = j}        
            if( jC<=nColPla && iC<=nLinPla){
                iCf = iC
                if (LinMod>0 && iCf>lFrz) { iCf = lFrz }
                mergedCols          = Cells[divSheetId][iCf][jC]['mergedCols']
                mergedLins          = Cells[divSheetId][iCf][jC]['mergedLins']
                if(  (mergedCols<1 && j+mergedCols<cFrz && j>=cFrz) || (mergedLins<1 && i+mergedLins<lFrz && i>=lFrz)   ){ 
                    cplan0R = cplan0 ; lplan0R = lplan0
                    if (mergedCols<1 && j+mergedCols<cFrz && j>=cFrz) { cplan0 = jC+mergedCols ; printCell(i, cFrz, divSheet) ; cell = el(divSheetId+"-Pla:("+i+","+cFrz+")") ; cell.style.zIndex = '1' }
                    if (mergedLins<1 && i+mergedLins<lFrz && i>=lFrz) { lplan0 = iC+mergedLins ; printCell(lFrz, j, divSheet) ; cell = el(divSheetId+"-Pla:("+lFrz+","+j+")") ; cell.style.zIndex = '1' }                
                    cplan0 = cplan0R ; lplan0 = lplan0R                
                }
            }                
        }
    }


    // . . . foco inicial
    if(iElC==0){
        iElC = lplan0 ; jElC = cplan0 
        Cells[divSheetId][0][0]['iEl'] = iElC ; Cells[divSheetId][0][0]['jEl'] = jElC
        iSheet = iElC - lplan0 + lFrz       ;   jSheet = jElC - cplan0 + cFrz
        el(divSheetId+"-Pla:("+lFrz+","+cFrz+")").focus()
    }

    // ....
}    
// ----[Preenche Sheet a partir de Cells]

function printCell(i, j, divSheet){   // i, j  em Sheet
    divSheetId  = divSheet.id
    yTopP = yTop ; xTopP = xTop
    
    // corpo
    if(i>=lFrz)     { iC  = i + lplan0 - lFrz }
    if(j>=cFrz)     { jC  = j + cplan0 - cFrz }  

    // lateral e cabeçalho
    if(i<lFrz)  { iC = i ; yTopP = 0}
    if(j<cFrz)  { jC = j ; xTopP = 0}

    // ... printa célula de Sheet
    if(iC<=nLinPla && jC<=nColPla){
        cell = el(divSheetId+"-Pla:("+i+","+j+")")
        cell.style.zIndex = '2'
            
        // geometria de Cell
        iCf = iC ; hei = parseInt(Cells[divSheetId][lFrz][1]['height'])
        if(iC<lFrz  || LinMod==0)  {               cellTop = (Cells[divSheetId][iC][jC]['top']   + yTopP + margT)  }
        if(iC>=lFrz && LinMod>0)   {iCf = lFrz ;   cellTop =  Cells[divSheetId][lFrz][1]['top']  + multH[divSheetId][i-lFrz] + margT + 1 }

        cell.style.top                  = cellTop + 'px'
        cell.style.left                 = (Cells[divSheetId][1][jC]['left'] + xTopP + margL)  + 'px'
        cell.style.width                = Cells[divSheetId][iCf][jC]['width']                   + 'px'
        cell.style.height               = Cells[divSheetId][iCf][jC]['height']                 + 'px'

        // ... parâmetros de Cell
        backgroundColor     = Cells[divSheetId][iCf][jC]['backgroundColor']
        color               = Cells[divSheetId][iCf][jC]['color']
        textAlign           = Cells[divSheetId][iCf][jC]['textAlign']
        fontFamily          = Cells[divSheetId][iCf][jC]['fontFamily']
        fontSize            = Cells[divSheetId][iCf][jC]['fontSize']
        fontWeight          = Cells[divSheetId][iCf][jC]['fontWeight']
        fontStyle           = Cells[divSheetId][iCf][jC]['fontStyle']
        textDecorationLine  = Cells[divSheetId][iCf][jC]['textDecorationLine']
        readO               = Cells[divSheetId][iCf][jC]['readonly']
        mergedCols          = Cells[divSheetId][iCf][jC]['mergedCols']
        mergedLins          = Cells[divSheetId][iCf][jC]['mergedLins']

        if(mergedCols<0) { cell.style.top                  = '-3000px'}
        if(mergedLins<0) { cell.style.top                  = '-3000px'}

        // ...[parâmetros de Cell]

        // ---- TRAP  DE VALOR
        valor = ''
        if(iC<lFrz  || LinMod==0){ valor = Cells[divSheetId][iC][jC]['valor'] }

        planValorTrap(divSheetId, iC, jC)
        cell.value                      = valor
        //cell.innerHTML                      = valor
        // ----[TRAP  DE VALOR]

        // format Cell
        cell.style.backgroundColor      = backgroundColor
        cell.style.color                = color
        cell.style.textAlign            = textAlign
        cell.style.fontFamily           = fontFamily
        cell.style.fontSize             = fontSize  + 'px'
        cell.style.fontWeight           = fontWeight
        cell.style.fontStyle            = fontStyle
        cell.style.textDecorationLine   = textDecorationLine
        cell.removeAttribute('readonly')
        if (readO=='true')  { cell.setAttribute('readonly', true)  }

        desBordas(iCf, jC, cell)
    }

    // ...
}
// ------- function printCell()

// ----- Desenha Bordas
function desBordas(iCf, jC, eleBord, fomBord=''){
    // bordas
    if (iCf>0)  { bordas      = Cells[divSheetId][iCf][jC]['bordas'] }
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



// *************************[FUNCTIONS DE SISTEMA EM .js]
