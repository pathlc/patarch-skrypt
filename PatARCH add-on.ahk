#Requires AutoHotkey v2.0
#SingleInstance Force
global name := ""

^2::{
    if WinActive("PatARCH [SUKRAKOW]") {
        Send("+3mat")
        Sleep(100)
        Send("{Enter}")
        Sleep(100)
        Send("^{Home}")

        ib := InputBox("liczba materiałów")
        if ib.Result != "OK"
            return
        repeat := ib.Value

        Sleep(200)
        Loop repeat {
            Send("{Del}{Del}{Del}{End}")
            Sleep(20)
            Send("{Enter}[]{Enter}{Down}{Home}")
        }
        Send("^{Home}")
        return
    } else {
        Send("+3mat")
        Sleep(100)
        Send("{Enter}")

        ib := InputBox("liczba materiałów")
        if ib.Result != "OK"
            return
        repeat := ib.Value

        Send("^{Home}")
        Loop repeat {
            Send("{End}:")
            Sleep(100)
            Send("{Enter}[]{Enter}{Down}")
        }
        Send("^{Home}")
        return
    }
}


^3::   ;przenoszenie stopnia zaawansowania do pola podsumowania`
{
    ClipSaved := ClipboardAll()
    A_Clipboard := ""
    SendInput("^c")
    if !ClipWait(1, 1) || A_Clipboard = "" {
        SoundBeep(1500, 120)
        ToolTip("Brak zaznaczenia / schowek pusty.")
        SetTimer(__ToolOff, -1200)
        A_Clipboard := ClipSaved
        return
    }

    full := A_Clipboard

    pos := InStr(full, "stopień zaawansowania", false)
    if !pos {
        SoundBeep(1200, 150)
        ToolTip('Nie znalazłem "Stopień zaawansowania". Operacja przerwana.')
        SetTimer(__ToolOff, -1600)
        A_Clipboard := ClipSaved
        return
    }

    lineStart := InStr(full, "`n", false, pos, -1)
    lineStart := lineStart ? lineStart + 1 : 1
    lineEnd := InStr(full, "`n", false, pos)
    before := SubStr(full, 1, lineStart - 1)
    after := lineEnd ? SubStr(full, lineEnd + 1) : ""
    lower := before . after

    A_Clipboard := lower
    ClipWait(0.5, 1)
    Send("^v")
    Sleep(40)

    Send("+{Tab}")
    Sleep(40)
    A_Clipboard := full
    ClipWait(0.5, 1)
    Send("^v")
    Sleep(40)

    A_Clipboard := ClipSaved
}

    __ToolOff(*) {
        ToolTip()
    }

^4::    ;lista gotowców
{
MsgBox("
(
xplano - prawdop płaski płuco
xadeno - prawdop adeno płuco
xemuc - el morf. utk. chł
xpred - gotowiec na predykcyjne
xsarko - sarkoidoza węzłowa
xelster - Elster
xgastrit - gastritis
xnpd - znamię skórne
xnpc - znamię złożone
xnpj - znamię łączące
xcrad - torbiel korzeniowa
xczaw - torbiel zawiązkowa
xapp - ropowiczy wyrostek
xvs - brodawka łojotokowa
xctr - torbiel włosowa (tricholemmalna)
xcep - torbiel naskórkowa
xcc - przewlekłe zapalenie pęcherzyka
xkam - kamica bez zapalenia
xfnd - rozrost guzkowy tarczycy
xbaso - baziak
xtl - gruczolak cewkowy lg
xtvl - gruczolak cewkowo-kosmkowy lg
xhyp - polip hiperplastyczny
xgt - ziarniniak naczyniasty (po nowemu)
xlip - tłuszczak
xalip - naczyniakotłuszczak
xuropap - nieinwazyjny urotelialny
xteratoma - potworniak dojrzały jajnika
xcin - dysplazja szyjki macicy (biopsja i konizat)
xtbz - fragment/y tkanki tłuszczowej/włóknistej bez istotnych zmian patologicznych
xcx - Z03
wwp(ki,go,im) nabłonek wielowarstwowy płaski (mianownik, dopełniacz, narzędnik)
xtar - szablony raka tarczycy (GUI)
xpdl - skrypt na PD-L1. wybierz numer żeby wybrać nowotwór (GUI)
xlnc - licznik węzłów chłonnych (GUI)
)")
}

^5::{    ;wyświetlenie instrukcji
    MsgBox("
(

CTRL+1: dodawanie korelacji z poziomu zapisanego rozpoznania (działa w Chrome)
CTRL+2: lepszy #mat
CTRL+3: przenoszenie rozpoznania i st. zaawansowania z szablonu (dwie pierwsze linie) do pola podsumowania (kursor w pozycji 0)
CTRL+4: lista skrótów z kodami ICD
CTRL+5: ponowne wyświetlenie tego okna
CTRL+6: otwiera bluebooks w edge
CRTL+7: otwiera CAP w edge
CTRL+9: rezydenci - wybór specjalisty do kontrasygnaty
CTRL+Q: rezydenci - skrót do przekazania wyniku do kontrasygnaty przez specjalistę (kursor w polu rozpoznania)
ALT+Z: specjaliści - dołącz do wyniku jako Diagnozujący 2 (działa w Chrome)

)")
}
^6::    ;otwiera bluebooks w edge
{
    Run("microsoft-edge:https://tumourclassification.iarc.who.int/home")
}

^7::    ;otwiera CAP w edge
{
    Run("microsoft-edge:https://cap.org/protocols-and-guidelines/cancer-reporting-tools/cancer-protocol-templates")
}

^9::{
    global name
    
    ib := InputBox("Wpisz trzyliterowy skrót nazwy specjalisty", "Wybór specjalisty", "w300 h200")
if ib.Result != "OK"
    return

    name := ib.Value
    MsgBox("wybrany specjalista: " name)
}

^q::{
    global name
    if name = "" {
        MsgBox("Najpierw wybierz specjalistę (Ctrl+9).")
        return
    }
    Send("^f")
    Sleep(100)
    Send("Kontras")
    Sleep(50)
    Send("{Esc}")
    Sleep(50)
    Send("{Tab}{Space}{Tab}" name "{Tab}{Space}{Ctrl {Up}")
}

:*:xww::
{
    ShowNodesGui()
}

ShowNodesGui() {
    local nodesGui1, btn

    nodesGui1 := Gui(, "Węzły chłonne")
    nodesGui1.SetFont("s10", "Arial")

    nodesGui1.Add("Text",, "Podaj liczbę węzłów chłonnych znalezionych w preparacie:")
    nodesGui1.Add("Edit", "w60 vNodesFound")

    nodesGui1.Add("Text",, "Podaj liczbę węzłów chłonnych ze stwierdzonym utkaniem raka:")
    nodesGui1.Add("Edit", "w60 vNodesCancer", "0")

    nodesGui1.Add("Text",, "Podaj maksymalny wymiar przerzutowego utkania (mm) [opcjonalnie]:")
    nodesGui1.Add("Edit", "w60 vMaxDim")

    btn := nodesGui1.Add("Button", "Default", "Ok")
    btn.OnEvent("Click", Generate)

    nodesGui1.OnEvent("Escape", GuiEscape)
    nodesGui1.OnEvent("Close", GuiEscape)

    nodesGui1.Show()
}

Generate(btn, info) {
    nodesGui1 := btn.Gui
    data := nodesGui1.Submit()

    x      := data.NodesFound
    y      := data.NodesCancer
    maxDim := data.MaxDim

    lineA := "Liczba węzłów chłonnych znalezionych w materiale tkankowym: " x ".`n"
           . "Liczba węzłów chłonnych ze stwierdzonym utkaniem raka: " y "."

    extraLine := (maxDim != "")
        ? "Maksymalny wymiar przerzutowego utkania: " maxDim " mm."
        : ""

    lineB := (y != "0" && y != 0)
        ? "[Nie stwierdzono naciekania pozatorebkowego (ENE-).][Obecne naciekanie pozatorebkowego (ENE+).]"
        : ""

    output := lineA
    if (extraLine)
        output .= "`n" extraLine
    if (lineB)
        output .= "`n" lineB

    nodesGui1.Destroy()
    Sleep(100)

    A_Clipboard := output
    ClipWait(1)
    Send("^v")

    if (y != "0" && y != 0)
        Send("+{Tab}+{Tab}")
}

GuiEscape(nodesGui1, info := "") {
    try nodesGui1.Destroy()
}

:*:xwq::
{
    ShowFragNodesGui()
}

ShowFragNodesGui() {
    local nodesGui2, btnOk, rbNeg, rbPos, txtDim, editDim

    nodesGui2 := Gui(, "Węzły pofragmentowane")
    nodesGui2.SetFont("s10", "Arial")

    ; Radio: 1 = brak raka, 2 = jest rak
    rbNeg := nodesGui2.Add("Radio", "vMode Checked", "W pofragmentowanym utkaniu chłonnym – nie stwierdzono utkania raka.")
    rbPos := nodesGui2.Add("Radio",,               "W pofragmentowanym utkaniu chłonnym – stwierdzono utkanie raka.")

    txtDim  := nodesGui2.Add("Text", "Disabled", "Maksymalny wymiar przerzutowego utkania (mm) [opcjonalnie]:")
    editDim := nodesGui2.Add("Edit", "w80 vMaxDim Disabled")

    ; przełączanie enabled/disabled po kliknięciu radia
    toggle := (*) => ToggleDim(nodesGui2, txtDim, editDim)
    rbNeg.OnEvent("Click", toggle)
    rbPos.OnEvent("Click", toggle)

    btnOk := nodesGui2.Add("Button", "Default", "OK")
    btnOk.OnEvent("Click", XWQ_Generate)

    nodesGui2.OnEvent("Escape", XWQ_GuiClose)
    nodesGui2.OnEvent("Close",  XWQ_GuiClose)

    nodesGui2.Show()

    ; ustaw stan pól na starcie
    ToggleDim(nodesGui2, txtDim, editDim)
}

ToggleDim(guiObj, txtDim, editDim) {
    data := guiObj.Submit(false) ; NoHide
    if (data.Mode = 2) {
        txtDim.Enabled  := true
        editDim.Enabled := true
    } else {
        txtDim.Enabled  := false
        editDim.Enabled := false
        editDim.Value := ""
    }
}

XWQ_Generate(btn, info) {
    nodesGui2 := btn.Gui
    data := nodesGui2.Submit()

    mode   := data.Mode          ; 1 lub 2
    maxDim := data.MaxDim

    if (mode = 1) {
        output := "W pofragmentowanym utkaniu chłonnym – nie stwierdzono utkania raka."
    } else {
        output := "W pofragmentowanym utkaniu chłonnym – stwierdzono utkanie raka."
        if (maxDim != "")
            output .= "`nMaksymalny wymiar przerzutowego utkania: " maxDim " mm."
        output .= "`n[Nie stwierdzono naciekania pozatorebkowego (ENE-).][Obecne naciekanie pozatorebkowego (ENE+).]"
    }

    nodesGui2.Destroy()
    Sleep(100)

    A_Clipboard := output
    ClipWait(1)
    Send("^v")

    if (mode = 2) {
        Sleep(120)
        Send("+{Tab}+{Tab}")
    }
}

XWQ_GuiClose(nodesGui2, info := "") {
    try nodesGui2.Destroy()
}




; =========================
;  LNC = Lymph Node Counter (basic) - AHK v2
;  Prefiksowane nazwy, żeby uniknąć kolizji po wklejeniu do większego skryptu
;  - uruchamiany hotstringiem (AHK v2 wymaga bloku {} dla kodu)
;  - domyślnie WYŁĄCZONY (GUI tworzy się dopiero po pierwszym wywołaniu)
;  - Close/Esc = Hide (bez ExitApp)
;  - Toggle bez .Visible (Gui nie ma takiej właściwości)
; =========================

:*:xlnc::
{
  Sleep 30
  LNC_ToggleGui()   
}

global LNC_State := Map(
  "found", 0,
  "pos", 0,
  "hist", [] ; array of action objects
)

global LNC_UI := Map()     ; control refs
global LNC_Gui := 0        ; lazy init
global LNC_IsShown := false

; ---------- Hotkeys (aktywne tylko gdy GUI jest pokazane) ----------
#HotIf LNC_IsShown
NumpadAdd::LNC_AddFound(1)
F23::LNC_AddFound(1)
NumpadEnter::LNC_AddPos(1, true)
F24::LNC_AddPos(1, true)
Numpad0::LNC_ResetAll()
F21::LNC_ResetAll()
NumpadDot::LNC_CopySummary()
F22::LNC_CopySummary()
NumpadSub::LNC_Undo()
#HotIf

; ---------- GUI ----------
LNC_BuildGui() {
  global LNC_UI, LNC_IsShown
  g := Gui("+AlwaysOnTop -MinimizeBox", "LN counter")
  g.SetFont("s10", "Segoe UI")

  ; Counters
  g.AddText("xm w120", "Znalezione:")
  LNC_UI["tFound"] := g.AddText("x+m w80 Right", "0")

  g.AddText("xm w120", "Dodatnie:")
  LNC_UI["tPos"] := g.AddText("x+m w80 Right", "0")

  g.AddText("xm w120", "Ujemne:")
  LNC_UI["tNeg"] := g.AddText("x+m w80 Right", "0")

  g.AddText("xm w260", " ")

  ; Buttons row 1
  btnF1 := g.AddButton("xm w70 h60", "+1 NEG")
  btnF5 := g.AddButton("x+m w70 h60", "+5 NEG")
  btnFm1 := g.AddButton("x+m w70 h60", "-1 NEG")

  ; Buttons row 2
  btnP1  := g.AddButton("xm w70", "● +1 POS")
  btnPm1 := g.AddButton("x+m w70", "● -1 POS")
  btnUndo := g.AddButton("x+m w70", "Undo")
  btnP1.Opt("+Default")  ; wizualne wyróżnienie

  ; Buttons row 3
  btnCopy := g.AddButton("xm w110", "Kopiuj")
  btnReset := g.AddButton("x+m w110", "Reset")

  ; Summary preview
  g.AddText("xm w320", "Wynik:")
  LNC_UI["eSummary"] := g.AddEdit("xm w320 r5 ReadOnly +VScroll", "")

  ; Optional extra line (only when pos>0)
  LNC_UI["cbMaxDim"] := g.AddCheckBox("xm w320", "Dodaj maksymalny wymiar przerzutowego utkania")
  LNC_UI["cbMaxDim"].Value := 0
  LNC_UI["cbMaxDim"].Enabled := false

  ; Events
  btnF1.OnEvent("Click", (*) => LNC_AddFound(1))
  btnF5.OnEvent("Click", (*) => LNC_AddFound(5))
  btnFm1.OnEvent("Click", (*) => LNC_AddFound(-1))

  btnP1.OnEvent("Click", (*) => LNC_AddPos(1, true))
  btnPm1.OnEvent("Click", (*) => LNC_AddPos(-1, false))
  btnUndo.OnEvent("Click", (*) => LNC_Undo())

  btnCopy.OnEvent("Click", (*) => LNC_CopySummary())
  btnReset.OnEvent("Click", (*) => LNC_ResetAll())
  LNC_UI["cbMaxDim"].OnEvent("Click", (*) => LNC_UpdateUI())

  g.OnEvent("Close",  (*) => (g.Hide(), LNC_IsShown := false, LNC_ResetAll()))
  g.OnEvent("Escape", (*) => (g.Hide(), LNC_IsShown := false, LNC_ResetAll()))

  LNC_UpdateUI()
  return g
}

; ---------- Actions ----------
LNC_AddFound(delta) {
  global LNC_State
  if (delta = 0)
    return

  newFound := LNC_State["found"] + delta
  if (newFound < 0)
    return

  ; clamp pos not to exceed found
  newPos := LNC_State["pos"]
  if (newPos > newFound)
    newPos := newFound

  LNC_PushHist({ kind: "found", dFound: delta, dPos: (newPos - LNC_State["pos"]) })
  LNC_State["found"] := newFound
  LNC_State["pos"] := newPos
  LNC_UpdateUI()
}

LNC_AddPos(delta, alsoFound := true) {
  global LNC_State
  if (delta = 0)
    return

  dFound := 0
  dPos := delta

  if (alsoFound && delta > 0)
    dFound := delta

  newFound := LNC_State["found"] + dFound
  newPos := LNC_State["pos"] + dPos

  if (newFound < 0 || newPos < 0)
    return

  ; pos cannot exceed found
  if (newPos > newFound)
    return

  LNC_PushHist({ kind: "pos", dFound: dFound, dPos: dPos })
  LNC_State["found"] := newFound
  LNC_State["pos"] := newPos
  LNC_UpdateUI()
}

LNC_Undo() {
  global LNC_State
  hist := LNC_State["hist"]
  if (hist.Length = 0)
    return

  act := hist.Pop()
  if (act.kind = "reset") {
    LNC_State["found"] := act.prevFound
    LNC_State["pos"] := act.prevPos
    LNC_State["hist"] := act.prevHist
    LNC_UpdateUI()
    return
  }
  LNC_State["found"] := LNC_Max(0, LNC_State["found"] - act.dFound)
  LNC_State["pos"] := LNC_Max(0, LNC_State["pos"] - act.dPos)

  ; ensure invariant pos<=found
  if (LNC_State["pos"] > LNC_State["found"])
    LNC_State["pos"] := LNC_State["found"]

  LNC_UpdateUI()
}

LNC_ResetAll() {
  global LNC_State
  prevHist := LNC_State["hist"].Clone()
  prevFound := LNC_State["found"]
  prevPos := LNC_State["pos"]
  LNC_State["found"] := 0
  LNC_State["pos"] := 0
  LNC_State["hist"] := []
  LNC_State["hist"].Push({ kind: "reset", prevFound: prevFound, prevPos: prevPos, prevHist: prevHist })
  LNC_UpdateUI()
}

LNC_CopySummary() {
  A_Clipboard := LNC_GetSummaryText()
}

LNC_ToggleGui() {
  global LNC_Gui, LNC_IsShown
  if !IsObject(LNC_Gui) {
    LNC_Gui := LNC_BuildGui()
    LNC_Gui.Show("AutoSize NA")
    WinGetPos(&x, &y, &w, &h, "ahk_id " LNC_Gui.Hwnd)
    LNC_Gui.Move(A_ScreenWidth - w - 300, 50) ; 10 px margines

    LNC_IsShown := true
    return
  }

  if (LNC_IsShown) {
    LNC_Gui.Hide()
    LNC_IsShown := false
  } else {
    LNC_Gui.Show("NA")
    LNC_IsShown := true
  }
}

; ---------- Helpers ----------
LNC_PushHist(act) {
  global LNC_State
  LNC_State["hist"].Push(act)
}

LNC_UpdateUI() {
  global LNC_State, LNC_UI
  if (LNC_UI.Count = 0)
    return
  found := LNC_State["found"]
  pos := LNC_State["pos"]
  neg := found - pos

  LNC_UI["tFound"].Text := found
  LNC_UI["tPos"].Text := pos
  LNC_UI["tNeg"].Text := neg
  LNC_UI["eSummary"].Value := LNC_GetSummaryText()
  LNC_UI["cbMaxDim"].Enabled := (pos > 0)
  if (pos = 0)
    LNC_UI["cbMaxDim"].Value := 0
}

LNC_FormatHist(a) {
  if (a.kind = "found") {
    return (a.dFound > 0) ? Format("+{} znalezione", a.dFound) : Format("{} znalezione", a.dFound)
  } else {
    s := ""
    if (a.dPos != 0)
      s .= (a.dPos > 0) ? Format("+{} dodatnie", a.dPos) : Format("{} dodatnie", a.dPos)
    if (a.dFound != 0)
      s .= Format(" (i {} znalezione)", a.dFound > 0 ? "+" a.dFound : a.dFound)
    return s
  }
}

LNC_Max(a, b) => (a > b ? a : b)

LNC_GetSummaryText() {
  global LNC_State
  found := LNC_State["found"]
  pos := LNC_State["pos"]
  if (found = 0)
    return "Tkanka tłuszczowa bez utkania chłonnego."
  text := Format(
    "Liczba węzłów chłonnych znalezionych w materiale tkankowym: {}`nLiczba węzłów chłonnych ze stwierdzonym utkaniem raka: {}",
    found,
    pos
  )
  if (pos > 0)
    text .= "`n[Nie stwierdzono naciekania pozatorebkowego (ENE-).][Obecne naciekanie pozatorebkowe (ENE+).]"
  if (pos > 0 && LNC_UI.Has("cbMaxDim") && LNC_UI["cbMaxDim"].Value)
    text .= "`nMaksymalny wymiar przerzutowego utkania: [] mm."
  return text
}

;================


:*:xangio::{#}angio{Space}
:*:xcc::{#}cc{Space}
:*:xcep::{#}cep{Space}
:*:xcrad::{#}crad{Space}
:*:xczaw::{#}czaw{Space}
:*:xapp::{#}app{Space}
:*:xvs::{#}vs{Space}
:*:xctr::{#}ctr{Space}
:*:xkam::{#}kam{Space}
:*:xfnd::{#}fnd{Space}
:*:xbaso::{#}baso{Space}
:*:xtl::{#}tl{Space}
:*:xtvl::{#}tvl{Space}
:*:xhyp::{#}hyp{Space}
:*:xgt::{#}gt{Space}
:*:xlip::{#}lip{Space}
:*:xalip::{#}alip{Space}
:*:xuropap::{#}uropap{Space}
:*:xteratoma::{#}teratoma{Space}
:*:xplazmo::{#}plazmo{Space}
:*:xnpd::{#}npd{Space}
:*:xnpc::{#}npc{Space}
:*:xgastrit::{#}gastrit{Space}
:*:wwpki::nabłonek wielowarstwowy płaski
:*:wwpgo::nabłonka wielowarstwowego płaskiego
:*:wwpim::nabłonkiem wielowarstwowym płaskim
:*:xpilo::{#}pilo{Space}

:*:xcin::{
    Sleep(50)
    ib := InputBox(
        "
        (
        1: Biopsja 
        2: Konizat
    )", "WYBÓR TYPU MATERIAŁU", "h100"  
     )
    if ib.Result != "OK"
        return

    mode := ib.Value

    m := Map(
        "1", "{#}cin_b{Space}",
        "2", "{#}cin_k{Space}",      
    )

    if m.Has(mode)
        Send(m[mode])
    else
        MsgBox("Nieprawidłowy wybór.")
}

:*:xpdl::{
    Sleep(50)
    ib := InputBox(
        "
        (
        1: NSCLC 
        2: HNSCC 
        3: TNBC 
        4: UC 
        5: szyjka macicy 
        6: czerniak 
        7: żołądek i przełyk gruczołowy 
        8: przełyk płaski 
        9: generyczny
    )",
        "PD-L1 Wybierz rodzaj nowotworu", "h240"
    )
    if ib.Result != "OK"
        return

    mode := ib.Value

    m := Map(
        "1", "{#}pdl1nsclc{Space}",
        "2", "{#}pdl1hnscc{Space}",
        "3", "{#}pdl1tnbc{Space}",
        "4", "{#}pdl1uc{Space}",
        "5", "{#}pdl1cc{Space}",
        "6", "{#}pdl1mel{Space}",
        "7", "{#}pdl1gej{Space}",
        "8", "{#}pdl1escc{Space}",
        "9", "{#}pdl1gen{Space}"
    )

    if m.Has(mode)
        Send(m[mode])
    else
        MsgBox("Nieprawidłowy wybór.")
}


:*:xtar::{
    Sleep(50)
    ib := InputBox(
        "
(
1: brodawkowaty
2: brodawkowaty wieloogniskowy
3: pęcherzykowy/onkocytarny/IEFVPTC
4: rdzeniasty
5: HG/anaplastyczny
6: inne
7: gruczolaki i niepewne
)", "Szablon nowotwory tarczycy", "h240"
     )
    if ib.Result != "OK"
        return

    mode := ib.Value

    m := Map(
        "1", "{#}tar_rakb{Space}",
        "2", "{#}tar_rakb_multi{Space}",      
        "3", "{#}tar_rakp{Space}",
        "4", "{#}tar_rakr{Space}",
        "5", "{#}tar_rakh{Space}",
        "6", "{#}tar_raki{Space}",
        "7", "{#}tar_misc{Space}"
    )

    if m.Has(mode)
        Send(m[mode])
    else
        MsgBox("Nieprawidłowy wybór.")
}

:*:xcx::{
if WinActive("PatARCH [SUKRAKOW]") 
    Send("{tab}{tab}z03+{tab}+{Tab}")

else
    Send("{tab}z03+{tab}")
}
