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
xpdl - skrypt na PD-L1 (wybierz numer żeby wybrać nowotwór)
xtbz - fragment/y tkanki tłuszczowej/włóknistej bez istotnych zmian patologicznych
xsll - CLL/SLL
xnschl - Hodgkin NS
xcx - Z03
wwp(ki,go,im) nabłonek wielowarstwowy płaski (mianownik, dopełniacz, narzędnik)
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
    Sleep(50)
    Send("Kontras")
    Sleep(20)
    Send("{Esc}{Tab}{Space}{Tab}" name "{Tab}{Space}")
}

:*:xpdl::{
    Sleep(50)
    ib := InputBox(
        "1: NSCLC 2: HNSCC 3: TNBC 4: UC 5: szyjka macicy 6: czerniak 7: żołądek i przełyk gruczołowy 8: przełyk płaski 9: generyczny",
        "PD-L1"
    )
    if ib.Result != "OK"
        return

    mode := ib.Value

    m := Map(
        "1", "{#}pdl1nsclc{Enter}",
        "2", "{#}pdl1hnscc{Enter}",
        "3", "{#}pdl1tnbc{Enter}",
        "4", "{#}pdl1uc{Enter}",
        "5", "{#}pdl1cc{Enter}",
        "6", "{#}pdl1mel{Enter}",
        "7", "{#}pdl1gej{Enter}",
        "8", "{#}pdl1escc{Enter}",
        "9", "{#}pdl1gen{Enter}"
    )

    if m.Has(mode)
        Send(m[mode])
    else
        MsgBox("Nieprawidłowy wybór.")
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



:*:xangio::{#}angio{Enter}
:*:xcc::{#}cc{Enter}
:*:xcep::{#}cep{Enter}
:*:xcrad::{#}crad{Enter}
:*:xczaw::{#}czaw{Enter}
:*:xapp::{#}app{Enter}
:*:xvs::{#}vs{Enter}
:*:xctr::{#}ctr{Enter}
:*:xkam::{#}kam{Enter}
:*:xfnd::{#}fnd{Enter}
:*:xbaso::{#}baso{Enter}
:*:xtl::{#}tl{Enter}
:*:xtvl::{#}tvl{Enter}
:*:xhyp::{#}hyp{Enter}
:*:xgt::{#}gt{Enter}
:*:xlip::{#}lip{Enter}
:*:xalip::{#}alip{Enter}
:*:xuropap::{#}uropap{Enter}
:*:xteratoma::{#}teratoma{Enter}
:*:xplazmo::{#}plazmo{Enter}
:*:xnpd::{#}npd{Enter}
:*:xnpc::{#}npc{Enter}
:*:wwpki::nabłonek wielowarstwowy płaski
:*:wwpgo::nabłonka wielowarstwowego płaskiego
:*:wwpim::nabłonkiem wielowarstwowym płaskim
:*:xpilo::{#}pilo{Space}

:*:xcin::{
    Sleep(50)
    ib := InputBox(
        "1: Biopsja 2: Konizat", "WYBÓR TYPU MATERIAŁU"
     )
    if ib.Result != "OK"
        return

    mode := ib.Value

    m := Map(
        "1", "{#}cin_b{enter}",
        "2", "{#}cin_k{enter}",      
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
