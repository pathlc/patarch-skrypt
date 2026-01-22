#Requires AutoHotkey v2.0
#SingleInstance Force

; =========================
;  LNC = Lymph Node Counter (basic) - AHK v2
;  Prefiksowane nazwy, żeby uniknąć kolizji po wklejeniu do większego skryptu
; =========================

global LNC_State := Map(
  "found", 0,
  "pos", 0,
  "hist", [] ; array of action objects
)

global LNC_UI := Map()     ; control refs
global LNC_Gui := LNC_BuildGui()
LNC_Gui.Show("AutoSize")

; ---------- Hotkeys ----------
NumpadAdd::LNC_AddFound(1)
NumpadSub::LNC_AddFound(-1)
NumpadEnter::LNC_AddPos(1, true)

#HotIf WinActive("LN counter")
Backspace::LNC_Undo()
#HotIf

; ---------- GUI ----------
LNC_BuildGui() {
  global LNC_UI
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

  g.OnEvent("Close", (*) => ExitApp())

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
  LNC_State["found"] := LNC_Max(0, LNC_State["found"] - act.dFound)
  LNC_State["pos"] := LNC_Max(0, LNC_State["pos"] - act.dPos)

  ; ensure invariant pos<=found
  if (LNC_State["pos"] > LNC_State["found"])
    LNC_State["pos"] := LNC_State["found"]

  LNC_UpdateUI()
}

LNC_ResetAll() {
  global LNC_State
  LNC_State["found"] := 0
  LNC_State["pos"] := 0
  LNC_State["hist"] := []
  LNC_UpdateUI()
}

LNC_CopySummary() {
  A_Clipboard := LNC_GetSummaryText()
}

LNC_ToggleGui() {
  global LNC_Gui
  if !IsSet(LNC_Gui)
    return
  if (LNC_Gui.Visible)
    LNC_Gui.Hide()
  else
    LNC_Gui.Show()
}

; ---------- Helpers ----------
LNC_PushHist(act) {
  global LNC_State
  LNC_State["hist"].Push(act)
}

LNC_UpdateUI() {
  global LNC_State, LNC_UI
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
