SetKeyDelay, -1


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.

;#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance Force
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^2:: ;#mat lepszy
IfWinActive, PatARCH [SUKRAKOW]
  {
  Send +3mat
  Sleep 100
  Send {enter}
  Sleep 100
  Send ^{Home}
  InputBox, repeat, liczba materiałów
  Sleep 200
  Loop, %repeat%  {
      Send, {Del}{Del}{Del}{end}
      Sleep 20
      Send, {enter}[]{enter}{Down}{Home}
    }
  Send ^{home}
  return
  }
else
  Send +3mat
  Sleep 100
  Send {enter}
  InputBox, repeat, liczba materiałów
  Send ^{Home}
  Loop, %repeat%
    {
    Send, {end}:
    Sleep 100
    Send {enter}[]{enter}{Down}
    }
  Send ^{home}
return

^3:: ; przenoszenie rozpoznania i st. zaawansowania z szablonu do pola podsumowania, kursor w pozycji 0

    ClipSaved := ClipboardAll
    Clipboard := ""
    Send ^c
    ClipWait, 0.5
    if (ErrorLevel) {
        SoundBeep, 1500, 120
        Tooltip, Brak zaznaczenia / schowek pusty.
        SetTimer, __ToolOff, -1200
        Clipboard := ClipSaved
        return
    }

    full := Clipboard

    ; Podział: wszystko przed frazą "Stopień zaawansowania" zostaje na dole,
    ; a CAŁOŚĆ ląduje w polu wyżej.
    ; i) = case-insensitive, (?s) = kropka łapie nowe linie
    pat := "i)(?s)\bstopie(?:ń|n)\s+zaawansowania\b"
    pos := RegExMatch(full, pat)

    if (!pos) {
        ; Bez markera nie wiemy gdzie uciąć pierwsze zdanie -> nie kombinujemy.
        SoundBeep, 1200, 150
        Tooltip, Nie znalazłem „Stopień zaawansowania…”. Operacja przerwana.
        SetTimer, __ToolOff, -1600
        Clipboard := ClipSaved
        return
    }

    first := Trim(SubStr(full, 1, pos-1), " `t`r`n")

    ; 1) Zastąp zaznaczenie tylko pierwszym zdaniem (to ma zostać w tym polu)
    Clipboard := first
    ClipWait, 0.5
    Send ^v
    Sleep, 40

    ; 2) Przejdź do pola wyżej i wklej CAŁY oryginalny tekst
    Send +{Tab}
    Sleep, 40
    Clipboard := full
    ClipWait, 0.5
    Send ^v
    Sleep, 40

    ; Finisz
    Clipboard := ClipSaved
    VarSetCapacity(ClipSaved, 0)
return

__ToolOff:
Tooltip
return

^4:: ; lista gotowców
MsgBox,

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
xplazmo - chronic endometritis z CD138
xpdl - skrypt na PD-L1 (wybierz numer żeby wybrać nowotwór)
xtbz - fragment/y tkanki tłuszczowej/włóknistej bez istotnych zmian patologicznych
xsll - CLL/SLL
xnschl - Hodgkin NS
xcx - Z03
wwp(ki,go,im) nabłonek wielowarstwowy płaski (mianownik, dopełniacz, narzędnik)
)

return


^5:: ; ponowne wyświetlenie instrukcji
MsgBox,

(

CTRL+1: dodawanie korelacji z poziomu zapisanego rozpoznania (działa w Chrome)
CTRL+2: lepszy #mat
CTRL+3: przenoszenie rozpoznania i st. zaawansowania z szablonu (dwie pierwsze linie) do pola podsumowania (kursor w pozycji 0)
CTRL+4: lista skrótów z kodami ICD
CTRL+5: ponowne wyświetlenie tego okna
CTRL+6: otwiera bluebooks w edge
CRTL+7: otwiera CAP w edge
CTRL+8: nowe zakładki do kontrasygnaty
CTRL+9: rezydenci - wybór specjalisty do kontrasygnaty
CTRL+Q: rezydenci - skrót do przekazania wyniku do kontrasygnaty przez specjalistę (kursor w polu rozpoznania)
ALT+Z: specjaliści - dołącz do wyniku jako Diagnozujący 2 (działa w Chrome)

)

return

^6::
Run microsoft-edge:https://tumourclassification.iarc.who.int/home
return

^7::
Run microsoft-edge:https://www.cap.org/protocols-and-guidelines/cancer-reporting-tools/cancer-protocol-templates
return

^9:: ; wybór specjalisty do kontrasygnaty
InputBox, name, wybór specjalisty w Patarch, wpisz trzyliterowy skrót nazwy specjalisty, ,300, 200
MsgBox, wybrany specjalista: %name%
return


^q:: ; skrót do kontrasygnaty przez specjalistę, kursor w oknie rozpoznania

;SendInput {tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{space}{tab}%name%{tab}{space}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}

Send, ^f
Sleep 50
Send, Kontras
Sleep 20
Send, {esc}{tab}{space}{tab}%name%{tab}{space}
return

return


:*:xadeno::
dgn :="
(
Rak niedrobnokomórkowy.
Immunofenotyp komórek nowotworu: TTF1+=/p40-.
Zgodnie z terminologią WHO dla drobnych biopsji - obraz raka niedrobnokomórkowego, w pierwszym rzędzie gruczołowego (+'non-small cell carcinoma, favour adenocarcinoma+').
[Materiał reprezentatywny do badań predykcyjnych.]
)"
ICDO := "80463"
ICD10 := "C34"
ICDfill()
return

:*:xplano::
dgn :="
(
Rak niedrobnokomórkowy.
Immunofenotyp komórek nowotworu: TTF1-/p40+=.
Zgodnie z terminologią WHO dla drobnych biopsji - obraz raka niedrobnokomórkowego, w pierwszym rzędzie płaskonabłonkowego (+'non-small cell carcinoma, favour squamous cell carcinoma+').
[Materiał reprezentatywny do badań predykcyjnych.]
)"
ICDO := "80463"
ICD10 := "C34"
ICDfill()
return

:*:xnos::
dgn :="
(
Rak niedrobnokomórkowy.
Immunofenotyp komórek nowotworu: [TTF1-/p40-][TTF1+=/p40+=].
Zgodnie z terminologią WHO dla drobnych biopsji - obraz raka niedrobnokomórkowego, inaczej nieokreślonego (non-small cell carcinoma NOS).
[Komentarz: Obecna ekspresja obu markerów immunohistochemicznych w utkaniu nowotworu - co włącza w zakres diagnostyki różnicowej raka gruczołowo-płaskonabłonkowego (adenosquamous carcinoma), niemniej - to rozpoznanie nie może być postawione na podstawie materiału biopsyjnego.]
[Wskazana dokładna ocena kliniczna celem wykluczenia przerzutowego charakteru zmiany.]
[Materiał reprezentatywny do badań predykcyjnych.]
)"
ICDO := "80463"
ICD10 := "C34"
ICDfill()
return

:*:xpred::
Send Materiał z wyżej określonego bloczka parafinowego wybrano i zakwalifikowano do oceny ekspresji czynników predykcyjnych: EGFR, ALK, ROS1, PD-L1. Materiał zawiera ponad 100 komórek raka.
return

:*:xnpd::
dgn :="
(
Znamię barwnikowe śródskórne (dermal nevus).
Linie cięcia operacyjnego wolne od utkania zmiany.
)"
ICDO := "87200"
ICD10 := "D22"
ICDfill()
return

:*:xnpc::
dgn :="
(
Znamię barwnikowe złożone (compound nevus).
Linie cięcia operacyjnego wolne od utkania zmiany.
)"
ICDO := "87200"
ICD10 := "D22"
ICDfill()
return

:*:xnpj::
dgn :="
(
Znamię barwnikowe łączące (junctional nevus).
Linie cięcia operacyjnego wolne od utkania zmiany.
)"
ICDO := "87200"
ICD10 := "D22"
ICDfill()
return

:*:xctr::
dgn :="
(
Torbiel włosowa (trichilemmal cyst).
)"
ICDO :=
ICD10 := "L72.1"
ICDfill()
return

:*:xelster::
dgn :="
(
Polip z gruczołów trawieńcowych (Elstera).
)"
ICDO :=
ICD10 := "K31.7"
ICDfill()
return

:*:xgastrit::
dgn :="
(
Przewlekłe zapalenie błony śluzowej [trzonu][części przedodźwiernikowej] żołądka.
Cechy ilościowe zapalenia (w skali czterostopniowej: -, +=, +=+=, +=+=+=):
- nasilenie []
- aktywność []
- zanik []
- metaplazja jelitowa []
- H. pylori []
)"
ICDO :=
ICD10 := "K29"
ICDfill()
return

:*:xsarko::
dgn :="
(
Elementy morfotyczne utkania chłonnego z obecnością nabłonkowatokomórkowych ziarniniaków nieulegających martwicy. W klinicznej diagnostyce różnicowej można uwzględnić sarkoidozę.
Nie stwierdzono cytomorfologicznych cech procesu nowotworowego.
)"
ICDO :=
ICD10 := "I88"
ICDfill()
return

:*:xteratoma::
dgn :="
(
Potworniak dojrzały jajnika (WHO: mature teratoma of the ovary).
)"
ICDO := "90800"
ICD10 := "D27"
ICDfill()
return

:*:xcep::
dgn :="
(
Torbiel naskórkowa
)"
ICDO :=
ICD10 := "L72.0"
ICDfill()
return

:*:xfnd::
Send, +3fnd{space}
/*dgn :="
(
Rozrost guzkowy tarczycy (WHO: thyroid follicular nodular disease).
[przytarczyca, węzły chłonne]
)"
ICDO :=
ICD10 := "E07.8"
ICDfill()
*/
return

:*:xbaso::
/*dgn :="
(
Rak podstawnokomórkowy (basal cell carcinoma).
Typ [guzkowy][powierzchowny][naciekający].
Nie stwierdzono angio- i neuroinwazji.
Linie cięcia operacyjnego wolne od utkania zmiany. Minimalny margines operacyjny ([boczny][głęboki]): []mm.
)"
ICDO := "80903"
ICD10 := "C44"
ICDfill()
*/
Send, +3baso{Space}
return

:*:xtl::
dgn :="
(
Gruczolak cewkowy z dysplazją małego stopnia (tubular adenoma with low grade dysplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82110"
ICD10 := "D12"
ICDfill()
return

:*:xth::
dgn :="
(
Gruczolak cewkowy z dysplazją dużego stopnia (tubular adenoma with high grade dysplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82112"
ICD10 := "D12"
ICDfill()
return

:*:xtvl::
dgn :="
(
Gruczolak cewkowo-kosmkowy z dysplazją małego stopnia (tubulovillous adenoma with low grade dysplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82630"
ICD10 := "D12"
ICDfill()
return

:*:xtvh::
dgn :="
(
Gruczolak cewkowo-kosmkowy z dysplazją dużego stopnia (tubulovillous adenoma with high grade dysplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82632"
ICD10 := "D12"
ICDfill()
return


:*:xvs::
dgn :="
(
Brodawka łojotokowa (seborrhoeic keratosis).
Linie cięcia operacyjnego wolne od utkania zmiany.
)"
ICDO :=
ICD10 := "L82"
ICDfill()
return

:*:xcc::
dgn :="
(
Przewlekłe zapalenie pęcherzyka żółciowego.
)"
ICDO :=
ICD10 := "K81.1"
ICDfill()
return

:*:xcrad::
dgn :="
(
Przy zgodnych danych klinicznych obraz histologiczny odpowiada torbieli korzeniowej.
)"
ICDO :=
ICD10 := "K04.8"
ICDfill()
return

:*:xczaw::
dgn :="
(
Przy zgodnych danych klinicznych obraz histologiczny odpowiada torbieli zawiązkowej.
)"
ICDO :=
ICD10 := "K09.0"
ICDfill()
return

:*:xapp::
dgn :="
(
Ropowicze zapalenie wyrostka robaczkowego.
)"
ICDO :=
ICD10 := "K35"
ICDfill()
return


:*:xhyp::
dgn :="
(
Polip hiperplastyczny.
)"
ICDO :=
ICD10 := "K63.5"
ICDfill()
return

:*:xgt::
dgn :="
(
Naczyniak krwionośny włośniczkowy zrazikowy (lobular capillary hemangioma).
Linie cięcia operacyjnego wolne od utkania zmiany.
)"
ICDO := "91310"
ICD10 := "D23"
ICDfill()
return

:*:xlip::
dgn :="
(
Tłuszczak.
[Zmiana wyłuszczona.]
)"
ICDO := "88500"
ICD10 := "D17"
ICDfill()
return

:*:xalip::
dgn :="
(
Naczyniakotłuszczak.
[Zmiana wyłuszczona].
)"
ICDO := "88610"
ICD10 := "D17"
ICDfill()
return

:*:xuropap::
dgn := "
(
Nieinwazyjny brodawkowaty rak urotelialny, [low][high]-grade wg WHO/ISUP (non-invasive papillary urothelial carcinoma).
[Materiał nie obejmuje mięśniówki właściwej.][Materiał obejmuje mięśniówkę właściwą.]
)"
ICDO := "81302"
ICD10 := "D09.0"
ICDfill()
return

:*:xkam::
dgn := "
(
Pęcherzyk żółciowy o strukturze histologicznej w granicach normy.
)"
ICDO :=
ICD10 := "K80.2"
ICDfill()
return

:*:xemuc::
dgn :="
(
Elementy morfotyczne utkania chłonnego[.][ ze złogami brunatnego pyłu.]
Nie stwierdzono cytomorfologicznych cech procesu nowotworowego.
)"
ICDO :=
ICD10 :=
ICDfill()
return

:*:xtbz::
Send, Fragment[y] tkanki [tłuszczowej][tłuszczowo-włóknistej][włóknistej] bez istotnych zmian patologicznych.
return

:*:xsll::
dgn := "
(
Przekroje węzła chłonnego o strukturze zatartej przez naciek złożony dominująco z małych komórek B oraz rozproszonych większych, typu prolimfocyta/paraimmunoblasta, liczniejszych w wyodrębniających się w utkaniu centrach proliferacji.
Immunofenotyp komórek: CD3-/CD20+=/CD5+=/CD23+=/cyklinaD1-/CD43+=. Aktywność proliferacyjna zmienna, przeciętnie niska - ok. 5`% komórek nacieku Ki67+=, z podwyższeniem frakcji w obszarach odpowiadających centrom proliferacji - do ok. 10-15`% komórek Ki67+= w tej lokalizacji.
----
Obraz histologiczny odpowiada nieagresywnemu chłoniakowi z obwodowych limfocytów B, wg WHO: chronic lymphocytic leukaemia/small lymphocytic lymphoma (CLL/SLL).
)"
ICDO := "98233"
ICD10 := "C91.1"
ICDfill()
return

:*:xnschl::

dgn := "
(
Węzeł chłonny o strukturze zatartej, z przebudową guzkową wtórną do włóknienia. Skład komórkowy utkania polimorficzny, złożony z odczynowych, małych limfocytów T CD3+= i B CD20+=, histiocytów, eozynofilów, plazmocytów oraz rozproszonych komórek HRS, o immunofenotypie CD3-/CD20-/CD30+=/CD15+=/PAX5+=(ekspresja słabsza niż w odczynowych limfocytach B).
----
Obraz histologiczny odpowiada klasycznemu chłoniakowi Hodgkina, wg WHO: nodular sclerosis classic Hodgkin lymphoma.
)"
ICDO := "96633"
ICD10 := "C81.1"
ICDfill()
return



:*:xcin::
Sleep 50
InputBox, mode, , 1: biopsja 2: konizat
If mode=1
    Send {#}cin_b{enter}
Else if mode=2
    Send {#}cin_k{enter}

return


:*:xpromise::
dgn := "
(
Wynik badania obecności mutacji POLE w komórkach nowotworowych: [ujemny][dodatni].

[Nie stwierdzono utraty ekspresji białek odpowiadających za naprawę DNA (MSH6, PMS2) w komórkach nowotworowych, co nie daje podstaw dla stwierdzenia kancerogenezy drogą niestabilności mikrosatelitarnej. Status MMR: +'proficient+'.]
[Stwierdzono utratę ekspresji białek odpowiadających za naprawę DNA (utrata ][, brak utraty ][) w komórkach nowotworowych, co wskazuje na kancerogenezę drogą niestabilności mikrosatelitarnej. Status MMR: '+'deficient+'.]

[Nie stwierdzono patologicznej ekspresji białka p53 w komórkach nowotworowych.][Stwierdzono patologiczną ekspresję białka p53 w komórkach nowotworowych.]

----
Podtyp molekularny raka (zgodnie z klasyfikatorem ProMisE): [POLE-ultramutated (POLE EDM)][MMR-deficient (MMR-D)][p53 abnormal (p53 abn)][p53 wild-type (p53 wt)].
)"
ICDO := ""
ICD10 := "Z03"
ICDfill()
return

:*:xplazmo:: ; Define hotstring "xplazmo"
    ; Create GUI
	Gui, Add, Text, , ICD10 CE:
    Gui, Add, Button, w80 gSelectOption Default, Tak
    Gui, Add, Button, w80 gSelectOption, Nie
    Gui, Show, , Option Selector

    return

SelectOption:
    Gui, Submit, NoHide
    GuiControlGet, SelectedOption, , % A_GuiControl
    If (SelectedOption = "Tak") {
        SelectedValue := "N71.1"
    } Else If (SelectedOption = "Nie") {
        SelectedValue := "Z03"
    }
	Gui, Destroy ; Close the GUI
    dgn :="
(
[Częściowo mikropolipowato uformowane fragmenty][Fragmenty] błony śluzowej trzonu macicy w [proliferacyjnej fazie cyklu][sekrecyjnej fazie cyklu]. Podścielisko [nieco obrzękłe][, lokalnie z cechami wrzecionowatego przekształcenia], z obecnością[ ogniskowych nacieków limfocytarnych i] nieregularnie rozproszonych plazmocytów CD138+= (do [...] plazmocytów / 1 mm²).
Obraz [przemawia za rozpoznaniem przewlekłego zapalenia endometrium.][odpowiada przewlekłemu zapaleniu endometrium.]
)"
ICDO :=
ICD10 := % SelectedValue
ICDfill()
return



:*:xpdl::
Sleep 50
/*

NSCLC :=
(
"Status immunohistochemicznej ekspresji PD-L1:
[TPS: <1%. Ekspresja PD-L1 w poniżej 1% komórek nowotworu.]
[TPS: 1-49%. Ekspresja PD-L1 w około ][% komórek nowotworu.]
[TPS: ≥50%. Ekspresja PD-L1 w powyżej 50% komórek nowotworu.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

HNSCC :=
(
"[Status immunohistochemicznej ekspresji PD-L1: pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥1. Wartość CPS: ][.]
[Status immunohistochemicznej ekspresji PD-L1: negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <1.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

TNBC :=
(
"[Status immunohistochemicznej ekspresji PD-L1: pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥10.]
[Status immunohistochemicznej ekspresji PD-L1: negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <10.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

UC :=
(
"Status immunohistochemicznej ekspresji PD-L1:
[Negatywny. Ekspresja PD-L1 w poniżej 1% komórek nowotworu (TPS: <1%).]
[Pozytywny. Ekspresja PD-L1 w powyżej 1% komórek nowotworu (TPS: ≥1%). Wartość TPS: ][%.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)


CSCC :=
(
"[Status immunohistochemicznej ekspresji PD-L1: pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥1. Wartość CPS: ][.]
[Status immunohistochemicznej ekspresji PD-L1: negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <1.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

Melanoma :=
(
"Status immunohistochemicznej ekspresji PD-L1:
[Negatywny. Ekspresja PD-L1 w poniżej 1% komórek nowotworu (TPS: <1%).]
[Pozytywny. Ekspresja PD-L1 w powyżej 1% komórek nowotworu (TPS: ≥1%). Wartość TPS: ][%.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

GEJ :=
(
"[Status immunohistochemicznej ekspresji PD-L1: pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥10.]
[Status immunohistochemicznej ekspresji PD-L1: pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥5.]
[Status immunohistochemicznej ekspresji PD-L1: negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <5.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

ESCC :=
(
"Status immunohistochemicznej ekspresji PD-L1:
Odsetek komórek nowotworowych z ekspresją PD-L1 (TPS): ok. []`%.
Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): [].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

gen :=
(
"Status immunohistochemicznej ekspresji PD-L1:
Odsetek komórek nowotworowych z ekspresją PD-L1 (TPS): ok. []`%.
Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): [].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)
*/

/*

NSCLC :=
(
"[TPS: <1`%. Ekspresja PD-L1 w poniżej 1`% komórek nowotworu.]
[TPS: 1-49`%. Ekspresja PD-L1 w około ][`% komórek nowotworu.]
[TPS: ≥50`%. Ekspresja PD-L1 w powyżej 50`% komórek nowotworu.]
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

HNSCC :=
(
"[Wynik pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥1.]
[Wynik negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <1.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

TNBC :=
(
"[Wynik pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥10.]
[Wynik negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <10.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

UC :=
(
"[Wynik negatywny (TPS: <1`%). Ekspresja PD-L1 w poniżej 1`% komórek nowotworu.]
[Wynik pozytywny (TPS: ≥1`%). Ekspresja PD-L1 w powyżej 1`% komórek nowotworu.]
[Wartość TPS: ][]`%.
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)


CSCC :=
(
"[Wynik pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥1.]
[Wynik negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <1.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

Melanoma :=
(
"[Wynik negatywny (TPS: <1`%). Ekspresja PD-L1 w poniżej 1`% komórek nowotworu.]
[Wynik pozytywny (TPS: ≥1`%). Ekspresja PD-L1 w powyżej 1`% komórek nowotworu.]
[Wartość TPS: ][]`%.
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

GEJ :=
(
"[Wynik pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥10.]
[Wynik pozytywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): ≥5.]
[Wynik negatywny. Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): <5.]
[Wartość CPS: ][].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

ESCC :=
(
"Odsetek komórek nowotworowych z ekspresją PD-L1 (TPS): ok. []`%.
Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): [].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)

gen :=
(
"Odsetek komórek nowotworowych z ekspresją PD-L1 (TPS): ok. []`%.
Ekspresja PD-L1 oceniana jako +'combined positive score+' (CPS): [].
----
Badanie wykonano na materiale tkankowym z bloczka parafinowego nr []."
)
*/



/*


InputBox, mode, , 1: NSCLC 2: HNSCC 3:TNBC 4: UC 5: szyjka macicy 6: czerniak 7: żołądek i przełyk gruczołowy 8 przełyk płaski 9: generyczny
If mode=1
    Send % NSCLC
Else if mode=2
    Send % HNSCC
Else if mode=3
    Send % TNBC
Else if mode=4
    Send % UC
Else if mode=5
    Send % CSCC
Else if mode=6
    Send % Melanoma
Else if mode=7
    Send % GEJ
Else if mode=8
    Send % ESCC
Else if mode=9
    Send % gen

Send {tab}{tab}Z03+{tab}+{tab}^{home}
return
*/
InputBox, mode, , 1: NSCLC 2: HNSCC 3:TNBC 4: UC 5: szyjka macicy 6: czerniak 7: żołądek i przełyk gruczołowy 8 przełyk płaski 9: generyczny
If mode=1
    Send {#}pdl1nsclc{enter}
Else if mode=2
    Send {#}pdl1hnscc{enter}
Else if mode=3
    Send {#}pdl1tnbc{enter}
Else if mode=4
    Send {#}pdl1uc{enter}
Else if mode=5
    Send {#}pdl1cc{enter}
Else if mode=6
    Send {#}pdl1mel{enter}
Else if mode=7
    Send {#}pdl1gej{enter}
Else if mode=8
    Send {#}pdl1escc{enter}
Else if mode=9
    Send {#}pdl1gen{enter}

return


:*:xww::
Gui, Destroy
Gui, Font, s10, Arial  ; Ustawienie czcionki: 10-punkt Arial.
Gui, Add, Text,, Podaj liczbę węzłów chłonnych znalezionych w preparacie:
Gui, Add, Edit, w30 h20 vNodesFound,
Gui, Add, Text,, Podaj liczbę węzłów chłonnych ze stwierdzonym utkaniem raka:
Gui, Add, Edit, w30 h20 vNodesCancer, 0
Gui, Add, Text,, Podaj maksymalny wymiar przerzutowego utkania (mm) [opcjonalnie]:
Gui, Add, Edit, w30 h20 vMaxDim,
Gui, Add, Button, Default gGenerate, Ok
Gui, Show,, Węzły chłonne
return

Generate:
    ; Pobranie danych z pól GUI
    Gui, Submit, NoHide
    x := NodesFound
    y := NodesCancer

    ; Utworzenie linijki A
    lineA := "Liczba węzłów chłonnych znalezionych w materiale tkankowym: " x ".`nLiczba węzłów chłonnych ze stwierdzonym utkaniem raka: " y "."

    ; Utworzenie dodatkowej linii tylko, gdy pole maksymalnego wymiaru nie jest puste
    if (MaxDim != "")
        extraLine := "Maksymalny wymiar przerzutowego utkania: " MaxDim " mm."
    else
        extraLine := ""

    ; Utworzenie linijki B tylko, gdy y jest różne od 0
    if (y != "0" && y != 0)
        lineB := "[Nie stwierdzono naciekania pozatorebkowego (ENE-).][Obecne naciekanie pozatorebkowe (ENE+).]"
    else
        lineB := ""

    ; Łączenie wyniku – extraLine przed linią B
    output := lineA
    if (extraLine != "")
        output .= "`n" extraLine
    if (lineB != "")
        output .= "`n" lineB

    ; Zamknięcie GUI
       Gui, Destroy
    Sleep, 100

    ; Skopiowanie wyniku do schowka
    Clipboard := output
    ClipWait, 1

    ; Wklejenie zawartości schowka w aktywne pole (Ctrl+V)
    Send, ^v
    if (y != "0" && y != 0)
        Send, +{tab}+{tab}
    else
        return
return

GuiEscape:
    Gui, Destroy
return


:*:xwq::
Gui, Destroy
Gui, New, +AlwaysOnTop, Węzły pofragmentowane
Gui, +LabelXWQ_        ; wszystkie zdarzenia tego GUI trafią do etykiet XWQ_Gui*
Gui, Font, s10, Arial

; Radio: 1 = brak raka, 2 = stwierdzono raka
Gui, Add, Radio, vMode gXWQ_Toggle Checked, W pofragmentowanym utkaniu chłonnym – nie stwierdzono utkania raka.
Gui, Add, Radio,        gXWQ_Toggle,         W pofragmentowanym utkaniu chłonnym – stwierdzono utkanie raka.

Gui, Add, Text, vLblDim Disabled, Maksymalny wymiar przerzutowego utkania (mm) [opcjonalnie]:
Gui, Add, Edit, vMaxDim w120 Disabled

Gui, Add, Button, Default w110 gXWQ_Generate, OK
Gui, Show,, Węzły pofragmentowane
return

XWQ_Toggle:
    Gui, Submit, NoHide
    ; Mode = 1 lub 2 (bo zmienna jest przy pierwszym Radio)
    if (Mode = 2) {
        GuiControl, Enable, LblDim
        GuiControl, Enable, MaxDim
    } else {
        GuiControl, Disable, LblDim
        GuiControl, Disable, MaxDim
        GuiControl,, MaxDim
    }
return

; Wklejanie przez schowek z bezpiecznym przywróceniem
XWQ_PasteViaClipboard(text) {
    RestoreDelayMs := 220
    ClipSaved := ClipboardAll
    Clipboard := ""
    Sleep, 25
    Clipboard := text
    ClipWait, 1
    if (ErrorLevel) {
        Clipboard := text
        ClipWait, 1
    }
    SendInput, ^v
    Sleep, %RestoreDelayMs%
    Clipboard := ClipSaved
    VarSetCapacity(ClipSaved, 0)
}

XWQ_Generate:
    Gui, Submit, NoHide
    if (Mode = 1) {
        output := "W pofragmentowanym utkaniu chłonnym – nie stwierdzono utkania raka."
    } else if (Mode = 2) {
        output := "W pofragmentowanym utkaniu chłonnym – stwierdzono utkanie raka."
        if (MaxDim != "")
            output .= "`nMaksymalny wymiar przerzutowego utkania: " MaxDim " mm."
        output .= "`n[Nie stwierdzono naciekania pozatorebkowego (ENE-).][Obecne naciekanie pozatorebkowego (ENE+).]"
    } else {
        MsgBox, 48, Błąd, Wybierz jedną z opcji.
        return
    }

    Gui, Destroy
    Sleep, 80
    XWQ_PasteViaClipboard(output)

    if (Mode = 2) {
        Sleep, 120
        Send, +{Tab}+{Tab}
    }
return

; Zamknięcia przypięte przez +LabelXWQ_
XWQ_GuiEscape:
XWQ_GuiClose:
    Gui, Destroy
return



:*:xcx:: ;to co w funkcji tylko bez ^{home}
dgn :=
ICDO :=
ICD10 := "Z03"
Send %dgn%
IfWinActive, PatARCH [SUKRAKOW]
   {
       Send, {tab}%ICDO%{tab}%ICD10%+{tab}+{tab}
    return

}
else
   {
       Send, {tab}%ICD10%{tab}%ICDO%+{tab}+{tab}
    return

}

return

:*:xangio::
angiofill()
Send {Up}{Home}
return

--------------------------------
:*:wwpki::
Send nabłonek wielowarstwowy płaski
return

:*:wwpgo::
Send nabłonka wielowarstwowego płaskiego
return

:*:wwpim::
Send nabłonkiem wielowarstwowym płaskim
return


--------------------------------

ICDfill(){
global
Send %dgn%
IfWinActive, PatARCH [SUKRAKOW]
   {
       Send, {tab}%ICDO%{tab}%ICD10%+{tab}+{tab}^{Home}
    return

}
else
   {
       Send, {tab}%ICD10%{tab}%ICDO%+{tab}+{tab}^{Home}
    return

}
}
return


angiofill(){ ;funkcja angio- i neuroinwazja
Send [Nie stwierdzono cech inwazji naczyń (LVI0).][Obecne cechy angioinwazji (LVI1).]{enter}[Nie stwierdzono cech inwazji okołonerwowej (PNI0).][Obecne cechy inwazji okołonerwowej (PNI1).]
}

