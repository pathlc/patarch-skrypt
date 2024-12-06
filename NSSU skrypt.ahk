SetKeyDelay, -1



#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance Force
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^p::
ExitApp
return

F1:: ;otwieranie skierowania
IfWinActive, PatARCH [SUKRAKOW] - Google Chrome
{
Send ^f
Sleep 100
Send status
Sleep 100
Send {enter}
Sleep 100
Send {Esc}
Sleep 200
Send +{tab}
Sleep 200
Send {enter}
MouseMove, (A_ScreenWidth // 2), (A_ScreenHeight // 2), 0
Sleep 20
Click
Sleep 100
Send {tab}{tab}{Enter}
MouseMove, (A_ScreenWidth // 2), (A_ScreenHeight // 2), 0
Sleep 20
Click
Send {Esc}
}
return


^1::
IfWinActive, PatARCH [SUKRAKOW] - Google Chrome
{
Send ^f
Sleep 100
Send ami    [
Sleep 100
Send {enter}
Sleep 100
Send {Esc}
Sleep 50
Send {enter}
Sleep 300
MouseMove, (A_ScreenWidth // 2), ((A_ScreenHeight // 5)), 0
Sleep 50
Click
Sleep 50
Send {tab}
}
return


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

^Numpad3::
Send,  {Space}/ {del}{end}
return


;^Numpad2:: ;#mat lepszy KSS
;Send +3mat
;Sleep 100
;Send {enter}
;InputBox, repeat, liczba materiałów
;Send ^{Home}
;Loop, %repeat%
 ; {
 ; Send, {end}:
 ; Sleep 50
 ; Send {enter}[]{enter}{Down}
 ; }
;Send ^{home}
;return



^3:: ; przenoszenie rozpoznania i st. zaawansowania z szablonu do pola podsumowania, kursor w pozycji 0
SendInput, +{down}+{down}^x{del}+{tab}^v{Backspace}{tab}
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
xctr - torbiel tricholemmalna
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
xcin(od 0-3) - dysplazja szyjki macicy, lub jej brak (0)
xplazmo - chronic endometritis z CD138
xpdl - skrypt na PD-L1 (wybierz numer żeby wybrać nowotwór)
xtbz - fragment/y tkanki tłuszczowej/włóknistej bez istotnych zmian patologicznych
xsll - CLL/SLL
xnschl - Hodgkin NS
xcx - Z03
wwp(m,d,n) nabłonek wielowarstwowy płaski (mianownik, dopełniacz, narzędnik)
)

return


^5:: ; ponowne wyświetlenie instrukcji
MsgBox,

(
F1: otwieranie skierowania
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
SendInput {tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{tab}{space}{tab}%name%{tab}{space}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}+{tab}
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
Torbiel tricholemmalna.
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
dgn :="
(
Rozrost guzkowy tarczycy (WHO: thyroid follicular nodular disease).
[przytarczyca, węzły chłonne]
)"
ICDO :=
ICD10 := "E07.8"
ICDfill()
return

:*:xbaso::
dgn :="
(
Rak podstawnokomórkowy (basal cell carcinoma).
Typ [guzkowy][powierzchowny][naciekający].
Nie stwierdzono angio- i neuroinwazji.
Linie cięcia operacyjnego wolne od utkania zmiany. Minimalny margines operacyjny ([boczny][głęboki]): []mm.
)"
ICDO := "80903"
ICD10 := "C44"
ICDfill()
return

:*:xtl::
dgn :="
(
Gruczolak cewkowy z dysplazją małego stopnia (tubular adenoma with low grade intraepithelial neoplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82110"
ICD10 := "D12"
ICDfill()
return

:*:xtvl::
dgn :="
(
Gruczolak cewkowo-kosmkowy z dysplazją małego stopnia (tubulovillous adenoma with low grade intraepithelial neoplasia).
[Linie cięcia operacyjnego wolne od utkania zmiany.][Brak możliwości oceny marginesów usunięcia zmiany.]
)"
ICDO := "82630"
ICD10 := "D12"
ICDfill()
return

:*:xvs::
dgn :="
(
Brodawka łojotokowa (seborrheic keratosis).
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
Obraz histologiczny odpowiada indolentnemu chłoniakowi z obwodowych limfocytów B, wg WHO: chronic lymphocytic leukaemia/small lymphocytic lymphoma (CLL/SLL).
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



:*:xcin0::
dgn := "
(
Fragmenty błony śluzowej okolicy ujścia zewnętrznego kanału szyjki macicy (obejmujące strefę T) bez cech dysplazji.
)"
ICDO := ""
ICD10 := "z03"
ICDfill()
return

:*:xcin1::
dgn := "
(
Fragmenty błony śluzowej okolicy ujścia zewnętrznego kanału szyjki macicy (obejmujące strefę T) z dysplazją niskiego stopnia (LSIL/CIN1).
)"
ICDO := "80770"
ICD10 := "N87"
ICDfill()
return

:*:xcin2::
dgn := "
(
Fragmenty błony śluzowej okolicy ujścia zewnętrznego kanału szyjki macicy (obejmujące strefę T) z dysplazją średniego stopnia - HSIL (CIN 2).
)"
ICDO := "80772"
ICD10 := "N87"
ICDfill()
return

:*:xcin3::
dgn := "
(
Fragmenty błony śluzowej okolicy ujścia zewnętrznego kanału szyjki macicy (obejmujące strefę T) z dysplazją dużego stopnia - HSIL (CIN 3).
)"
ICDO := "80772"
ICD10 := "N87"
ICDfill()
return

:*:xpromise::
dgn := "
(
Wynik badania obecności mutacji POLE w komórkach nowotworowych: [ujemny][dodatni].

[Nie stwierdzono utraty ekspresji białek odpowiadających za naprawę DNA (MLH1, MSH2, MSH6, PMS2) w komórkach nowotworowych, co nie daje podstaw dla stwierdzenia kancerogenezy drogą niestabilności mikrosatelitarnej. Status MMR: +'proficient+'.]
[Stwierdzono utratę ekspresji białek odpowiadających za naprawę DNA (utrata: MLH1, PMS2, brak utraty MSH2, MSH6) w komórkach nowotworowych, co wskazuje na kancerogenezę drogą niestabilności mikrosatelitarnej. Status MMR: '+'deficient+'.]

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
"[Wynik negatywny (TPS: <1`%). Ekspresja PD-L1 w poniżej 1`% komórek nowotworu.]
[Wynik pozytywny (TPS: ≥1`%). Ekspresja PD-L1 w powyżej 1`% komórek nowotworu.]
[Wartość TPS: ][]`%.
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

InputBox, mode, , 1: NSCLC 2: HNSCC 3:TNBC 4: UC 5: szyjka macicy 6: czerniak 7: żołądek 8 przełyk płaski 9: generyczny
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

Send {tab}{tab}Z03+{tab}+{tab}
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
:*:wwpm::
Send nabłonek wielowarstwowy płaski
return

:*:wwpd::
Send nabłonka wielowarstwowego płaskiego
return

:*:wwpn::
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
