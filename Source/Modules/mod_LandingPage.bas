Attribute VB_Name = "mod_LandingPage"
Option Explicit

' ============================================================================
'  mod_LandingPage - handlers dla Form Controls buttons na sh_LandingPage.
'
'  DLACZEGO standalone module (nie Sheet code module):
'  Form Controls (Insert -> Form Controls -> Button) mogą wołać macra TYLKO
'  z standalone modules (via Assign Macro dialog). Nie mogą wołać bezpośrednio
'  z Sheet code module. Dlatego handlers są tutaj jako Public Subs.
'
'  DLACZEGO Form Controls (nie ActiveX):
'  Production robustness - ActiveX moze byc zablokowany przez Trust Center
'  policy IT (typowo w korporacyjnych deploymentach), nie dziala na Mac /
'  Excel Online. Form Controls dzialaja WSZEDZIE - Windows/Mac/Web, kazda
'  polityka bezpieczenstwa. 20+ handlowcow = 20+ potencjalnych IT policies =
'  ryzyko zero-day deployment blockers. Form Controls to zero risk.
'
'  Trade-off: Form Controls wygladaja starzej (bardziej Excel 2003), maja
'  mniej properties (brak per-button Font control). Ale dla launcher pattern
'  (5 buttons z tekstem) jest to akceptowalne - dashboard to landing page,
'  nie decorative UI.
'
'  Guard pattern (mod_Utils.IsFormOpen): defensywny check przeciwko
'  double-instance przy cascade otwarciach.
'
'  Post-form refresh (2026-08-16): kazdy handler po Show vbModal (blocking)
'  wywoluje sh_LandingPage.RefreshDashboard. Powod: Worksheet_Activate NIE
'  fire'uje po Me.Hide modal formu, bo sheet byl juz aktywny (form tylko go
'  przykrywal, nie zmienial aktywnego sheeta). Bez explicit refresh dashboard
'  pokazuje stary user info po SwitchUser (bug 2026-08-16).
'
'  Patrz: Source/Sheets/sh_LandingPage.LAYOUT.md
'         Source/Sheets/sh_LandingPage.code.txt (Worksheet_Activate + Refresh)
' ============================================================================

' ----- Button handlers (assigned via Excel right-click -> Assign Macro) ----

Public Sub OpenMain()
    If mod_Utils.IsFormOpen("frm_Main") Then Exit Sub
    frm_Main.Show vbModal
    sh_LandingPage.RefreshDashboard
End Sub

Public Sub OpenLog()
    If mod_Utils.IsFormOpen("frm_Log") Then Exit Sub
    frm_Log.Show vbModal
    sh_LandingPage.RefreshDashboard
End Sub

Public Sub OpenPicker()
    If mod_Utils.IsFormOpen("frm_UserPicker") Then Exit Sub
    frm_UserPicker.Show vbModal
    sh_LandingPage.RefreshDashboard
End Sub

Public Sub OpenSetup()
    If mod_Utils.IsFormOpen("frm_Setup") Then Exit Sub
    frm_Setup.Show vbModal
    sh_LandingPage.RefreshDashboard
End Sub

Public Sub OpenTutorial()
    If mod_Utils.IsFormOpen("frm_Tutorial") Then Exit Sub
    frm_Tutorial.Show vbModal
    sh_LandingPage.RefreshDashboard
End Sub
