Attribute VB_Name = "mod_LandingPage"
Option Explicit

' ============================================================================
'  mod_LandingPage
'  Handlery dla Form Controls buttonów na sh_LandingPage.
'
'  Form Controls muszą wołać macra ze standalone modules (Assign Macro dialog
'  nie widzi Sheet code modules) - dlatego handlery tutaj, nie w sh_LandingPage.
'
'  Form Controls (NIE ActiveX) dla production robustness: Trust Center policies
'  IT mogą blokować ActiveX; Form Controls działają wszędzie (Windows/Mac/Web).
'
'  Post-form refresh (2026-08-16): każdy handler po Show vbModal woła
'  RefreshDashboard. Worksheet_Activate NIE fire'uje po Me.Hide modal formu
'  (sheet był już aktywny, form tylko go przykrywał).
' ============================================================================

' ----- Button handlers (przypisane w Excel przez Assign Macro) -------------

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
