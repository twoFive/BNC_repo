Attribute VB_Name = "mod_LandingPage"
Option Explicit

' ============================================================================
'  mod_LandingPage
'  Handlery dla Form Controls buttonów na sh_LandingPage.
'
'  Form Controls musz¹ wo³aæ macra ze standalone modules (Assign Macro dialog
'  nie widzi Sheet code modules) - dlatego handlery tutaj, nie w sh_LandingPage.
'
'  Form Controls (NIE ActiveX) dla production robustness: Trust Center policies
'  IT mog¹ blokowaæ ActiveX; Form Controls dzia³aj¹ wszêdzie (Windows/Mac/Web).
'
'  Post-form refresh (2026-08-16): ka¿dy handler po Show vbModal wo³a
'  RefreshDashboard. Worksheet_Activate NIE fire'uje po Me.Hide modal formu
'  (sheet by³ ju¿ aktywny, form tylko go przykrywa³).
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
