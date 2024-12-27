Attribute VB_Name = "Manifest"
Option Explicit
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
 
Public Sub InitManifest()
    'Iniciamos permisos para ejecutar como administrador
    InitCommonControls
End Sub
 
