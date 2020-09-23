Attribute VB_Name = "InitialMod"
Public fMainForm As frmMain

'Variable holder for the database path
Public sDatabasePath As String

'variable holder for the universal connectionstring
'used through out the example.
Global sConnString As String



Sub Main()
    sConnString = GetSetting(App.title, "Database", "ConnString", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\nauction.mdb" & ";Jet OLEDB")
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub

