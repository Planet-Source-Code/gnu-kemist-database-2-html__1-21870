Attribute VB_Name = "OpenDatabase"
'*************************************************
'* Project: Database 2 HTML                                                    *
'* Programmer: Gnu Kemist GnuKemist@yahoo.com                *
'* Version: 0.0.1 (as of March. 22, 2001)                                   *
'* Known Bugs: None                                                                *
'*************************************************
Global cnnHockey As New ADODB.Connection
Global cmdGrid As New ADODB.Command
Global rstGrid As New ADODB.Recordset

Public Sub Constructor()
'# Opes a connection to the Hockey Database                          #
'# The DataSource section MUST match the path to the database on        #
'# your own computer (i.e. Data Source= [Correct Path] & "\Hockey.MDB    #

cnnHockey.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & " Data Source=" & App.Path & "\Hockey.MDB;Persist Security Info=False"
cnnHockey.Open

End Sub

Public Sub Destructor()
'# Closes the object when done #
cnnHockey.Close
End Sub
