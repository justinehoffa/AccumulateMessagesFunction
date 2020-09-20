'Justine Hoffa
'RCET0265
'Fall2020
'AccumulationMessagesFunction
'https://github.com/justinehoffa/AccumulateMessagesFunction

Option Strict On
Option Explicit On
Option Compare Text

Module AccumulateMessagesFunction

    Sub Main()
        Dim userInput As String
        Dim message As String
        Dim clearData As Boolean

        Do
            userInput = Console.ReadLine()
            If userInput = "call" Then
                MsgBox(message)
            ElseIf userInput = "clear" Then
                clearData = True
            End If

            message = AccumulateMessage(userInput, clearData)
            clearData = False

        Loop

    End Sub

    Function AccumulateMessage(ByVal newMessage As String, ByVal clear As Boolean) As String
        Static userMessage As String
        If clear = True Then
            userMessage = ""
        ElseIf newMessage = "call" Then
        Else
            userMessage &= newMessage & vbNewLine
        End If

        Return userMessage
    End Function

End Module
