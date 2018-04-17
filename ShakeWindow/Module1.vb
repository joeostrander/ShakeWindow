Imports System.Drawing
Imports System.Text
Imports System.Runtime.InteropServices


Module Module1


    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean

    Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Int32, ByRef lpRect As RECT) As Boolean
    Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As UInteger)

    Const HWND_TOPMOST As Integer = -1
    Const HWND_NOTOPMOST As Integer = -2

    Const SWP_NOSIZE As Integer = 1

    Public Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure 'RECT

    Sub Main()
        Dim strWindowTitle As String = ""
        Dim strWindowClass As String = vbNullString

        Dim args As Array = Environment.GetCommandLineArgs

        If args Is Nothing Or args.Length < 2 Then
            WriteUsage()
        Else
            strWindowTitle = args(1)
            If args.Length = 3 Then strWindowClass = args(2)
        End If


        Dim handle As IntPtr = FindWindow(strWindowClass, strWindowTitle)
        If Not IsWindow(handle) Then
            Console.WriteLine("Could not find window titled:  " & strWindowTitle)
            Exit Sub
        End If

        shakeMe(handle)

    End Sub

    Public Sub shakeMe(ByRef hWnd As IntPtr)
        Dim rect As New RECT
        GetWindowRect(hWnd.ToInt32, rect)

        Dim myLoc As Point = New Point(rect.left, rect.top)
        Dim myLocStart As Point = New Point(rect.left, rect.top)

        For i As Integer = 0 To 10
            For x As Integer = 0 To 4
                Select Case x
                    Case 0
                        myLoc.X = myLocStart.X + 10
                    Case 1
                        myLoc.X = myLocStart.X - 10
                    Case 2
                        myLoc.Y = myLocStart.Y - 10
                    Case 3
                        myLoc.Y = myLocStart.Y + 10
                    Case 4
                        myLoc = myLocStart
                End Select
                SetWindowPos(hWnd, HWND_TOPMOST, myLoc.X, myLoc.Y, 0, 0, SWP_NOSIZE)
                System.Threading.Thread.Sleep(15)
            Next
        Next
        SetWindowPos(hWnd, HWND_NOTOPMOST, myLocStart.X, myLocStart.Y, 0, 0, SWP_NOSIZE)

    End Sub

    Private Sub WriteUsage()
        Console.WriteLine()
        Console.WriteLine("Description:  Shake a window")
        Console.WriteLine("Author:       Joe Ostrander")
        Console.WriteLine("Date:         2014.06.16")
        Console.WriteLine()
        Console.WriteLine("USAGE:")
        Console.WriteLine("ShakeWindow.exe ""Your Window Title""")
        Console.WriteLine("OR")
        Console.WriteLine("ShakeWindow.exe ""Your Window Title"" ""window_class_name""")
        Console.WriteLine()
        Console.WriteLine("Example:")
        Console.WriteLine("ShakeWindow.exe ""Conversations"" ""wcl_manager1""")
        Environment.Exit(0)
    End Sub

End Module
