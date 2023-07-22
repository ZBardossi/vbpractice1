Module Module1

    Sub Main()
        'Dim a, b, c

        'a = 10
        'b = "Hello"
        'c = 3.14
        'MsgBox(a)
        'MsgBox(VarType(b))
        'MsgBox(TypeName(c))
        'areaOfCircle()
        'evenOrOdd()
        'moreOrLessThenZero()
        sampleSelectCase()
        'loop1()
        'loop3()
        'forLoop()
        'whileWendLoop()







    End Sub

    Sub areaOfCircle()
        Dim aoc, r
        Const pi = 3.14

        r = InputBox("Add a radius of circle")
        aoc = r * r * pi
        MsgBox("Area of circle is: " & aoc)

    End Sub

    Sub evenOrOdd()
        Dim a

        a = InputBox("Enter a value")
        If a Mod 2 = 0 Then
            MsgBox("even")
        Else
            MsgBox("odd")
        End If
    End Sub

    Sub moreOrLessThenZero()
        Dim a
        a = InputBox("Enter a number")

        If a < 0 Then
            MsgBox("It is below zero")
        ElseIf a > 0 Then
            MsgBox("It is above zero")
        ElseIf a = 0 Then
            MsgBox("you entered zero")

        End If
    End Sub

    Sub sampleSelectCase()
        On Error Resume Next

        Dim a, b, expr

        a = CInt(InputBox("Enter a value of a"))
        b = CInt(InputBox("Enter a value of b"))
        expr = InputBox("Add a operation: 1: Additon, 2: Subtraction, 3: Multiplication")

        If Err.Number <> 0 Then
            MsgBox(“Number of the Error and Description is “ & Err.Number & ” ” & Err.Description) '(Give details about the Error)
            Err.Clear() '(Will Clear the Error)
        End If

        Select Case expr
            Case 1
                MsgBox("Addition: " & (a + b))
            Case 2
                MsgBox("Subtraction: " & (a - b))
            Case 3
                MsgBox("Multiplication: " & (a * b))
            Case Else
                MsgBox("Wrong choice")
        End Select

    End Sub

    Sub loop1()
        Dim myCount
        myCount = 0


        Console.WriteLine("Counting to ten with a Do loop")

        Do While myCount < 10
            myCount = myCount + 1
            Console.WriteLine(myCount)
            Threading.Thread.Sleep(1000) 'Slowing down the console
        Loop

    End Sub

    Sub loop2()
        Dim counter
        counter = 0

        Do
            Console.WriteLine(counter)
            counter = counter + 1
            Threading.Thread.Sleep(1000)
        Loop While counter < 1
    End Sub

    Sub loop3()
        Dim myCount
        myCount = 0


        Console.WriteLine("Counting to ten with a Do Until loop")

        Do Until myCount > 9
            myCount = myCount + 1
            Console.WriteLine(myCount)
            Threading.Thread.Sleep(1000) 'Slowing down the console
        Loop

    End Sub

    Sub forLoop()
        Dim counter

        For counter = 1 To 10
            Console.WriteLine(counter)
            Threading.Thread.Sleep(500)

        Next
    End Sub


    While counter < 11
            Console.WriteLine(counter)
            counter = counter + 1
            Threading.Thread.Sleep(500)
        End While


    End Sub
End Module
