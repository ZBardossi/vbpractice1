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
        Dim a, b, exp

        a = InputBox("Enter a value of a")
        b = InputBox("Enter a value of b")
        exp = InputBox("Add a operation: 1: Additon, 2: Subtraction, 3: Multiplication")



    End Sub
End Module
