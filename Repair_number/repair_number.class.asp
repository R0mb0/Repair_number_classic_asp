<%
    Class repair_number

        ' Initialization and destruction'
        Sub class_initialize()
        End Sub
        
        Sub class_terminate()
        End Sub

        'Function to check if a number is an integer
        Private Function is_integer(number)
            If InStr(number, ",") <> 0 or InStr(number, ".") <> 0 Then 
                is_integer = false
            Else
            is_integer = true
            End If
        End Function

        'Function to split a number when is not possible know how the interpreter works
        Private Function my_split(number)
            If Not is_integer(number) Then
                If InStr(number, ",") <> 0 Then 
                    my_split = Split(number,",")
                    Exit Function 
                End If 
                If InStr(number, ".") <> 0 Then 
                    my_split = Split(number,".")
                    Exit Function 
                End If 
            End If 
                Call Err.Raise(vbObjectError + 10, "repair_number", "The number: " & number & " is not regular ")
        End Function

        'Function to convert a number in a array
        Private Function string_to_array(text)
            Dim length
            length = Len(text)
            Dim outArray() 
            Dim index 
            For index = 0 to length - 1
                Redim preserve outArray(length)
                outArray(index) = Left(Right(text,(length - index)), (1))
            Next 
            Redim preserve outArray(length - 1)
            string_to_array = outArray
        End Function

        ' Number must be an integer -> it returns the number inside the most recurring digit and the postion of the sequence 
        'Array structure (0) = start index (1) = end structure (2) = sample number (3) = number occurences
        Private Function analyse_number(number)
            Dim number_array
            number_array = string_to_array(number)
            Dim index 
            index = 0
            Dim temp 
            Dim array1(3)
            array1(3) = 1
            Dim array2(3)
            array2(3) = 0
            For Each temp In number_array
                If temp <> last_number Then
                    If  array1(3) >= array2(3) Then 
                        array2(0) = array1(0)
                        array2(1) = index - 1
                        array2(2) = array1(2)
                        array2(3) = array1(3)
                    End If 
                    array1(0) = index 
                    array1(2) = temp
                    array1(3) = 1
                    last_number = temp
                Else 
                    array1(3) = array1(3) + 1
                End If 
                index = index + 1
            Next 
            If array1(2) <> array2(2) and array1(3) >= array2(3) Then 
                array2(0) = array1(0)
                array2(1) = index
                array2(2) = array1(2)
                array2(3) = array1(3)
            End If 
            analyse_number = array2
        End Function 

        'Function to convert an array in a number 
        Private Function array_to_string(my_array)
            Dim temp_string 
            temp_string = ""
            Dim temp 
            For Each temp In my_array
                temp_string = temp_string & temp
            Next 
            array_to_string = temp_string
        End Function

        'Function to sum at the right position a integer number of one digit 
        Private Function Sum_number_to_decimal(number, number_to_sum)
            Dim my_number
            my_number = "0,"
            Dim temp 
            For Each temp In string_to_array(my_split(number)(1))
                my_number = my_number & "0"
            Next 
            my_number = my_number & number_to_sum
            Sum_number_to_decimal = Cdbl(number) + Cdbl(Replace(my_number, "0" & number_to_sum, number_to_sum))
        End Function

        'Function to find the inde of an element in the array 
        Private Function get_firtst_index_of_element(my_array, to_find)
            Dim index 
            index = 0 
            Dim temp 
            For Each temp In my_array
            If temp = to_find Then 
                get_firtst_index_of_element = index
                Exit Function 
            End If 
            index = index + 1
            Next 
        End Function 

        'Function to repair a number that cames from a bad operations 
        Public Function repair_number(number)
            Dim splitted_number
            splitted_number = my_split(number)
            Dim number_array
            Dim number_properties
            Dim number_to_add
            If Not is_integer(number) Then 
                If Len(splitted_number(1)) > 10 Then 
                    number_properties = analyse_number(splitted_number(1))
                    number_array = string_to_array(number)
                    Redim Preserve number_array(Len(number) - Len(splitted_number(1)) + number_properties(0))
                    'Resolve a special case: 9,9999999999999999 some times return a wrong number
                    If Abs(splitted_number(0)) = 9 and number_array(get_firtst_index_of_element(number_array, ",") + 1) = 9 and number_properties(2) = 9 Then 
                        If number_array(0) = "-" Then 
                            'if the numebr is in scientific notation
                            If InStr(LCase(number), "e-") <> 0 Then
                                repair_number = -1 / 10 ^ Int(Split(LCase(number), "e-")(1))
                                Exit Function
                            Else
                                repair_number = -1
                                Exit Function
                            End If 
                        Else
                            'if the numebr is in scientific notation
                            If InStr(LCase(number), "e-") <> 0 Then
                                repair_number = 1 / 10 ^ Int(Split(LCase(number), "e-")(1))
                                Exit Function
                            Else
                                repair_number = 1
                                Exit Function
                            End If 
                        End If 
                    End If 
                    'Resolve the case when the most part of number is 0: 8, 1200000000000003
                    If number_properties(2) = 0 Then 
                        'if the numebr is in scientific notation
                        If InStr(LCase(number), "e-") <> 0 Then
                            repair_number = array_to_string(number_array) / 10 ^ Int(Split(LCase(number), "e-")(1))
                            Exit Function
                        Else
                            repair_number = array_to_string(number_array)
                            Exit Function
                        End If 
                    End If 
                    'Resolve normal cases 
                    'if the numebr is in scientific notation
                    If InStr(LCase(number), "e-") <> 0 Then 
                        repair_number = (Sum_number_to_decimal(array_to_string(number_array), 10 - number_properties(2))) / 10 ^ Int(Split(LCase(number), "e-")(1))
                        Exit Function
                    Else
                        repair_number = Sum_number_to_decimal(array_to_string(number_array), 10 - number_properties(2))
                        Exit Function
                    End If
                Else
                    repair_number = number
                    Exit Function 
                End If 
            Else 
                repair_number = number
                Exit Function 
            End If
        End Function
    End Class 
%>