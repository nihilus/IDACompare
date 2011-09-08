'each match function is named Match_1 -> Match_4
'
'they all take the same prototype where 
'     c is the class object of CFuntion from list A
'     h is the class object of CFuntion from list B
'     identifier is a string you return on successful match
'     if you set a return value of true then match will be made.
'
' Class CFunction looks like this
'
'	 Public Length As Long
'	 Public Calls As Long
'	 Public Name As String
'	 Public Jumps As Long
'	 Public Pushs As Long
'	 Public esp As Long
'	 Public fxCalls As New Collection
'	 Public Constants As New Collection
'
'	 Function ConstantExists(key) As Boolean


Function isWithin(cnt, v1, v2, min) 
    
    If v1 <= min Or v2 <= min Then Exit Function
    
    If v1 = v2 Then
        isWithin = True
        Exit Function
    End If
    
    If v1 < v2 then low = v1 else low = v2
    
    high = v1
    If low = v1 Then high = v2
    
    If low + cnt >= high Then isWithin = True
    
End Function



function Match_1(c,h,identifier)

      If c.Calls = h.Calls And c.Pushs = h.Pushs Then  'same num of calls and pushs
	    If isWithin(60, c.Length, h.Length, 80) Then     'and length is close
                   If isWithin(4, c.Jumps, h.Jumps,1) Then    'and num jmps is close
			identifier = "Call/Push Match"
			Match_1 = true
                   End If
             End If
      End If

end function


function Match_2(c,h,identifier)

      If isWithin(80, c.Length, h.Length, 80) Then
           If c.esp <> 0 And c.esp = h.esp And isWithin(40, c.Length, h.Length,20) Then
		identifier = "ESP Match"
		Match_2 = true
           End If
      End If

end function


function Match_3(c,h,identifier)

	If h.fxCalls.Count = c.fxCalls.Count And h.fxCalls.Count > 0 Then

               j = 0
               i = 0

               For Each t In h.fxCalls
                   i = i + 1
                   If t = c.fxCalls(i) Then
                        j = j + 1
                   End If
               Next

               If j = h.fxCalls.Count Then
			identifier = "API Profile Match"
			Match_3 = true	
               End If
          
         End If

end function


function Match_4(c,h,identifier)

	If isWithin(3, c.Constants.Count, h.Constants.Count,1) And isWithin(60, c.Length, h.Length,30) Then

              j = 0
              For Each x In c.Constants
                   If h.ConstantExists(x) Then j = j + 1
              Next
                            
              If isWithin(3, c.Constants.Count, j, 1) Then
			identifier = "Const Match"
			Match_4 = true
              End If
         
     End If

end function

