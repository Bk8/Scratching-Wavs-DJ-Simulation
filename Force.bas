Attribute VB_Name = "Force"
Option Explicit
 
Public Function Calculate_Normal_Force(m As Single, g As Single) As Single

   Calculate_Normal_Force = m * g

End Function

Public Function Calculate_Frictional_Force(u As Single, N As Single) As Single

    Calculate_Frictional_Force = u * N

End Function
