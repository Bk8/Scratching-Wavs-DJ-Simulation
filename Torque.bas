Attribute VB_Name = "Torque"
Option Explicit

Public Sub Reset_Torque(Obj As PHYSICS2D)

    Obj.Torque.Net = 0

End Sub

Public Sub Add_Torque_2D(Obj As PHYSICS2D, Torque As Single)

    Obj.Torque.Net = Obj.Torque.Net + Torque

End Sub

Public Sub Calculate_Torque_2Da(Obj As PHYSICS2D, Force As Single, Radius As Single)

    Obj.Torque.Net = Force * Radius

End Sub

Public Sub Calculate_Torque_2Db(Obj As PHYSICS2D, Inertia As Single, Angular_Acceleration As Single)

    Obj.Torque.Net = Inertia * Angular_Acceleration

End Sub
