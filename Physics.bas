Attribute VB_Name = "Physics"
Option Explicit

Public Type POINT2D

    X As Single
    Y As Single

End Type

Public Type VECTOR2D

    X As Single
    Y As Single

End Type

Public Type FORCE2D

    Net As VECTOR2D
    Gravity As VECTOR2D
    Friction As VECTOR2D
    Spring As VECTOR2D
    Applied As VECTOR2D
    Magnetic As VECTOR2D
    Wind As VECTOR2D
    Drag As VECTOR2D

End Type

Public Type TORQUE2D

    Net As Single
    Gravity As Single
    Friction As Single
    Spring As Single
    Applied As Single
    Magnetic As Single
    Wind As Single
    Drag As Single

End Type

Public Type PHYSICS2D
    
    Position As POINT2D
    Velocity As VECTOR2D
    Acceleration As VECTOR2D
    Angle As Single
    Angular_Velocity As Single
    Angular_Acceleration As Single
    Force As FORCE2D
    Torque As TORQUE2D
    Mass As Single
    One_Over_Mass As Single
    Inertia As Single
    One_Over_Inertia As Single
    Elasticity As Single
    Momentum As VECTOR2D
    Impulse As VECTOR2D
    Angular_Momentum As VECTOR2D
    Angular_Impulse As VECTOR2D
    
End Type

Public Const CENTIMETERS_TO_INCH As Single = 0.393700787

Public Const POUNDS_TO_KG As Single = 0.45359237

Public Const ONE_KGF_CM_TO_NEWTONS_CM As Single = 9.80665
Public Const TWO_KGF_CM_TO_NEWTONS_CM As Single = 9.80665 * 2
Public Const THREE_KGF_CM_TO_NEWTONS_CM As Single = 9.80665 * 3
Public Const FOUR_KGF_CM_TO_NEWTONS_CM As Single = 9.80665 * 4

Public Const EARTH_GRAVITY As Single = 9.80665
Public Const AIR_DENSITY As Single = 1.29

