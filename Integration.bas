Attribute VB_Name = "Integration"
Option Explicit

'void simulation_world::Integrate( real DeltaTime )
'{
'    int Counter;

'    for(Counter = 0;Counter < NumberOfBodies;Counter++)
'    {
'        rigid_body::configuration &Source =
'            aBodies[Counter].aConfigurations[SourceConfigurationIndex];
'        rigid_body::configuration &Target =
'            aBodies[Counter].aConfigurations[TargetConfigurationIndex];

'        // integrate primary quantities

'        Target.CMPosition = Source.CMPosition +
'                DeltaTime * Source.CMVelocity;

'        Target.Orientation = Source.Orientation +
'                DeltaTime *
'                matrix_3x3(Source.AngularVelocity,matrix_3x3::SkewSymmetric) *
'                Source.Orientation;

'        Target.CMVelocity = Source.CMVelocity +
'                (DeltaTime * aBodies[Counter].OneOverMass) * Source.CMForce;

'        Target.AngularMomentum = Source.AngularMomentum +
'                DeltaTime * Source.Torque;

'        OrthonormalizeOrientation(Target.Orientation);

'        // compute auxiliary quantities

'        Target.InverseWorldInertiaTensor = Target.Orientation *
'                aBodies[Counter].InverseBodyInertiaTensor *
'                Transpose(Target.Orientation);

'        Target.AngularVelocity = Target.InverseWorldInertiaTensor *
'                Target.AngularMomentum;
'    }
'}

Public Enum CONST_INTEGRATOR

    FORWARD_EULER = 0
    SECOND_ORDER_EULER = 1
    VERLET = 2
    VELOCITY_VERLET = 3
    SECOND_ORDER_RUNGE_KUTTA = 4
    THIRD_ORDER_RUNGE_KUTTA = 5
    FORTH_ORDER_RUNGE_KUTTA = 6
    
End Enum

Private Old_Position As POINT2D
Private Old_Velocity As POINT2D
Private Old_Acceleration As POINT2D
Private Old_Angle As Single
Private Old_Angular_Velocity As Single
Private Old_Angular_Acceleration As Single

Public Sub Integrate2D(Obj As PHYSICS2D, dt As Single, Integrator As CONST_INTEGRATOR)

    Dim k1 As POINT2D, k2 As POINT2D, k3 As POINT2D, k4 As POINT2D
    Dim l1 As VECTOR2D, l2 As VECTOR2D, l3 As VECTOR2D, l4 As VECTOR2D
    
    Dim m1 As Single, m2 As Single, m3 As Single, m4 As Single
    Dim n1 As Single, n2 As Single, n3 As Single, n4 As Single
    
    With Obj

        .Acceleration.X = .Force.Net.X * .One_Over_Mass
        .Acceleration.Y = .Force.Net.Y * .One_Over_Mass
        
        .Angular_Acceleration = .Torque.Net * .One_Over_Inertia
        
        Select Case Integrator
        
            Case FORWARD_EULER
        
                'Target.CMPosition = Source.CMPosition + Source.CMVelocity * DeltaTime;
                'Target.CMVelocity = Source.CMVelocity + (aBodies[Counter].OneOverMass * Source.CMForce) * DeltaTime;
                'Target.AngularMomentum = Source.AngularMomentum + Source.Torque * DeltaTime;
                'Target.InverseWorldInertiaTensor = Target.Orientation * aBodies[Counter].InverseBodyInertiaTensor * Transpose(Target.Orientation);
                'Target.AngularVelocity = Target.InverseWorldInertiaTensor * Target.AngularMomentum;
                
                .Position.X = .Position.X + .Velocity.X * dt
                .Velocity.X = .Velocity.X + .Acceleration.X * dt
            
                .Position.Y = .Position.Y + .Velocity.Y * dt
                .Velocity.Y = .Velocity.Y + .Acceleration.Y * dt
                
                .Angle = .Angle + .Angular_Velocity * dt
                .Angular_Velocity = .Angular_Velocity + .Angular_Acceleration * dt
                
            Case SECOND_ORDER_EULER
            
                .Position.X = .Position.X + .Velocity.X * dt + 0.5 * .Acceleration.X * dt * dt
                .Velocity.X = .Velocity.X + .Acceleration.X * dt
            
                .Position.Y = .Position.Y + .Velocity.Y * dt + 0.5 * .Acceleration.Y * dt * dt
                .Velocity.Y = .Velocity.Y + .Acceleration.Y * dt
                
                .Angle = .Angle + .Angular_Velocity * dt + 0.5 * .Angular_Acceleration * dt * dt
                .Angular_Velocity = .Angular_Velocity + .Angular_Acceleration * dt
            
            Case VERLET
            
                .Velocity.X = .Position.X - Old_Position.X + .Acceleration.X * dt * dt
                Old_Position.X = .Position.X
                .Position.X = .Position.X + .Velocity.X
                
                .Velocity.Y = .Position.Y - Old_Position.Y + .Acceleration.Y * dt * dt
                Old_Position.Y = .Position.Y
                .Position.Y = .Position.Y + .Velocity.Y
                
                .Angular_Velocity = .Angle - Old_Angle + .Angular_Acceleration * dt * dt
                Old_Angle = .Angle
                .Angle = .Angle + .Angular_Velocity
                
            Case VELOCITY_VERLET
                
                Old_Acceleration.X = .Acceleration.X
                .Position.X = .Position.X + .Velocity.X * dt + 0.5 * Old_Acceleration.X * dt * dt
                .Velocity.X = .Velocity.X + 0.5 * (Old_Acceleration.X + .Acceleration.X) * dt
            
                Old_Acceleration.Y = .Acceleration.Y
                .Position.Y = .Position.Y + .Velocity.Y * dt + 0.5 * Old_Acceleration.Y * dt * dt
                .Velocity.Y = .Velocity.Y + 0.5 * (Old_Acceleration.Y + .Acceleration.Y) * dt
            
                Old_Angular_Acceleration = .Angular_Acceleration
                .Angle = .Angle + .Angular_Velocity * dt + 0.5 + Old_Angular_Acceleration * dt * dt
                .Angular_Velocity = .Angular_Velocity + 0.5 * (Old_Angular_Acceleration + .Angular_Acceleration) * dt
            
            Case SECOND_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity + m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                .Position.X = .Position.X + k2.X
                .Position.Y = .Position.Y + k2.Y
                .Velocity.X = .Velocity.X + l2.X
                .Velocity.Y = .Velocity.Y + l2.Y
                
                
                .Angle = .Angle + m2
                .Angular_Velocity = .Angular_Velocity + n2
            
            Case THIRD_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                k3.X = dt * (.Velocity.X - k1.X + 2 * k2.X)
                k3.Y = dt * (.Velocity.Y - k1.Y + 2 * k2.Y)
                l3.X = dt * .Acceleration.X
                l3.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity * m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                m3 = dt * (.Angular_Velocity - m1 + 2 * m2)
                n3 = dt * .Angular_Acceleration
                
                .Position.X = .Position.X + k1.X * 1 / 6 + k2.X * 2 / 3 + k3.X * 1 / 6
                .Position.Y = .Position.Y + k1.Y * 1 / 6 + k2.Y * 2 / 3 + k3.Y * 1 / 6
                .Velocity.X = .Velocity.X + l1.X * 1 / 6 + l2.X * 2 / 3 + l3.X * 1 / 6
                .Velocity.Y = .Velocity.Y + l1.Y * 1 / 6 + l2.Y * 2 / 3 + l3.Y * 1 / 6
                
                .Angle = .Angle + m1 * 1 / 6 + m2 * 2 / 3 + m3 * 1 / 6
                .Angular_Velocity = .Angular_Velocity + n1 * 1 / 6 + n2 * 2 / 3 + n3 * 1 / 6
                
            Case FORTH_ORDER_RUNGE_KUTTA
            
                k1.X = dt * .Velocity.X
                k1.Y = dt * .Velocity.Y
                l1.X = dt * .Acceleration.X
                l1.Y = dt * .Acceleration.Y
                
                k2.X = dt * (.Velocity.X + k1.X / 2)
                k2.Y = dt * (.Velocity.Y + k1.Y / 2)
                l2.X = dt * .Acceleration.X
                l2.Y = dt * .Acceleration.Y
                
                k3.X = dt * (.Velocity.X + k2.X / 2)
                k3.Y = dt * (.Velocity.Y + k2.Y / 2)
                l3.X = dt * .Acceleration.X
                l3.Y = dt * .Acceleration.Y
                
                k4.X = dt * (.Velocity.X + k3.X)
                k4.Y = dt * (.Velocity.Y + k3.Y)
                l4.X = dt * .Acceleration.X
                l4.Y = dt * .Acceleration.Y
                
                m1 = dt * .Angular_Velocity
                n1 = dt * .Angular_Acceleration
                
                m2 = dt * (.Angular_Velocity + m1 / 2)
                n2 = dt * .Angular_Acceleration
                
                m3 = dt * (.Angular_Velocity + m2 / 2)
                n3 = dt * .Angular_Acceleration
                
                m4 = dt * (.Angular_Velocity + m3)
                n4 = dt * .Angular_Acceleration
                 
                .Position.X = .Position.X + k1.X / 6 + k2.X / 3 + k3.X / 3 + k4.X / 6
                .Position.Y = .Position.Y + k1.Y / 6 + k2.Y / 3 + k3.Y / 3 + k4.Y / 6
                .Velocity.X = .Velocity.X + l1.X / 6 + l2.X / 3 + l3.X / 3 + l4.X / 6
                .Velocity.Y = .Velocity.Y + l1.Y / 6 + l2.Y / 3 + l3.Y / 3 + l4.Y / 6
                
                .Angle = .Angle + m1 / 6 + m2 / 3 + m3 / 3 + m4 / 6
                .Angular_Velocity = .Angular_Velocity + n1 / 6 + n2 / 3 + n3 / 3 + n4 / 6
                
        End Select
        
    End With

End Sub


