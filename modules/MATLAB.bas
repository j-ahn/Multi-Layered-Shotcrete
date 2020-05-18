Attribute VB_Name = "MATLAB"
Dim MCLUtil As Object
Dim bModuleInitialized As Boolean
Dim Class1 As Object

Public beam_h, beam_b, fc_28, s_comp, steel_fy, steel_E1, rein_type, inc_tensile, tens_model, fr_1, fr_4, f1_5, etu, kh, e1c, e2c, einc, sections, sections_points
Public maxcurv, mincurv, curvinc, comp_model, alpha_cc, Eco_fib, alpha_E_fib, rho_AS3600

Dim fp_full(), ecp(), ecu(), E0(), Xi_u_CEB(), N_CEB(), n_AS3600(), k_pl(), fcmi(), fr_1_layers(), fr_4_layers(), fl_5_layers()
Dim Layers(), Ages(), rein_d, A_rebar(), Depths()

Public numberOfStages As Integer
Public num_layers As Integer, num_reo As Integer
Public r, q


Sub Master_Calcs()
Attribute Master_Calcs.VB_ProcData.VB_Invoke_Func = "m\n14"

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer
  
numberOfStages = Range("Shotcrete").Columns.Count

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim ws As Worksheet
Set ws = Worksheets("RS2_Staging")

Dim ShotcretePropertiesTable As Range: Set ShotcretePropertiesTable = Range("RS2_Staging!$B$71:$I$80")

' Pulling constants from spreadsheet
beam_b = ws.Range("beam_b")
fc_28 = ws.Range("fc_28")
s_comp = ws.Range("s_comp")
steel_fy = ws.Range("fy")
steel_E1 = ws.Range("Esteel")

rein_type = ws.Range("rein_type")

inc_tensile = ws.Range("inc_tensile")
tens_model = ws.Range("tens_model")
fr_1 = ws.Range("fr_1")
fr_4 = ws.Range("fr_4")
f1_5 = ws.Range("f1_5")
etu = ws.Range("etu")

e1c = ws.Range("e1c")
e2c = ws.Range("e2c")
einc = ws.Range("einc")

sections = ws.Range("sections")
sections_points = ws.Range("sections_points")

maxcurv = ws.Range("maxcurv")
mincurv = ws.Range("mincurv")
curvinc = ws.Range("curvinc")

comp_model = ws.Range("comp_model")

alpha_cc = ws.Range("alpha_cc")

Eco_fib = ws.Range("Eco_fib")
alpha_E_fib = ws.Range("alpha_E_fib")

rho_AS3600 = ws.Range("rho_AS3600")

Dim pctCompl As Single
progress 0
ufProgress.Show

Call GenerateReinforcementTable

' Declare arrays for inputs to MATLAB calcs
ReDim Layers(1 To numberOfStages, 1) As Variant, Ages(1 To numberOfStages, 1) As Variant, rein_d(1 To 5, 1) As Variant, A_rebar(1 To 5, 1) As Variant
' String for spreadsheet range to get inputs from
Dim LayersR As String, AgesR As String, rein_dR As String, A_rebarR As String
' Covert convert spreadsheet column index number to alphabets
Dim ColLet As String, PasteRng As String

For i = 1 To numberOfStages
    ColLet = Col_Letter(1 + i)
    Let LayersR = ColLet & 71 & ":" & ColLet & 75
    Let AgesR = ColLet & 76 & ":" & ColLet & 80
    Layers = ws.Range(LayersR)
    Ages = ws.Range(AgesR)
    
    Worksheets("ReinforcementTemp").Activate
    Let rein_dR = "A" & ((i - 1) * 8 + 1 + 2) & ":" & "A" & ((i - 1) * 8 + 5 + 2)
    Let A_rebarR = "B" & ((i - 1) * 8 + 1 + 2) & ":" & "B" & ((i - 1) * 8 + 5 + 2)
    rein_d = Range(rein_dR)
    A_rebar = Range(A_rebarR)

    Application.StatusBar = "Solving Layer " & i

    res = MN_diagram(Layers, Ages, rein_d, A_rebar)
    
    ' Check that a worksheet called MNi exists
    On Error Resume Next
        sheetName = Worksheets("MN_" & i).Name
    On Error GoTo 0
    ' If it doesnt exist, create it
    If sheetName = "" Then
        Worksheets.Add.Name = "MN_" & i
    End If
    
    Worksheets("MN_" & i).Activate
    Worksheets("MN_" & i).Cells.Clear
    
    Let PasteRng = "A" & 1 & ":" & "D" & (sections_points * 2 + 3)
    Range(PasteRng) = res
    
    pctCompl = i / numberOfStages * 100
    progress pctCompl

Next i

ufProgress.Hide
Worksheets("Master").Activate
Application.Calculation = xlCalculationAutomatic

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub

Private Function GenerateReinforcementTable()
    Dim reincorcements As Range, reinforcementRow As Range
    
    On Error Resume Next
        sheetName = Worksheets("ReinforcementTemp").Name
    On Error GoTo 0

    If sheetName = "" Then
        Worksheets.Add.Name = "ReinforcementTemp"
    End If

    Worksheets("ReinforcementTemp").Activate
    Worksheets("ReinforcementTemp").Cells.Clear

    Set reinforcements = Range("Reinforcement")
    
    For Stage = 1 To numberOfStages
        Worksheets("ReinforcementTemp").Cells((Stage - 1) * 8 + 1, 1).Value = "Stage " & Stage
        Worksheets("ReinforcementTemp").Cells((Stage - 1) * 8 + 2, 1).Value = "rein_d"
        Worksheets("ReinforcementTemp").Cells((Stage - 1) * 8 + 2, 2).Value = "A_rebar"
        RowCount = 1
        
        For Each reinforcementRow In reinforcements.Rows
            If reinforcementRow.Cells(1, 7).Value <= Stage And IsNumeric(reinforcementRow.Cells(1, 3).Value) And IsNumeric(reinforcementRow.Cells(1, 6).Value) Then
                Worksheets("ReinforcementTemp").Cells((Stage - 1) * 8 + RowCount + 2, 1).Value = reinforcementRow.Cells(1, 3).Value
                Worksheets("ReinforcementTemp").Cells((Stage - 1) * 8 + RowCount + 2, 2).Value = reinforcementRow.Cells(1, 6).Value
                RowCount = RowCount + 1
            End If
        Next
    Next

End Function


Sub progress(pctCompl As Single)

ufProgress.Text.Caption = pctCompl & "% Completed"
ufProgress.Bar.Width = pctCompl * 2

DoEvents

End Sub

'----------------------------------------------------------------------------------------
'-------------------------'MATLAB' CALCULATIONS BELOW------------------------------------
'----------------------------------------------------------------------------------------

Private Function MN_diagram(Layers, Ages, rein_d, A_rebar) As Variant()

    Depths = generateCumulativeArray(Layers)
    
    num_ages = UBound(Ages) - LBound(Ages) + 1
    num_reo_ = UBound(rein_d) - LBound(rein_d) + 1
    
    num_layers = 0
    num_reo = 0
    
    For i = 1 To num_ages
        If Ages(i, 1) > 0 Then
            num_layers = num_layers + 1
        End If
    Next i
    
    For i = 1 To num_reo_
        If rein_d(i, 1) > 0 Then
            num_reo = num_reo + 1
        End If
    Next i
    
    beam_h = WorksheetFunction.Sum(Layers)
    s_tens = 0.33 'Tension strength gain (ft = tensile strength = ft*exp(s*(1-sqrt(28/t)))

    If inc_tensile = "No" Then
        fl_1 = 0 'MPa
        fr_4 = 0 'MPa
        etu = 0
        kh = 0 ' TR63 Secton 6.2.3 size effects factor
    ElseIf inc_tensile = "Yes" Then
        ' TR63 Secton 6.2.3 size effects factor
        If beam_h <= 0.125 Then
            kh = 1
        ElseIf beam_h >= 0.6 Then
            kh = 0.68
        Else
            kh = (1.6 - beam_h) / 1.475
        End If
    Else
        Err.Raise Number:=vbObjectError + 517, Description:="Unknown inc_tensile option, it should be 'Yes' or 'No'"
    End If
    
    ReDim fp_full(1 To num_layers, 1), ecp(1 To num_layers, 1), ecu(1 To num_layers, 1), E0(1 To num_layers, 1), Xi_u_CEB(1 To num_layers, 1), N_CEB(1 To num_layers, 1), n_AS3600(1 To num_layers, 1), k_pl(1 To num_layers, 1), fcmi(1 To num_layers, 1)
    
    For i = 1 To num_layers
        
        fp_full(i, 1) = fc_28 * Exp(s_comp * (1 - Sqr(28 / Ages(i, 1))))
    
        If comp_model = "AS3600_2009" Then
            If fp_full(i, 1) <= 40 Then
                fp_full_sum1 = fp_full_sum1 + fp_full(i, 1)
            Else
                fp_full_sum2 = fp_full_sum2 + fp_full(i, 1)
            End If
        End If
                
        If comp_model = "ENG1992_1_1" Then
            If fp_full(i, 1) < 50 Then
                fp_full_sum3 = fp_full_sum3 + fp_full(i, 1)
            Else
                fp_full_sum4 = fp_full_sum4 + fp_full(i, 1)
            End If
        End If

        If comp_model = "CEB_1998" Then
            ecp(i, 1) = 0.0022
            ecu(i, 1) = 0.025
            E0(i, 1) = 0.142 * 10 ^ 4 * (fp_full(i, 1) * alpha_cc * 0.14503773773 / 0.142) ^ (1 / 3) / 0.14503773773
            Es = fp_full(i, 1) * alpha_cc / ecp(i, 1)
            k_pl(i, 1) = E0(i, 1) / Es
            Xi_u_CEB(i, 1) = ((1 + 0.5 * k_pl(i, 1)) + Sqr(((-0.5 - 0.5 * k_pl(i, 1)) ^ 2) - 4 * 1 * 0.5)) / 2
            N_CEB(i, 1) = 4 * (Xi_u_CEB(i, 1) ^ 2 * (k_pl(i, 1) - 2) + 2 * Xi_u_CEB(i, 1) - k_pl(i, 1)) / (Xi_u_CEB(i, 1) * (k_pl(i, 1) - 2) + 1) ^ 2
            n_AS3600(i, 1) = -1
        ElseIf comp_model = "fib_2010" Then
            ecp(i, 1) = -1 * ((8.82767211332514 * 10 ^ -8) * (fp_full(i, 1) ^ 2) - (2.15138346961857E-05) * fp_full(i, 1) - 1.69196033099264E-03)
            ecu(i, 1) = 0.0035
            k_pl(i, 1) = -0.58167770410286 * Log(fp_full(i, 1)) + 3.96781843837164
            E0(i, 1) = Eco_fib * alpha_E_fib * (((fp_full(i, 1) + 8) / 10) ^ (1 / 3))
            Es(i, 1) = E0(i, 1) / k_pl(i, 1)
            Xi_u_CEB(i, 1) = -1
            N_CEB(i, 1) = -1
            n_AS3600(i, 1) = -1
        ElseIf comp_model = "AS3600_2009" Then
            fcmi(i, 1) = 0.9 * fp_full(i, 1) * (1.2875 - 0.001875 * fp_full(i, 1)) 'Function built for fp_full between 20 and 100 MPa
            If fp_full_sum1 > 0 And fp_full(i, 1) <= 40 Then
                E0(i, 1) = (rho_AS3600 ^ 1.5) * (0.043 * Sqr(fcmi(i, 1)))
            End If
            If fp_full_sum2 > 0 And fp_full(i, 1) > 40 Then
                E0(i, 1) = (rho_AS3600 ^ 1.5) * (0.12 + 0.024 * Sqr(fcmi(i, 1)))
            End If
            ecp(i, 1) = 4.11 * ((alpha_cc * fp_full(i, 1)) ^ 0.75) / E0(i, 1)
            ecu(i, 1) = 0.025
            k_pl(i, 1) = WorksheetFunction.Max(1, 0.67 + (alpha_cc * fp_full(i, 1)) / 62)
            Es = (alpha_cc * fp_full(i, 1)) / ecp(i, 1)
            Xi_u_CEB(i, 1) = -1
            N_CEB(i, 1) = -1
            n_AS3600(i, 1) = E0(i, 1) / (E0(i, 1) - Es)
        ElseIf comp_model = "EN1992_1_1" Then
            ecp(i, 1) = 0.001 * WorksheetFunction.Min(2.8, 0.7 * ((fp_full(i, 1) + 8) ^ 0.31))
            If fp_full_sum3 > 0 Then
                ecu(i, 1) = 0.0035
            End If
            If fp_full_sum4 > 0 Then
                ecu(i, 1) = 0.001 * (2.8 + 27 * ((98 - fp_full(i, 1) + 8) / 100) ^ 4)
            End If
            E0(i, 1) = 22000 * (((fp_full(i, 1) + 8) / 10) ^ 0.3)
            Es = -1
            k_pl(i, 1) = 1.05 * E0(i, 1) * ecp(i, 1) / (fp_full(i, 1) + 8)
            Xi_u_CEB(i, 1) = -1
            N_CEB(i, 1) = -1
            n_AS3600(i, 1) = -1
        Else
            Err.Raise Number:=vbObjectError + 513, Description:="Unkown compressive stress model"
        End If
    Next i
    
    ' INITIALIZE VARIABLES BASED ON INPUTS
    ' Concrete tensile strength (in MPa): 0 for plain concrete, <0 for fibre-reinforced
    ReDim fr_1_layers(1 To num_layers, 1), fr_4_layers(1 To num_layers, 1), fl_5_layers(1 To num_layers, 1)
    
    If inc_tensile = "Yes" Then
        For i = 1 To num_layers
            fr_1_layers(i, 1) = -1 * fr_1 * Exp(s_tens * (1 - Sqr(28 / Ages(i, 1))))
            fr_4_layers(i, 1) = -1 * fr_4 * Exp(s_tens * (1 - Sqr(28 / Ages(i, 1))))
            fl_5_layers(i, 1) = -1 * f1_5 * Exp(s_tens * (1 - Sqr(28 / Ages(i, 1))))
        Next i
    End If
    
    'A_beam = beam_b*beam_h; # Area of beam (in m^2)

    If rein_type = 1 Then
        A_steel = 0 ' Area of steel (in m^2)
        rein_d = 0
        A_rein = 0
    ElseIf rein_type > 1 Then 'Install reins
        bal_et = -steel_fy / steel_E1 ' Tensile strain at balanced condition
        ' Area calculations
        A_steel = WorksheetFunction.Sum(A_rein) ' Total area of rein in beam (in m^2)
    End If
    
    'A_conc = A_beam-A_steel # Area of concrete in beam (in m^2)

    ' Plastic Centroid
    ec_p = ecp(1, 1) 'concrete cracking limit - Assumed to equal the strain for peak compressive strength of the oldest shotcrete layer
    ReDim curv_p(1 To 1, 1)
    curv_p(1, 1) = 0
    
    cumresults_p = calculate_MN(curv_p, ec_p, -1)
    cumforce_p = cumresults_p(1, 1)
    cummoment_p = cumresults_p(1, 2)
    
    pc = cummoment_p / cumforce_p

    ' Calculate pure tension point
    If rein_type = 1 Then
        ' no mesh, set the strain to a large value for fibre-reinforced
        ec_t = -0.02
    Else
        ' Mesh in the section, set the train equal to the uniform strain (i.e. onset of necking) for L Class Mesh
        ec_t = -0.015
    End If
    Dim curv_t(1 To 1, 1) As Variant
    curv_t(1, 1) = 0
    
    cumresults_t = calculate_MN(curv_t, ec_t, pc)
    Prten = cumresults_t(1, 1)
    Prten_M = cumresults_t(1, 2)
    
    'Solve for axial force and bending moment for each strain
    'Establish vector of strains to sample (each strain is a point along the M-N curve
    ec_range = linspace(e1c, e2c, Round(Abs((e2c - e1c) / einc), 0), True)
    curv_range = linspace(mincurv, maxcurv, Round(Abs(maxcurv - mincurv) / curvinc), True)
    q = arrlen(ec_range)
    r = arrlen(curv_range)

    ' ULS capacity reduction factor
    phi_AS = 0
    If rein_type = 1 Then
        phi_AS = 0.6
    ElseIf rein_type = 2 Then
        phi_AS = 0.65
    End If
    
    ' Add MN diagram results to a 4 column array
    ReDim MNcurveFull(1 To (r * q), 1 To 4) As Variant
    
    'ReDim cumforce_mn(1 To r, 1 To q) As Variant, cummoment_mn(1 To r, 1 To q) As Variant
    ReDim cumforce_mn(1 To (r * q), 1), cummoment_mn(1 To (r * q), 1)
    For i = 1 To q
        ec = ec_range(i, 1)
        cum_temp = calculate_MN(curv_range, ec, pc)
        
        For j = 1 To r
            row_idx = (i - 1) * r + j
            cumforce_mn(row_idx, 1) = cum_temp(j, 1)
            cummoment_mn(row_idx, 1) = cum_temp(j, 2)
            Next j
    Next i

    Max_Axial = Application.WorksheetFunction.Max(cumforce_mn)
    Min_Axial = Application.WorksheetFunction.Min(cumforce_mn)
    
    Axial_Increment = (Max_Axial - Min_Axial) / sections_points
    axial_range = linspace(Min_Axial, Max_Axial, (sections_points + 1))
    
    For i = 1 To (r * q)
        If cumforce_mn(i, 1) < Prten Then
        ' The assumptions behind Prten are not producing the minimal axial force
            Prten = cumforce_mn(i, 1)
            Prten_M = cummoment_mn(i, 1)
        End If
        
        If cumforce_mn(i, 1) > cumforce_p Then
            cumforce_p = cumforce_mn(i, 1)
            cummoment_p = cummoment_mn(i, 1)
        End If
    
    Next i
    
    ' ULS capacity reduction factor
    phi_AS = 0
    If rein_type = 1 Then
        phi_AS = 0.6
    ElseIf rein_type = 2 Then
        phi_AS = 0.65
    End If
    
    ' Add MN diagram results to a 4 column array
    ReDim MNcurveFull(0 To sections_points * 2 + 2, 1 To 4) As Variant
    
    MNcurveFull(0, 2) = Prten
    MNcurveFull(0, 1) = Prten_M
    MNcurveFull(sections_points * 2 + 2, 2) = Prten
    MNcurveFull(sections_points * 2 + 2, 1) = Prten_M
    
    MNcurveFull(0, 3) = Prten * phi_AS
    MNcurveFull(0, 4) = Prten_M * phi_AS
    MNcurveFull(sections_points * 2 + 2, 4) = Prten * phi_AS
    MNcurveFull(sections_points * 2 + 2, 3) = Prten_M * phi_AS
    
    MNcurveFull(sections_points + 1, 2) = cumforce_p
    MNcurveFull(sections_points + 1, 1) = cummoment_p
    
    MNcurveFull(sections_points + 1, 4) = cumforce_p * phi_AS
    MNcurveFull(sections_points + 1, 3) = cummoment_p * phi_AS
    
    For i = 1 To sections_points
        x_1 = axial_range(i, 1)
        x_2 = axial_range((i + 1), 1)
        
        minF_temp = 0
        maxF_temp = 0
        minM_temp = 1000000
        maxM_temp = -1000000
        
        For j = 1 To (r * q)
            If cumforce_mn(j, 1) > x_1 And cumforce_mn(j, 1) <= x_2 Then
                If cummoment_mn(j, 1) > maxM_temp Then
                    maxF_temp = cumforce_mn(j, 1)
                    maxM_temp = cummoment_mn(j, 1)
                ElseIf cummoment_mn(j, 1) < minM_temp Then
                    minF_temp = cumforce_mn(j, 1)
                    minM_temp = cummoment_mn(j, 1)
                End If
            End If
        Next j
        
        row_idx1 = i
        row_idx2 = sections_points * 2 - i + 2

        ' Unfactored MN Diagram
        MNcurveFull(row_idx1, 2) = minF_temp
        MNcurveFull(row_idx1, 1) = minM_temp
        MNcurveFull(row_idx2, 2) = maxF_temp
        MNcurveFull(row_idx2, 1) = maxM_temp
        ' Factored MN Diagram
        MNcurveFull(row_idx1, 4) = minF_temp * phi_AS
        MNcurveFull(row_idx1, 3) = minM_temp * phi_AS
        MNcurveFull(row_idx2, 4) = maxF_temp * phi_AS
        MNcurveFull(row_idx2, 3) = maxM_temp * phi_AS
        
    Next i
    
    MN_diagram = MNcurveFull
    
End Function
Private Function calculate_MN(curv, ec, pc) As Variant()
    
    Dim i As Integer, j As Integer, in_layer As Integer
    Dim x1 As Variant, x2 As Variant, xm As Variant
    
    num_curvs = arrlen(curv)
    ReDim result(1 To num_curvs, 2) As Variant
    
    For i = 1 To num_curvs
    
        conc_cumforce = 0
        conc_cummoment = 0
        rein_cumforce = 0
        rein_cummoment = 0
        
        For j = 1 To sections
        
            x1 = (j - 1) * (beam_h / sections)
            x2 = j * (beam_h / sections)
            xm = (x1 + x2) / 2
            
            conc_stress1 = concrete_stress(x1, curv(i, 1), ec)
            conc_stress2 = concrete_stress(x2, curv(i, 1), ec)
            conc_stress = 0.5 * (conc_stress1 + conc_stress2)
            
            conc_force = conc_stress * beam_b * (beam_h / sections)
            
            If pc = -1 Then
                conc_moment = conc_force * xm
            Else
                conc_moment = conc_force * (pc - xm)
            End If
            
            conc_cumforce = conc_cumforce + conc_force
            conc_cummoment = conc_cummoment + conc_moment
            
        Next j
        
        If rein_type > 1 Then
            For k = 1 To num_reo
                x = rein_d(k, 1)
                st_stress = steel_stress(x, curv(i, 1), ec)
                rein_force = st_stress * A_rebar(k, 1)
                
                If pc = -1 Then
                    rein_moment = rein_force * x
                Else
                    rein_moment = rein_force * (pc - x)
                End If
                
                rein_cumforce = rein_cumforce + rein_force
                rein_cummoment = rein_cummoment + rein_moment

            Next k
        End If
            
        result(i, 1) = conc_cumforce + rein_cumforce
        result(i, 2) = conc_cummoment + rein_cummoment

    Next i
    
    calculate_MN = result
                
End Function

Private Function concrete_stress(x, curv, ec) As Variant
    ' Calculate the strain point
    Dim comp_idx As Boolean, tens_idx As Boolean
    
    in_layer = 1
    For k = 1 To (num_layers - 1)
        If x > Depths(k, 1) Then
            in_layer = k + 1
        End If
    Next k
    
    strain = ec - curv * x
    ' Compressive strain
    If strain > 0 And strain < ecu(k, 1) Then
        ' Calcualte compressive stresses based on compressive model
        If comp_model = "CEB_1998" Then
            Xi = strain / ecp(k, 1)
            fp = fp_full(k, 1) * alpha_cc
            If Xi <= Xi_u_CEB(k, 1) Then
                result = fp * (k_pl(k, 1) * Xi - (Xi ^ 2)) / (1 + (k_pl(k, 1) - 2) * Xi)
            ElseIf Xi > Xi_u_CEB(k, 1) Then
                result = fp * ((((N_CEB(k, 1) / Xi_u_CEB(k, 1)) - (2 / (Xi_u_CEB(k, 1) ^ 2))) * (Xi ^ 2) + Xi * ((4 / Xi_u_CEB(k, 1)) - N_CEB(k, 1))) ^ -1)
            End If
        ElseIf comp_model = "fib_2010" Then
            fp = (fp_full(k, 1) + 8) * alpha_cc * (fp_full(k, 1) / (fp_full(k, 1) + 8))
            neta = strain / ecp(k, 1)
            result = fp * ((k_pl(k, 1) * neta - (neta ^ 2)) / (1 + (k_pl(i, 1) - 2) * neta))
        ElseIf comp_model = "AS3600_2009" Then
            fp = fp_full(k, 1) * alpha_cc
            neta = strain / ecp(k, 1)
            result = fp * n_AS3600(k, 1) * neta / (n_AS3600(k, 1) - 1 + (neta) ^ (n_AS3600(k, 1) * k_pl(k, 1)))
        ElseIf comp_model = "EN1992_1_1" Then
            fp = (fp_full(k, 1) + 8) * alpha_cc * (fp_full(k, 1) / (fp_full(k, 1) + 8))
            neta = strain / ecp(k, 1)
            result = fp * (k_pl(k, 1) * neta - (neta ^ 2)) / (1 + (k_pl(k, 1) - 2) * neta)
        Else:
            Err.Raise Number:=vbObjectError + 513, Description:="Unkown compressive stress model"
        End If
        
    End If
        
    ' Tensile strain
    If strain < 0 And strain > etu Then
        If tens_model = "zero" Then
            result = 0
        ElseIf tens_model = "SFRS_TR63" Then
            ' Uses the TR63 tensile stress block from Section 6.2.3, which is based on EN14651 beam test parameters
            fb = 0.37 * kh * fr_4_layers(k, 1)
            sigma_2 = 0.45 * kh * fr_1_layers(k, 1)
            result = fb - (etu - strain) * (fb - sigma_2) / etu
        ElseIf tens_model = "SFRC_DAFStb" Then
            ' Uses the rectangular tensile stress block from Figure R.2, it is assumed the f,f,ctr,u strength is equivalent to fr,4 from an EN14651 beam test
            ' Assuming 0.37 factor to convert fr_4 to residual tensile strength and a 0.85 long term effects factor
            result = 0.85 * 0.37 * fr_4_layers(k, 1)
        ElseIf tens_model = "SFRCbars_AS" Then
            ' Uses the rectangular tensile stress block from AS3600:2018 Clause 16.4.2, note this method REQUIRES bars or mesh to be present
            result = 0.9 * 1 * fl_5_layers(k, 1) ' Conservatively assuming kg factor = 1 (as we don't know the ultimate strength tensile zone dimensions)
        Else:
            Err.Raise Number:=vbObjectError + 514, Description:="Unkown tensile stress model"
        End If
        
    End If

    concrete_stress = result

End Function

Private Function steel_stress(x, curv, ec) As Variant
    ' Calculate the strain point
    
    strain = ec - curv * x
    If Abs(strain) * steel_E1 < steel_fy Then
        result = strain * steel_E1
    ElseIf Abs(strain) * steel_E1 >= steel_fy Then
        result = Sgn(strain) * steel_fy
    Else
        result = 0
    End If

    steel_stress = result
    
End Function

