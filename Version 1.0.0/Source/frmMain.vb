'Dave Kurtz
'OctoPlan - an Octopath Traveler Character Planner
'Last Revised 8-1-18
'Latest Published Version: 1.0.0
'Currently Working On: DONE!

Public Class frmMain
    'Declare variables to be used in multiple subs
    Dim intRunningTotal As Integer
    Dim intPassivesArray(12, 4) As Integer
    Dim blnDupes As Boolean = False

    Private Sub combo_boxes_Input_Validation(sender As Object, e As KeyPressEventArgs) Handles cmb1PartyMember.KeyPress, cmb1Passive1.KeyPress,
    cmb1Passive2.KeyPress, cmb1Passive3.KeyPress, cmb1Passive4.KeyPress, cmb1SubJob.KeyPress, cmb2PartyMember.KeyPress, cmb2Passive1.KeyPress,
    cmb2Passive3.KeyPress, cmb2Passive2.KeyPress, cmb2Passive4.KeyPress, cmb2SubJob.KeyPress, cmb3PartyMember.KeyPress, cmb3Passive1.KeyPress,
    cmb3Passive2.KeyPress, cmb3Passive3.KeyPress, cmb3Passive4.KeyPress, cmb3SubJob.KeyPress, cmb4PartyMember.KeyPress, cmb4Passive1.KeyPress,
    cmb4Passive2.KeyPress, cmb4Passive3.KeyPress, cmb4Passive4.KeyPress, cmb4SubJob.KeyPress
        'We don't want people typing random stuff into these combo boxes
        e.Handled = True
    End Sub

    Private Sub cmb1PartyMember_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb1PartyMember.SelectedIndexChanged, cmb1SubJob.SelectedIndexChanged
        'Check to make sure there aren't duplicate party members
        If cmb1PartyMember.SelectedIndex > 0 Then
            If cmb1PartyMember.SelectedIndex = cmb2PartyMember.SelectedIndex Or cmb1PartyMember.SelectedIndex = cmb3PartyMember.SelectedIndex Or cmb1PartyMember.SelectedIndex = cmb4PartyMember.SelectedIndex Then
                'Here there be Dupes
                MsgBox("You can't have " & cmb1PartyMember.SelectedItem & " in your party more than once!", MsgBoxStyle.OkOnly, "Duplicate Member Detected")
                cmb1PartyMember.SelectedIndex = 0
            End If
        End If

        'Check to make sure Selected Sub-job isn't duped or invalid
        If cmb1SubJob.SelectedIndex > 0 Then
            If cmb1SubJob.SelectedIndex = cmb1PartyMember.SelectedIndex Then
                'Invalid sub-job
                MsgBox(cmb1PartyMember.SelectedItem & " cannot be assigned the " & cmb1SubJob.SelectedItem & " sub-job!", MsgBoxStyle.OkOnly, "Invalid Sub-Job Detected")
                cmb1SubJob.SelectedIndex = 0
            ElseIf cmb1SubJob.SelectedIndex = cmb2SubJob.SelectedIndex Or cmb1SubJob.SelectedIndex = cmb3SubJob.SelectedIndex Or cmb1SubJob.SelectedIndex = cmb4SubJob.SelectedIndex Then
                'Duplicate sub-job
                MsgBox("Another character already has the " & cmb1SubJob.SelectedItem & " sub-job equipped!", MsgBoxStyle.OkOnly, "Duplicate Sub-Job Detected")
                cmb1SubJob.SelectedIndex = 0
            End If
        End If

        'Clear all the stuffs!
        chk1Axe.Checked = False
        chk1Bow.Checked = False
        chk1Dagger.Checked = False
        chk1Dark.Checked = False
        chk1Fire.Checked = False
        chk1Ice.Checked = False
        chk1Lance.Checked = False
        chk1Light.Checked = False
        chk1Staff.Checked = False
        chk1Sword.Checked = False
        chk1Thunder.Checked = False
        chk1Wind.Checked = False
        txt1Stats.Text = String.Empty
        pic1.Image = Nothing
        lbl1Path.BackColor = Nothing
        lbl1Path.Text = Nothing

        'Change Party Member 1 pic+path+primary proficiencies to the selected person (or none)
        Select Case cmb1PartyMember.SelectedIndex
            Case 1
                pic1.Image = OctoPlan.My.Resources.Ophelia
                lbl1Path.BackColor = Color.Tomato
                lbl1Path.Text = "Guide"
                chk1Staff.Checked = True
                chk1Light.Checked = True
            Case 2
                pic1.Image = OctoPlan.My.Resources.Cyrus
                lbl1Path.BackColor = Color.MediumSlateBlue
                lbl1Path.Text = "Scrutinize"
                chk1Staff.Checked = True
                chk1Ice.Checked = True
                chk1Fire.Checked = True
                chk1Thunder.Checked = True
            Case 3
                pic1.Image = OctoPlan.My.Resources.Tressa
                lbl1Path.BackColor = Color.Yellow
                lbl1Path.Text = "Purchase"
                chk1Lance.Checked = True
                chk1Bow.Checked = True
                chk1Wind.Checked = True
            Case 4
                pic1.Image = OctoPlan.My.Resources.Olberic
                lbl1Path.BackColor = Color.LightGreen
                lbl1Path.Text = "Challenge"
                chk1Sword.Checked = True
                chk1Lance.Checked = True
            Case 5
                pic1.Image = OctoPlan.My.Resources.Primrose
                lbl1Path.BackColor = Color.Tomato
                lbl1Path.Text = "Allure"
                chk1Dagger.Checked = True
                chk1Dark.Checked = True
            Case 6
                pic1.Image = OctoPlan.My.Resources.Alfyn
                lbl1Path.BackColor = Color.MediumSlateBlue
                lbl1Path.Text = "Inquire"
                chk1Axe.Checked = True
                chk1Ice.Checked = True
            Case 7
                pic1.Image = OctoPlan.My.Resources.Therion
                lbl1Path.BackColor = Color.Yellow
                lbl1Path.Text = "Steal"
                chk1Sword.Checked = True
                chk1Dagger.Checked = True
                chk1Fire.Checked = True
            Case 8
                pic1.Image = OctoPlan.My.Resources.H_aanit
                lbl1Path.BackColor = Color.LightGreen
                lbl1Path.Text = "Provoke"
                chk1Axe.Checked = True
                chk1Bow.Checked = True
                chk1Thunder.Checked = True
        End Select

        'Set additional proficiencies + bonus stats based on sub-job chosen
        Select Case cmb1SubJob.SelectedIndex
            Case 1 'Cleric
                chk1Staff.Checked = True
                chk1Light.Checked = True
                txt1Stats.Text = "SP: +17.78%" & vbNewLine & "El. Attack: +4.40%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 26.35%"
            Case 2 'Scholar
                chk1Staff.Checked = True
                chk1Ice.Checked = True
                chk1Fire.Checked = True
                chk1Thunder.Checked = True
                txt1Stats.Text = "SP: +11.11%" & vbNewLine & "El. Attack: +9.89%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 25.17%"
            Case 3 'Merchant
                chk1Lance.Checked = True
                chk1Bow.Checked = True
                chk1Wind.Checked = True
                txt1Stats.Text = "HP: +7.86%" & vbNewLine & "SP: +6.67%" & vbNewLine & "Phys Attack: +2.75%" & vbNewLine & "Phys Defense: +2.78%" & vbNewLine & "El. Attack: +2.20%" & vbNewLine & "El. Defense: +2.78%" & vbNewLine & "Total Increase: 25.04%"
            Case 4 'Warrior
                chk1Sword.Checked = True
                chk1Lance.Checked = True
                txt1Stats.Text = "HP: +17.83%" & vbNewLine & "Phys Attack: +4.59%" & vbNewLine & "Phys Defense: +4.17%" & vbNewLine & "Total Increase: 26.59%"
            Case 5 'Dancer
                chk1Dagger.Checked = True
                chk1Dark.Checked = True
                txt1Stats.Text = "El. Attack: +7.69%" & vbNewLine & "Speed: +9.89%" & vbNewLine & "Evasion: +10%" & vbNewLine & "Total Increase: 27.58%"
            Case 6 'Apothecary
                chk1Axe.Checked = True
                chk1Ice.Checked = True
                txt1Stats.Text = "HP: +19.96%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 26.28%"
            Case 7 'Thief
                chk1Sword.Checked = True
                chk1Dagger.Checked = True
                chk1Fire.Checked = True
                txt1Stats.Text = "Phys Attack: +1.83%" & vbNewLine & "Accuracy: +4.59%" & vbNewLine & "Speed: +7.69%" & vbNewLine & "Crit: +5%" & vbNewLine & "Evade: +8%" & vbNewLine & "Total Increase: 27.11%"
            Case 8 'Hunter
                chk1Axe.Checked = True
                chk1Bow.Checked = True
                chk1Thunder.Checked = True
                txt1Stats.Text = "Phys Attack: +7.84%" & vbNewLine & "Accuracy: +5.38%" & vbNewLine & "Speed: +4.05%" & vbNewLine & "Crit: +5.38%" & vbNewLine & "Evade: +2.41%" & vbNewLine & "Total Increase: 25.06%"
            Case 9 'Warmaster
                chk1Sword.Checked = True
                chk1Lance.Checked = True
                chk1Dagger.Checked = True
                chk1Axe.Checked = True
                chk1Bow.Checked = True
                chk1Staff.Checked = True
                txt1Stats.Text = "Phys Attack: +11.93%" & vbNewLine & "Phys Defense: +11.11%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Speed: +5.49%" & vbNewLine & "Evade: +4%" & vbNewLine & "Total Increase: 38.03%"
            Case 10 'Runelord
                chk1Sword.Checked = True
                chk1Axe.Checked = True
                chk1Ice.Checked = True
                chk1Fire.Checked = True
                chk1Thunder.Checked = True
                chk1Wind.Checked = True
                chk1Dark.Checked = True
                chk1Light.Checked = True
                txt1Stats.Text = "Phys Attack: +6.42%" & vbNewLine & "Phys Defense: +5.56%" & vbNewLine & "El. Attack: +6.59%" & vbNewLine & "El. Defense: +5.56%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Crit: +8%" & vbNewLine & "Total Increase: 37.63%"
            Case 11 'Starseer
                chk1Lance.Checked = True
                chk1Dagger.Checked = True
                chk1Wind.Checked = True
                chk1Dark.Checked = True
                chk1Light.Checked = True
                txt1Stats.Text = "HP: +14.86%" & vbNewLine & "SP: +13.33%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "El. Defense: +1.39%" & vbNewLine & "Evade: +2%" & vbNewLine & "Total Increase: 35.90%"
            Case 12 'Sorceror
                chk1Bow.Checked = True
                chk1Staff.Checked = True
                chk1Ice.Checked = True
                chk1Fire.Checked = True
                chk1Thunder.Checked = True
                chk1Wind.Checked = True
                chk1Dark.Checked = True
                chk1Light.Checked = True
                txt1Stats.Text = "SP: +20%" & vbNewLine & "El. Attack: +8.79%" & vbNewLine & "El. Defense: +8.33%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 39.12%"
        End Select
    End Sub

    Private Sub lbl1Path_BackColorChanged(sender As Object, e As EventArgs) Handles lbl1Path.BackColorChanged,
        lbl2Path.BackColorChanged, lbl3Path.BackColorChanged, lbl4Path.BackColorChanged
        'Check to see if Tomato is covered
        If lbl1Path.BackColor = Color.Tomato Or lbl2Path.BackColor = Color.Tomato Or lbl3Path.BackColor = Color.Tomato Or lbl4Path.BackColor = Color.Tomato Then
            'At least one member is Tomato; throw Red into Covered status
            lblCoverageRed.Text = "Covered!"
        Else
            'No Tomato for the potato :(
            lblCoverageRed.Text = "Not in Party"
        End If

        'Check to see if MediumSlateBlue is covered
        If lbl1Path.BackColor = Color.MediumSlateBlue Or lbl2Path.BackColor = Color.MediumSlateBlue Or lbl3Path.BackColor = Color.MediumSlateBlue Or lbl4Path.BackColor = Color.MediumSlateBlue Then
            'At least one member Slated; throw Blue into Covered status
            lblCoverageBlue.Text = "Covered!"
        Else
            'No slate in this estate
            lblCoverageBlue.Text = "Not in Party"
        End If

        'Check for old yeller
        If lbl1Path.BackColor = Color.Yellow Or lbl2Path.BackColor = Color.Yellow Or lbl3Path.BackColor = Color.Yellow Or lbl4Path.BackColor = Color.Yellow Then
            'It's yellow; let it mellow
            lblCoverageYellow.Text = "Covered!"
        Else
            'It's brown; flush it down
            lblCoverageYellow.Text = "Not in Party"
        End If

        'Why can't I, hold all these limes?
        If lbl1Path.BackColor = Color.LightGreen Or lbl2Path.BackColor = Color.LightGreen Or lbl3Path.BackColor = Color.LightGreen Or lbl4Path.BackColor = Color.LightGreen Then
            'TEENAGE MUTANT NINJA TURTLES
            lblCoverageGreen.Text = "Covered!"
        Else
            'Do not pass Go, do not collect 100 leaves
            lblCoverageGreen.Text = "Not in Party"
        End If
    End Sub

    Private Sub cmb2PartyMember_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb2PartyMember.SelectedIndexChanged, cmb2SubJob.SelectedIndexChanged
        'Check to make sure there aren't duplicate party members
        If cmb2PartyMember.SelectedIndex > 0 Then
            If cmb2PartyMember.SelectedIndex = cmb1PartyMember.SelectedIndex Or cmb2PartyMember.SelectedIndex = cmb3PartyMember.SelectedIndex Or cmb2PartyMember.SelectedIndex = cmb4PartyMember.SelectedIndex Then
                'Here there be Dupes
                MsgBox("You can't have " & cmb2PartyMember.SelectedItem & " in your party more than once!", MsgBoxStyle.OkOnly, "Duplicate Member Detected")
                cmb2PartyMember.SelectedIndex = 0
            End If
        End If

        'Check to make sure Selected Sub-job isn't duped or invalid
        If cmb2SubJob.SelectedIndex > 0 Then
            If cmb2SubJob.SelectedIndex = cmb2PartyMember.SelectedIndex Then
                'Invalid sub-job
                MsgBox(cmb2PartyMember.SelectedItem & " cannot be assigned the " & cmb2SubJob.SelectedItem & " sub-job!", MsgBoxStyle.OkOnly, "Invalid Sub-Job Detected")
                cmb2SubJob.SelectedIndex = 0
            ElseIf cmb2SubJob.SelectedIndex = cmb1SubJob.SelectedIndex Or cmb2SubJob.SelectedIndex = cmb3SubJob.SelectedIndex Or cmb2SubJob.SelectedIndex = cmb4SubJob.SelectedIndex Then
                'Duplicate sub-job
                MsgBox("Another character already has the " & cmb2SubJob.SelectedItem & " sub-job equipped!", MsgBoxStyle.OkOnly, "Duplicate Sub-Job Detected")
                cmb2SubJob.SelectedIndex = 0
            End If
        End If

        'Clear all the stuffs!
        chk2Axe.Checked = False
        chk2Bow.Checked = False
        chk2Dagger.Checked = False
        chk2Dark.Checked = False
        chk2Fire.Checked = False
        chk2Ice.Checked = False
        chk2Lance.Checked = False
        chk2Light.Checked = False
        chk2Staff.Checked = False
        chk2Sword.Checked = False
        chk2Thunder.Checked = False
        chk2Wind.Checked = False
        txt2Stats.Text = String.Empty
        pic2.Image = Nothing
        lbl2Path.BackColor = Nothing
        lbl2Path.Text = Nothing

        'Change Party Member 2 pic+path+primary proficiencies to the selected person (or none)
        Select Case cmb2PartyMember.SelectedIndex
            Case 1
                pic2.Image = OctoPlan.My.Resources.Ophelia
                lbl2Path.BackColor = Color.Tomato
                lbl2Path.Text = "Guide"
                chk2Staff.Checked = True
                chk2Light.Checked = True
            Case 2
                pic2.Image = OctoPlan.My.Resources.Cyrus
                lbl2Path.BackColor = Color.MediumSlateBlue
                lbl2Path.Text = "Scrutinize"
                chk2Staff.Checked = True
                chk2Ice.Checked = True
                chk2Fire.Checked = True
                chk2Thunder.Checked = True
            Case 3
                pic2.Image = OctoPlan.My.Resources.Tressa
                lbl2Path.BackColor = Color.Yellow
                lbl2Path.Text = "Purchase"
                chk2Lance.Checked = True
                chk2Bow.Checked = True
                chk2Wind.Checked = True
            Case 4
                pic2.Image = OctoPlan.My.Resources.Olberic
                lbl2Path.BackColor = Color.LightGreen
                lbl2Path.Text = "Challenge"
                chk2Sword.Checked = True
                chk2Lance.Checked = True
            Case 5
                pic2.Image = OctoPlan.My.Resources.Primrose
                lbl2Path.BackColor = Color.Tomato
                lbl2Path.Text = "Allure"
                chk2Dagger.Checked = True
                chk2Dark.Checked = True
            Case 6
                pic2.Image = OctoPlan.My.Resources.Alfyn
                lbl2Path.BackColor = Color.MediumSlateBlue
                lbl2Path.Text = "Inquire"
                chk2Axe.Checked = True
                chk2Ice.Checked = True
            Case 7
                pic2.Image = OctoPlan.My.Resources.Therion
                lbl2Path.BackColor = Color.Yellow
                lbl2Path.Text = "Steal"
                chk2Sword.Checked = True
                chk2Dagger.Checked = True
                chk2Fire.Checked = True
            Case 8
                pic2.Image = OctoPlan.My.Resources.H_aanit
                lbl2Path.BackColor = Color.LightGreen
                lbl2Path.Text = "Provoke"
                chk2Axe.Checked = True
                chk2Bow.Checked = True
                chk2Thunder.Checked = True
        End Select

        'Set additional proficiencies + bonus stats based on sub-job chosen
        Select Case cmb2SubJob.SelectedIndex
            Case 1 'Cleric
                chk2Staff.Checked = True
                chk2Light.Checked = True
                txt2Stats.Text = "SP: +17.78%" & vbNewLine & "El. Attack: +4.40%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 26.35%"
            Case 2 'Scholar
                chk2Staff.Checked = True
                chk2Ice.Checked = True
                chk2Fire.Checked = True
                chk2Thunder.Checked = True
                txt2Stats.Text = "SP: +11.11%" & vbNewLine & "El. Attack: +9.89%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 25.17%"
            Case 3 'Merchant
                chk2Lance.Checked = True
                chk2Bow.Checked = True
                chk2Wind.Checked = True
                txt2Stats.Text = "HP: +7.86%" & vbNewLine & "SP: +6.67%" & vbNewLine & "Phys Attack: +2.75%" & vbNewLine & "Phys Defense: +2.78%" & vbNewLine & "El. Attack: +2.20%" & vbNewLine & "El. Defense: +2.78%" & vbNewLine & "Total Increase: 25.04%"
            Case 4 'Warrior
                chk2Sword.Checked = True
                chk2Lance.Checked = True
                txt2Stats.Text = "HP: +17.83%" & vbNewLine & "Phys Attack: +4.59%" & vbNewLine & "Phys Defense: +4.17%" & vbNewLine & "Total Increase: 26.59%"
            Case 5 'Dancer
                chk2Dagger.Checked = True
                chk2Dark.Checked = True
                txt2Stats.Text = "El. Attack: +7.69%" & vbNewLine & "Speed: +9.89%" & vbNewLine & "Evasion: +10%" & vbNewLine & "Total Increase: 27.58%"
            Case 6 'Apothecary
                chk2Axe.Checked = True
                chk2Ice.Checked = True
                txt2Stats.Text = "HP: +19.96%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 26.28%"
            Case 7 'Thief
                chk2Sword.Checked = True
                chk2Dagger.Checked = True
                chk2Fire.Checked = True
                txt2Stats.Text = "Phys Attack: +1.83%" & vbNewLine & "Accuracy: +4.59%" & vbNewLine & "Speed: +7.69%" & vbNewLine & "Crit: +5%" & vbNewLine & "Evade: +8%" & vbNewLine & "Total Increase: 27.11%"
            Case 8 'Hunter
                chk2Axe.Checked = True
                chk2Bow.Checked = True
                chk2Thunder.Checked = True
                txt2Stats.Text = "Phys Attack: +7.84%" & vbNewLine & "Accuracy: +5.38%" & vbNewLine & "Speed: +4.05%" & vbNewLine & "Crit: +5.38%" & vbNewLine & "Evade: +2.41%" & vbNewLine & "Total Increase: 25.06%"
            Case 9 'Warmaster
                chk2Sword.Checked = True
                chk2Lance.Checked = True
                chk2Dagger.Checked = True
                chk2Axe.Checked = True
                chk2Bow.Checked = True
                chk2Staff.Checked = True
                txt2Stats.Text = "Phys Attack: +11.93%" & vbNewLine & "Phys Defense: +11.11%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Speed: +5.49%" & vbNewLine & "Evade: +4%" & vbNewLine & "Total Increase: 38.03%"
            Case 10 'Runelord
                chk2Sword.Checked = True
                chk2Axe.Checked = True
                chk2Ice.Checked = True
                chk2Fire.Checked = True
                chk2Thunder.Checked = True
                chk2Wind.Checked = True
                chk2Dark.Checked = True
                chk2Light.Checked = True
                txt2Stats.Text = "Phys Attack: +6.42%" & vbNewLine & "Phys Defense: +5.56%" & vbNewLine & "El. Attack: +6.59%" & vbNewLine & "El. Defense: +5.56%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Crit: +8%" & vbNewLine & "Total Increase: 37.63%"
            Case 11 'Starseer
                chk2Lance.Checked = True
                chk2Dagger.Checked = True
                chk2Wind.Checked = True
                chk2Dark.Checked = True
                chk2Light.Checked = True
                txt2Stats.Text = "HP: +14.86%" & vbNewLine & "SP: +13.33%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "El. Defense: +1.39%" & vbNewLine & "Evade: +2%" & vbNewLine & "Total Increase: 35.90%"
            Case 12 'Sorceror
                chk2Bow.Checked = True
                chk2Staff.Checked = True
                chk2Ice.Checked = True
                chk2Fire.Checked = True
                chk2Thunder.Checked = True
                chk2Wind.Checked = True
                chk2Dark.Checked = True
                chk2Light.Checked = True
                txt2Stats.Text = "SP: +20%" & vbNewLine & "El. Attack: +8.79%" & vbNewLine & "El. Defense: +8.33%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 39.12%"
        End Select
    End Sub

    Private Sub cmb3PartyMember_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb3PartyMember.SelectedIndexChanged, cmb3SubJob.SelectedIndexChanged
        'Check to make sure there aren't duplicate party members
        If cmb3PartyMember.SelectedIndex > 0 Then
            If cmb3PartyMember.SelectedIndex = cmb1PartyMember.SelectedIndex Or cmb3PartyMember.SelectedIndex = cmb2PartyMember.SelectedIndex Or cmb3PartyMember.SelectedIndex = cmb4PartyMember.SelectedIndex Then
                'Here there be Dupes
                MsgBox("You can't have " & cmb3PartyMember.SelectedItem & " in your party more than once!", MsgBoxStyle.OkOnly, "Duplicate Member Detected")
                cmb3PartyMember.SelectedIndex = 0
            End If
        End If

        'Check to make sure Selected Sub-job isn't duped or invalid
        If cmb3SubJob.SelectedIndex > 0 Then
            If cmb3SubJob.SelectedIndex = cmb3PartyMember.SelectedIndex Then
                'Invalid sub-job
                MsgBox(cmb3PartyMember.SelectedItem & " cannot be assigned the " & cmb3SubJob.SelectedItem & " sub-job!", MsgBoxStyle.OkOnly, "Invalid Sub-Job Detected")
                cmb3SubJob.SelectedIndex = 0
            ElseIf cmb3SubJob.SelectedIndex = cmb1SubJob.SelectedIndex Or cmb3SubJob.SelectedIndex = cmb2SubJob.SelectedIndex Or cmb3SubJob.SelectedIndex = cmb4SubJob.SelectedIndex Then
                'Duplicate sub-job
                MsgBox("Another character already has the " & cmb3SubJob.SelectedItem & " sub-job equipped!", MsgBoxStyle.OkOnly, "Duplicate Sub-Job Detected")
                cmb3SubJob.SelectedIndex = 0
            End If
        End If

        'Clear all the stuffs!
        chk3Axe.Checked = False
        chk3Bow.Checked = False
        chk3Dagger.Checked = False
        chk3Dark.Checked = False
        chk3Fire.Checked = False
        chk3Ice.Checked = False
        chk3Lance.Checked = False
        chk3Light.Checked = False
        chk3Staff.Checked = False
        chk3Sword.Checked = False
        chk3Thunder.Checked = False
        chk3Wind.Checked = False
        txt3Stats.Text = String.Empty
        pic3.Image = Nothing
        lbl3Path.BackColor = Nothing
        lbl3Path.Text = Nothing

        'Change Party Member 3 pic+path+primary proficiencies to the selected person (or none)
        Select Case cmb3PartyMember.SelectedIndex
            Case 1
                pic3.Image = OctoPlan.My.Resources.Ophelia
                lbl3Path.BackColor = Color.Tomato
                lbl3Path.Text = "Guide"
                chk3Staff.Checked = True
                chk3Light.Checked = True
            Case 2
                pic3.Image = OctoPlan.My.Resources.Cyrus
                lbl3Path.BackColor = Color.MediumSlateBlue
                lbl3Path.Text = "Scrutinize"
                chk3Staff.Checked = True
                chk3Ice.Checked = True
                chk3Fire.Checked = True
                chk3Thunder.Checked = True
            Case 3
                pic3.Image = OctoPlan.My.Resources.Tressa
                lbl3Path.BackColor = Color.Yellow
                lbl3Path.Text = "Purchase"
                chk3Lance.Checked = True
                chk3Bow.Checked = True
                chk3Wind.Checked = True
            Case 4
                pic3.Image = OctoPlan.My.Resources.Olberic
                lbl3Path.BackColor = Color.LightGreen
                lbl3Path.Text = "Challenge"
                chk3Sword.Checked = True
                chk3Lance.Checked = True
            Case 5
                pic3.Image = OctoPlan.My.Resources.Primrose
                lbl3Path.BackColor = Color.Tomato
                lbl3Path.Text = "Allure"
                chk3Dagger.Checked = True
                chk3Dark.Checked = True
            Case 6
                pic3.Image = OctoPlan.My.Resources.Alfyn
                lbl3Path.BackColor = Color.MediumSlateBlue
                lbl3Path.Text = "Inquire"
                chk3Axe.Checked = True
                chk3Ice.Checked = True
            Case 7
                pic3.Image = OctoPlan.My.Resources.Therion
                lbl3Path.BackColor = Color.Yellow
                lbl3Path.Text = "Steal"
                chk3Sword.Checked = True
                chk3Dagger.Checked = True
                chk3Fire.Checked = True
            Case 8
                pic3.Image = OctoPlan.My.Resources.H_aanit
                lbl3Path.BackColor = Color.LightGreen
                lbl3Path.Text = "Provoke"
                chk3Axe.Checked = True
                chk3Bow.Checked = True
                chk3Thunder.Checked = True
        End Select

        'Set additional proficiencies + bonus stats based on sub-job chosen
        Select Case cmb3SubJob.SelectedIndex
            Case 1 'Cleric
                chk3Staff.Checked = True
                chk3Light.Checked = True
                txt3Stats.Text = "SP: +17.78%" & vbNewLine & "El. Attack: +4.40%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 26.35%"
            Case 2 'Scholar
                chk3Staff.Checked = True
                chk3Ice.Checked = True
                chk3Fire.Checked = True
                chk3Thunder.Checked = True
                txt3Stats.Text = "SP: +11.11%" & vbNewLine & "El. Attack: +9.89%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 25.17%"
            Case 3 'Merchant
                chk3Lance.Checked = True
                chk3Bow.Checked = True
                chk3Wind.Checked = True
                txt3Stats.Text = "HP: +7.86%" & vbNewLine & "SP: +6.67%" & vbNewLine & "Phys Attack: +2.75%" & vbNewLine & "Phys Defense: +2.78%" & vbNewLine & "El. Attack: +2.20%" & vbNewLine & "El. Defense: +2.78%" & vbNewLine & "Total Increase: 25.04%"
            Case 4 'Warrior
                chk3Sword.Checked = True
                chk3Lance.Checked = True
                txt3Stats.Text = "HP: +17.83%" & vbNewLine & "Phys Attack: +4.59%" & vbNewLine & "Phys Defense: +4.17%" & vbNewLine & "Total Increase: 26.59%"
            Case 5 'Dancer
                chk3Dagger.Checked = True
                chk3Dark.Checked = True
                txt3Stats.Text = "El. Attack: +7.69%" & vbNewLine & "Speed: +9.89%" & vbNewLine & "Evasion: +10%" & vbNewLine & "Total Increase: 27.58%"
            Case 6 'Apothecary
                chk3Axe.Checked = True
                chk3Ice.Checked = True
                txt3Stats.Text = "HP: +19.96%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 26.28%"
            Case 7 'Thief
                chk3Sword.Checked = True
                chk3Dagger.Checked = True
                chk3Fire.Checked = True
                txt3Stats.Text = "Phys Attack: +1.83%" & vbNewLine & "Accuracy: +4.59%" & vbNewLine & "Speed: +7.69%" & vbNewLine & "Crit: +5%" & vbNewLine & "Evade: +8%" & vbNewLine & "Total Increase: 27.11%"
            Case 8 'Hunter
                chk3Axe.Checked = True
                chk3Bow.Checked = True
                chk3Thunder.Checked = True
                txt3Stats.Text = "Phys Attack: +7.84%" & vbNewLine & "Accuracy: +5.38%" & vbNewLine & "Speed: +4.05%" & vbNewLine & "Crit: +5.38%" & vbNewLine & "Evade: +2.41%" & vbNewLine & "Total Increase: 25.06%"
            Case 9 'Warmaster
                chk3Sword.Checked = True
                chk3Lance.Checked = True
                chk3Dagger.Checked = True
                chk3Axe.Checked = True
                chk3Bow.Checked = True
                chk3Staff.Checked = True
                txt3Stats.Text = "Phys Attack: +11.93%" & vbNewLine & "Phys Defense: +11.11%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Speed: +5.49%" & vbNewLine & "Evade: +4%" & vbNewLine & "Total Increase: 38.03%"
            Case 10 'Runelord
                chk3Sword.Checked = True
                chk3Axe.Checked = True
                chk3Ice.Checked = True
                chk3Fire.Checked = True
                chk3Thunder.Checked = True
                chk3Wind.Checked = True
                chk3Dark.Checked = True
                chk3Light.Checked = True
                txt3Stats.Text = "Phys Attack: +6.42%" & vbNewLine & "Phys Defense: +5.56%" & vbNewLine & "El. Attack: +6.59%" & vbNewLine & "El. Defense: +5.56%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Crit: +8%" & vbNewLine & "Total Increase: 37.63%"
            Case 11 'Starseer
                chk3Lance.Checked = True
                chk3Dagger.Checked = True
                chk3Wind.Checked = True
                chk3Dark.Checked = True
                chk3Light.Checked = True
                txt3Stats.Text = "HP: +14.86%" & vbNewLine & "SP: +13.33%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "El. Defense: +1.39%" & vbNewLine & "Evade: +2%" & vbNewLine & "Total Increase: 35.90%"
            Case 12 'Sorceror
                chk3Bow.Checked = True
                chk3Staff.Checked = True
                chk3Ice.Checked = True
                chk3Fire.Checked = True
                chk3Thunder.Checked = True
                chk3Wind.Checked = True
                chk3Dark.Checked = True
                chk3Light.Checked = True
                txt3Stats.Text = "SP: +20%" & vbNewLine & "El. Attack: +8.79%" & vbNewLine & "El. Defense: +8.33%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 39.12%"
        End Select
    End Sub

    Private Sub cmb4PartyMember_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb4PartyMember.SelectedIndexChanged, cmb4SubJob.SelectedIndexChanged
        'Check to make sure there aren't duplicate party members
        If cmb4PartyMember.SelectedIndex > 0 Then
            If cmb4PartyMember.SelectedIndex = cmb1PartyMember.SelectedIndex Or cmb4PartyMember.SelectedIndex = cmb2PartyMember.SelectedIndex Or cmb4PartyMember.SelectedIndex = cmb3PartyMember.SelectedIndex Then
                'Here there be Dupes
                MsgBox("You can't have " & cmb4PartyMember.SelectedItem & " in your party more than once!", MsgBoxStyle.OkOnly, "Duplicate Member Detected")
                cmb4PartyMember.SelectedIndex = 0
            End If
        End If

        'Check to make sure Selected Sub-job isn't duped or invalid
        If cmb4SubJob.SelectedIndex > 0 Then
            If cmb4SubJob.SelectedIndex = cmb4PartyMember.SelectedIndex Then
                'Invalid sub-job
                MsgBox(cmb4PartyMember.SelectedItem & " cannot be assigned the " & cmb4SubJob.SelectedItem & " sub-job!", MsgBoxStyle.OkOnly, "Invalid Sub-Job Detected")
                cmb4SubJob.SelectedIndex = 0
            ElseIf cmb4SubJob.SelectedIndex = cmb1SubJob.SelectedIndex Or cmb4SubJob.SelectedIndex = cmb2SubJob.SelectedIndex Or cmb4SubJob.SelectedIndex = cmb3SubJob.SelectedIndex Then
                'Duplicate sub-job
                MsgBox("Another character already has the " & cmb4SubJob.SelectedItem & " sub-job equipped!", MsgBoxStyle.OkOnly, "Duplicate Sub-Job Detected")
                cmb4SubJob.SelectedIndex = 0
            End If
        End If

        'Clear all the stuffs!
        chk4Axe.Checked = False
        chk4Bow.Checked = False
        chk4Dagger.Checked = False
        chk4Dark.Checked = False
        chk4Fire.Checked = False
        chk4Ice.Checked = False
        chk4Lance.Checked = False
        chk4Light.Checked = False
        chk4Staff.Checked = False
        chk4Sword.Checked = False
        chk4Thunder.Checked = False
        chk4Wind.Checked = False
        txt4Stats.Text = String.Empty
        pic4.Image = Nothing
        lbl4Path.BackColor = Nothing
        lbl4Path.Text = Nothing

        'Change Party Member 4 pic+path+primary proficiencies to the selected person (or none)
        Select Case cmb4PartyMember.SelectedIndex
            Case 1
                pic4.Image = OctoPlan.My.Resources.Ophelia
                lbl4Path.BackColor = Color.Tomato
                lbl4Path.Text = "Guide"
                chk4Staff.Checked = True
                chk4Light.Checked = True
            Case 2
                pic4.Image = OctoPlan.My.Resources.Cyrus
                lbl4Path.BackColor = Color.MediumSlateBlue
                lbl4Path.Text = "Scrutinize"
                chk4Staff.Checked = True
                chk4Ice.Checked = True
                chk4Fire.Checked = True
                chk4Thunder.Checked = True
            Case 3
                pic4.Image = OctoPlan.My.Resources.Tressa
                lbl4Path.BackColor = Color.Yellow
                lbl4Path.Text = "Purchase"
                chk4Lance.Checked = True
                chk4Bow.Checked = True
                chk4Wind.Checked = True
            Case 4
                pic4.Image = OctoPlan.My.Resources.Olberic
                lbl4Path.BackColor = Color.LightGreen
                lbl4Path.Text = "Challenge"
                chk4Sword.Checked = True
                chk4Lance.Checked = True
            Case 5
                pic4.Image = OctoPlan.My.Resources.Primrose
                lbl4Path.BackColor = Color.Tomato
                lbl4Path.Text = "Allure"
                chk4Dagger.Checked = True
                chk4Dark.Checked = True
            Case 6
                pic4.Image = OctoPlan.My.Resources.Alfyn
                lbl4Path.BackColor = Color.MediumSlateBlue
                lbl4Path.Text = "Inquire"
                chk4Axe.Checked = True
                chk4Ice.Checked = True
            Case 7
                pic4.Image = OctoPlan.My.Resources.Therion
                lbl4Path.BackColor = Color.Yellow
                lbl4Path.Text = "Steal"
                chk4Sword.Checked = True
                chk4Dagger.Checked = True
                chk4Fire.Checked = True
            Case 8
                pic4.Image = OctoPlan.My.Resources.H_aanit
                lbl4Path.BackColor = Color.LightGreen
                lbl4Path.Text = "Provoke"
                chk4Axe.Checked = True
                chk4Bow.Checked = True
                chk4Thunder.Checked = True
        End Select

        'Set additional proficiencies + bonus stats based on sub-job chosen
        Select Case cmb4SubJob.SelectedIndex
            Case 1 'Cleric
                chk4Staff.Checked = True
                chk4Light.Checked = True
                txt4Stats.Text = "SP: +17.78%" & vbNewLine & "El. Attack: +4.40%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 26.35%"
            Case 2 'Scholar
                chk4Staff.Checked = True
                chk4Ice.Checked = True
                chk4Fire.Checked = True
                chk4Thunder.Checked = True
                txt4Stats.Text = "SP: +11.11%" & vbNewLine & "El. Attack: +9.89%" & vbNewLine & "El. Defense: +4.17%" & vbNewLine & "Total Increase: 25.17%"
            Case 3 'Merchant
                chk4Lance.Checked = True
                chk4Bow.Checked = True
                chk4Wind.Checked = True
                txt4Stats.Text = "HP: +7.86%" & vbNewLine & "SP: +6.67%" & vbNewLine & "Phys Attack: +2.75%" & vbNewLine & "Phys Defense: +2.78%" & vbNewLine & "El. Attack: +2.20%" & vbNewLine & "El. Defense: +2.78%" & vbNewLine & "Total Increase: 25.04%"
            Case 4 'Warrior
                chk4Sword.Checked = True
                chk4Lance.Checked = True
                txt4Stats.Text = "HP: +17.83%" & vbNewLine & "Phys Attack: +4.59%" & vbNewLine & "Phys Defense: +4.17%" & vbNewLine & "Total Increase: 26.59%"
            Case 5 'Dancer
                chk4Dagger.Checked = True
                chk4Dark.Checked = True
                txt4Stats.Text = "El. Attack: +7.69%" & vbNewLine & "Speed: +9.89%" & vbNewLine & "Evasion: +10%" & vbNewLine & "Total Increase: 27.58%"
            Case 6 'Apothecary
                chk4Axe.Checked = True
                chk4Ice.Checked = True
                txt4Stats.Text = "HP: +19.96%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 26.28%"
            Case 7 'Thief
                chk4Sword.Checked = True
                chk4Dagger.Checked = True
                chk4Fire.Checked = True
                txt4Stats.Text = "Phys Attack: +1.83%" & vbNewLine & "Accuracy: +4.59%" & vbNewLine & "Speed: +7.69%" & vbNewLine & "Crit: +5%" & vbNewLine & "Evade: +8%" & vbNewLine & "Total Increase: 27.11%"
            Case 8 'Hunter
                chk4Axe.Checked = True
                chk4Bow.Checked = True
                chk4Thunder.Checked = True
                txt4Stats.Text = "Phys Attack: +7.84%" & vbNewLine & "Accuracy: +5.38%" & vbNewLine & "Speed: +4.05%" & vbNewLine & "Crit: +5.38%" & vbNewLine & "Evade: +2.41%" & vbNewLine & "Total Increase: 25.06%"
            Case 9 'Warmaster
                chk4Sword.Checked = True
                chk4Lance.Checked = True
                chk4Dagger.Checked = True
                chk4Axe.Checked = True
                chk4Bow.Checked = True
                chk4Staff.Checked = True
                txt4Stats.Text = "Phys Attack: +11.93%" & vbNewLine & "Phys Defense: +11.11%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Speed: +5.49%" & vbNewLine & "Evade: +4%" & vbNewLine & "Total Increase: 38.03%"
            Case 10 'Runelord
                chk4Sword.Checked = True
                chk4Axe.Checked = True
                chk4Ice.Checked = True
                chk4Fire.Checked = True
                chk4Thunder.Checked = True
                chk4Wind.Checked = True
                chk4Dark.Checked = True
                chk4Light.Checked = True
                txt4Stats.Text = "Phys Attack: +6.42%" & vbNewLine & "Phys Defense: +5.56%" & vbNewLine & "El. Attack: +6.59%" & vbNewLine & "El. Defense: +5.56%" & vbNewLine & "Accuracy: +5.50%" & vbNewLine & "Crit: +8%" & vbNewLine & "Total Increase: 37.63%"
            Case 11 'Starseer
                chk4Lance.Checked = True
                chk4Dagger.Checked = True
                chk4Wind.Checked = True
                chk4Dark.Checked = True
                chk4Light.Checked = True
                txt4Stats.Text = "HP: +14.86%" & vbNewLine & "SP: +13.33%" & vbNewLine & "Phys Attack: +1.83%" & vbNewLine & "Phys Defense: +1.39%" & vbNewLine & "El. Attack: +1.10%" & vbNewLine & "El. Defense: +1.39%" & vbNewLine & "Evade: +2%" & vbNewLine & "Total Increase: 35.90%"
            Case 12 'Sorceror
                chk4Bow.Checked = True
                chk4Staff.Checked = True
                chk4Ice.Checked = True
                chk4Fire.Checked = True
                chk4Thunder.Checked = True
                chk4Wind.Checked = True
                chk4Dark.Checked = True
                chk4Light.Checked = True
                txt4Stats.Text = "SP: +20%" & vbNewLine & "El. Attack: +8.79%" & vbNewLine & "El. Defense: +8.33%" & vbNewLine & "Crit: +2%" & vbNewLine & "Total Increase: 39.12%"
        End Select
    End Sub

    Sub Flush_Array()
        For intOuterIncrementor As Integer = 1 To 12
            For intInnerIncrementor As Integer = 1 To 4
                intPassivesArray(intOuterIncrementor, intInnerIncrementor) = 0
            Next
        Next
        intRunningTotal = 0
    End Sub

    Sub Check_For_Duplicate_Passives(intSkill1 As Integer, intSkill2 As Integer, intSkill3 As Integer, intSkill4 As Integer)
        If intSkill1 = intSkill2 And intSkill1 > 0 Then
            blnDupes = True
        ElseIf intSkill1 = intSkill3 And intSkill1 > 0 Then
            blnDupes = True
        ElseIf intSkill1 = intSkill4 And intSkill1 > 0 Then
            blnDupes = True
        ElseIf intSkill2 = intSkill3 And intSkill2 > 0 Then
            blnDupes = True
        ElseIf intSkill2 = intSkill4 And intSkill2 > 0 Then
            blnDupes = True
        ElseIf intSkill3 = intSkill4 And intSkill3 > 0 Then
            blnDupes = True
        Else
            blnDupes = False
        End If
    End Sub

    Sub Populate_Array(intSkill1 As Integer, intSkill2 As Integer, intSkill3 As Integer, intSkill4 As Integer)
        Select Case intSkill1
            Case 1
                intPassivesArray(8, 4) = 2000
            Case 2
                intPassivesArray(9, 2) = 2000
            Case 3
                intPassivesArray(9, 3) = 2000
            Case 4
                intPassivesArray(12, 1) = 130
            Case 5
                intPassivesArray(6, 3) = 2000
            Case 6
                intPassivesArray(9, 4) = 2000
            Case 7
                intPassivesArray(4, 2) = 500
            Case 8
                intPassivesArray(8, 3) = 2000
            Case 9
                intPassivesArray(7, 2) = 500
            Case 10
                intPassivesArray(6, 4) = 2000
            Case 11
                intPassivesArray(3, 4) = 3000
            Case 12
                intPassivesArray(5, 1) = 130
            Case 13
                intPassivesArray(12, 3) = 1000
            Case 14
                intPassivesArray(7, 1) = 130
            Case 15
                intPassivesArray(2, 3) = 1000
            Case 16
                intPassivesArray(11, 1) = 4000
            Case 17
                intPassivesArray(3, 2) = 500
            Case 18
                intPassivesArray(10, 2) = 500
            Case 19
                intPassivesArray(11, 3) = 2000
            Case 20
                intPassivesArray(5, 2) = 500
            Case 21
                intPassivesArray(1, 2) = 500
            Case 22
                intPassivesArray(5, 3) = 1000
            Case 23
                intPassivesArray(9, 1) = 4000
            Case 24
                intPassivesArray(4, 1) = 130
            Case 25
                intPassivesArray(1, 4) = 3000
            Case 26
                intPassivesArray(10, 1) = 130
            Case 27
                intPassivesArray(2, 2) = 500
            Case 28
                intPassivesArray(1, 1) = 130
            Case 29
                intPassivesArray(10, 4) = 3000
            Case 30
                intPassivesArray(8, 1) = 4000
            Case 31
                intPassivesArray(4, 4) = 3000
            Case 32
                intPassivesArray(7, 3) = 1000
            Case 33
                intPassivesArray(2, 1) = 130
            Case 34
                intPassivesArray(11, 4) = 2000
            Case 35
                intPassivesArray(1, 3) = 1000
            Case 36
                intPassivesArray(2, 4) = 3000
            Case 37
                intPassivesArray(4, 3) = 1000
            Case 38
                intPassivesArray(3, 3) = 1000
            Case 39
                intPassivesArray(10, 3) = 1000
            Case 40
                intPassivesArray(6, 2) = 2000
            Case 41
                intPassivesArray(5, 4) = 3000
            Case 42
                intPassivesArray(11, 2) = 2000
            Case 43
                intPassivesArray(6, 1) = 4000
            Case 44
                intPassivesArray(8, 2) = 2000
            Case 45
                intPassivesArray(12, 2) = 500
            Case 46
                intPassivesArray(12, 4) = 3000
            Case 47
                intPassivesArray(3, 1) = 130
            Case 48
                intPassivesArray(7, 4) = 3000
        End Select

        Select Case intSkill2
            Case 1
                intPassivesArray(8, 4) = 2000
            Case 2
                intPassivesArray(9, 2) = 2000
            Case 3
                intPassivesArray(9, 3) = 2000
            Case 4
                intPassivesArray(12, 1) = 130
            Case 5
                intPassivesArray(6, 3) = 2000
            Case 6
                intPassivesArray(9, 4) = 2000
            Case 7
                intPassivesArray(4, 2) = 500
            Case 8
                intPassivesArray(8, 3) = 2000
            Case 9
                intPassivesArray(7, 2) = 500
            Case 10
                intPassivesArray(6, 4) = 2000
            Case 11
                intPassivesArray(3, 4) = 3000
            Case 12
                intPassivesArray(5, 1) = 130
            Case 13
                intPassivesArray(12, 3) = 1000
            Case 14
                intPassivesArray(7, 1) = 130
            Case 15
                intPassivesArray(2, 3) = 1000
            Case 16
                intPassivesArray(11, 1) = 4000
            Case 17
                intPassivesArray(3, 2) = 500
            Case 18
                intPassivesArray(10, 2) = 500
            Case 19
                intPassivesArray(11, 3) = 2000
            Case 20
                intPassivesArray(5, 2) = 500
            Case 21
                intPassivesArray(1, 2) = 500
            Case 22
                intPassivesArray(5, 3) = 1000
            Case 23
                intPassivesArray(9, 1) = 4000
            Case 24
                intPassivesArray(4, 1) = 130
            Case 25
                intPassivesArray(1, 4) = 3000
            Case 26
                intPassivesArray(10, 1) = 130
            Case 27
                intPassivesArray(2, 2) = 500
            Case 28
                intPassivesArray(1, 1) = 130
            Case 29
                intPassivesArray(10, 4) = 3000
            Case 30
                intPassivesArray(8, 1) = 4000
            Case 31
                intPassivesArray(4, 4) = 3000
            Case 32
                intPassivesArray(7, 3) = 1000
            Case 33
                intPassivesArray(2, 1) = 130
            Case 34
                intPassivesArray(11, 4) = 2000
            Case 35
                intPassivesArray(1, 3) = 1000
            Case 36
                intPassivesArray(2, 4) = 3000
            Case 37
                intPassivesArray(4, 3) = 1000
            Case 38
                intPassivesArray(3, 3) = 1000
            Case 39
                intPassivesArray(10, 3) = 1000
            Case 40
                intPassivesArray(6, 2) = 2000
            Case 41
                intPassivesArray(5, 4) = 3000
            Case 42
                intPassivesArray(11, 2) = 2000
            Case 43
                intPassivesArray(6, 1) = 4000
            Case 44
                intPassivesArray(8, 2) = 2000
            Case 45
                intPassivesArray(12, 2) = 500
            Case 46
                intPassivesArray(12, 4) = 3000
            Case 47
                intPassivesArray(3, 1) = 130
            Case 48
                intPassivesArray(7, 4) = 3000
        End Select

        Select Case intSkill3
            Case 1
                intPassivesArray(8, 4) = 2000
            Case 2
                intPassivesArray(9, 2) = 2000
            Case 3
                intPassivesArray(9, 3) = 2000
            Case 4
                intPassivesArray(12, 1) = 130
            Case 5
                intPassivesArray(6, 3) = 2000
            Case 6
                intPassivesArray(9, 4) = 2000
            Case 7
                intPassivesArray(4, 2) = 500
            Case 8
                intPassivesArray(8, 3) = 2000
            Case 9
                intPassivesArray(7, 2) = 500
            Case 10
                intPassivesArray(6, 4) = 2000
            Case 11
                intPassivesArray(3, 4) = 3000
            Case 12
                intPassivesArray(5, 1) = 130
            Case 13
                intPassivesArray(12, 3) = 1000
            Case 14
                intPassivesArray(7, 1) = 130
            Case 15
                intPassivesArray(2, 3) = 1000
            Case 16
                intPassivesArray(11, 1) = 4000
            Case 17
                intPassivesArray(3, 2) = 500
            Case 18
                intPassivesArray(10, 2) = 500
            Case 19
                intPassivesArray(11, 3) = 2000
            Case 20
                intPassivesArray(5, 2) = 500
            Case 21
                intPassivesArray(1, 2) = 500
            Case 22
                intPassivesArray(5, 3) = 1000
            Case 23
                intPassivesArray(9, 1) = 4000
            Case 24
                intPassivesArray(4, 1) = 130
            Case 25
                intPassivesArray(1, 4) = 3000
            Case 26
                intPassivesArray(10, 1) = 130
            Case 27
                intPassivesArray(2, 2) = 500
            Case 28
                intPassivesArray(1, 1) = 130
            Case 29
                intPassivesArray(10, 4) = 3000
            Case 30
                intPassivesArray(8, 1) = 4000
            Case 31
                intPassivesArray(4, 4) = 3000
            Case 32
                intPassivesArray(7, 3) = 1000
            Case 33
                intPassivesArray(2, 1) = 130
            Case 34
                intPassivesArray(11, 4) = 2000
            Case 35
                intPassivesArray(1, 3) = 1000
            Case 36
                intPassivesArray(2, 4) = 3000
            Case 37
                intPassivesArray(4, 3) = 1000
            Case 38
                intPassivesArray(3, 3) = 1000
            Case 39
                intPassivesArray(10, 3) = 1000
            Case 40
                intPassivesArray(6, 2) = 2000
            Case 41
                intPassivesArray(5, 4) = 3000
            Case 42
                intPassivesArray(11, 2) = 2000
            Case 43
                intPassivesArray(6, 1) = 4000
            Case 44
                intPassivesArray(8, 2) = 2000
            Case 45
                intPassivesArray(12, 2) = 500
            Case 46
                intPassivesArray(12, 4) = 3000
            Case 47
                intPassivesArray(3, 1) = 130
            Case 48
                intPassivesArray(7, 4) = 3000
        End Select

        Select Case intSkill4
            Case 1
                intPassivesArray(8, 4) = 2000
            Case 2
                intPassivesArray(9, 2) = 2000
            Case 3
                intPassivesArray(9, 3) = 2000
            Case 4
                intPassivesArray(12, 1) = 130
            Case 5
                intPassivesArray(6, 3) = 2000
            Case 6
                intPassivesArray(9, 4) = 2000
            Case 7
                intPassivesArray(4, 2) = 500
            Case 8
                intPassivesArray(8, 3) = 2000
            Case 9
                intPassivesArray(7, 2) = 500
            Case 10
                intPassivesArray(6, 4) = 2000
            Case 11
                intPassivesArray(3, 4) = 3000
            Case 12
                intPassivesArray(5, 1) = 130
            Case 13
                intPassivesArray(12, 3) = 1000
            Case 14
                intPassivesArray(7, 1) = 130
            Case 15
                intPassivesArray(2, 3) = 1000
            Case 16
                intPassivesArray(11, 1) = 4000
            Case 17
                intPassivesArray(3, 2) = 500
            Case 18
                intPassivesArray(10, 2) = 500
            Case 19
                intPassivesArray(11, 3) = 2000
            Case 20
                intPassivesArray(5, 2) = 500
            Case 21
                intPassivesArray(1, 2) = 500
            Case 22
                intPassivesArray(5, 3) = 1000
            Case 23
                intPassivesArray(9, 1) = 4000
            Case 24
                intPassivesArray(4, 1) = 130
            Case 25
                intPassivesArray(1, 4) = 3000
            Case 26
                intPassivesArray(10, 1) = 130
            Case 27
                intPassivesArray(2, 2) = 500
            Case 28
                intPassivesArray(1, 1) = 130
            Case 29
                intPassivesArray(10, 4) = 3000
            Case 30
                intPassivesArray(8, 1) = 4000
            Case 31
                intPassivesArray(4, 4) = 3000
            Case 32
                intPassivesArray(7, 3) = 1000
            Case 33
                intPassivesArray(2, 1) = 130
            Case 34
                intPassivesArray(11, 4) = 2000
            Case 35
                intPassivesArray(1, 3) = 1000
            Case 36
                intPassivesArray(2, 4) = 3000
            Case 37
                intPassivesArray(4, 3) = 1000
            Case 38
                intPassivesArray(3, 3) = 1000
            Case 39
                intPassivesArray(10, 3) = 1000
            Case 40
                intPassivesArray(6, 2) = 2000
            Case 41
                intPassivesArray(5, 4) = 3000
            Case 42
                intPassivesArray(11, 2) = 2000
            Case 43
                intPassivesArray(6, 1) = 4000
            Case 44
                intPassivesArray(8, 2) = 2000
            Case 45
                intPassivesArray(12, 2) = 500
            Case 46
                intPassivesArray(12, 4) = 3000
            Case 47
                intPassivesArray(3, 1) = 130
            Case 48
                intPassivesArray(7, 4) = 3000
        End Select
    End Sub

    Sub Sum_Array()
        Dim intIncrementor As Integer
        For intIncrementor = 1 To 12
            'Check to see if multiple skills are from same job
            If intPassivesArray(intIncrementor, 1) = 0 And intPassivesArray(intIncrementor, 2) = 0 And intPassivesArray(intIncrementor, 3) = 0 And intPassivesArray(intIncrementor, 4) = 0 Then
                'No skills from this job, do nothing
            ElseIf intPassivesArray(intIncrementor, 4) > 0 Then
                'Tier-4 support was highest picked, add full cumulative cost of all 4 tiers to running total
                Select Case intPassivesArray(intIncrementor, 4)
                    Case 3000 'Basic
                        intRunningTotal += 4630
                    Case 2000 'Advanced
                        intRunningTotal += 10000
                End Select
            ElseIf intPassivesArray(intIncrementor, 3) > 0 Then
                'Tier-3 support was highest picked, add full cumulative cost of first 3 tiers to running total
                Select Case intPassivesArray(intIncrementor, 3)
                    Case 1000 'Basic
                        intRunningTotal += 1630
                    Case 2000 'Advanced
                        intRunningTotal += 8000
                End Select
            ElseIf intPassivesArray(intIncrementor, 2) > 0 Then
                'Tier-2 support was highest picked, add full cumulative cost of first 2 tiers to running total
                Select Case intPassivesArray(intIncrementor, 2)
                    Case 500 'Basic
                        intRunningTotal += 630
                    Case 2000 'Advanced
                        intRunningTotal += 6000
                End Select
            Else
                'Tier-1 support was the only one picked, add its cost to running total
                intRunningTotal += intPassivesArray(intIncrementor, 1)
            End If
        Next
    End Sub

    Private Sub cmb1Passive1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb1Passive1.SelectedIndexChanged, cmb1Passive2.SelectedIndexChanged, cmb1Passive3.SelectedIndexChanged, cmb1Passive4.SelectedIndexChanged
        'Check for duplicate passives
        Check_For_Duplicate_Passives(cmb1Passive1.SelectedIndex, cmb1Passive2.SelectedIndex, cmb1Passive3.SelectedIndex, cmb1Passive4.SelectedIndex)
        If blnDupes = True Then
            MsgBox("This character already has " & sender.selecteditem & " equipped!", MsgBoxStyle.OkOnly, "Duplicate Passive Detected")
            sender.selectedindex = 0
        End If

        'Reset the JP costs
        Flush_Array()

        'Set new JP costs
        Populate_Array(cmb1Passive1.SelectedIndex, cmb1Passive2.SelectedIndex, cmb1Passive3.SelectedIndex, cmb1Passive4.SelectedIndex)

        'Add the new entry's JP cost, if any, to the running total
        Sum_Array()

        'Update new JP requirement to user
        lbl1JP.Text = intRunningTotal.ToString("N0")
    End Sub

    Private Sub cmb2Passive1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb2Passive1.SelectedIndexChanged, cmb2Passive2.SelectedIndexChanged, cmb2Passive3.SelectedIndexChanged, cmb2Passive4.SelectedIndexChanged
        'Check for duplicate passives
        Check_For_Duplicate_Passives(cmb2Passive1.SelectedIndex, cmb2Passive2.SelectedIndex, cmb2Passive3.SelectedIndex, cmb2Passive4.SelectedIndex)
        If blnDupes = True Then
            MsgBox("This character already has " & sender.selecteditem & " equipped!", MsgBoxStyle.OkOnly, "Duplicate Passive Detected")
            sender.selectedindex = 0
        End If

        'Reset the JP costs
        Flush_Array()

        'Set new JP costs
        Populate_Array(cmb2Passive1.SelectedIndex, cmb2Passive2.SelectedIndex, cmb2Passive3.SelectedIndex, cmb2Passive4.SelectedIndex)

        'Add the new entry's JP cost, if any, to the running total
        Sum_Array()

        'Update new JP requirement to user
        lbl2JP.Text = intRunningTotal.ToString("N0")
    End Sub

    Private Sub cmb3Passive1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb3Passive1.SelectedIndexChanged, cmb3Passive2.SelectedIndexChanged, cmb3Passive3.SelectedIndexChanged, cmb3Passive4.SelectedIndexChanged
        'Check for duplicate passives
        Check_For_Duplicate_Passives(cmb3Passive1.SelectedIndex, cmb3Passive2.SelectedIndex, cmb3Passive3.SelectedIndex, cmb3Passive4.SelectedIndex)
        If blnDupes = True Then
            MsgBox("This character already has " & sender.selecteditem & " equipped!", MsgBoxStyle.OkOnly, "Duplicate Passive Detected")
            sender.selectedindex = 0
        End If

        'Reset the JP costs
        Flush_Array()

        'Set new JP costs
        Populate_Array(cmb3Passive1.SelectedIndex, cmb3Passive2.SelectedIndex, cmb3Passive3.SelectedIndex, cmb3Passive4.SelectedIndex)

        'Add the new entry's JP cost, if any, to the running total
        Sum_Array()

        'Update new JP requirement to user
        lbl3JP.Text = intRunningTotal.ToString("N0")
    End Sub

    Private Sub cmb4Passive1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb4Passive1.SelectedIndexChanged, cmb4Passive2.SelectedIndexChanged, cmb4Passive3.SelectedIndexChanged, cmb4Passive4.SelectedIndexChanged
        'Check for duplicate passives
        Check_For_Duplicate_Passives(cmb4Passive1.SelectedIndex, cmb4Passive2.SelectedIndex, cmb4Passive3.SelectedIndex, cmb4Passive4.SelectedIndex)
        If blnDupes = True Then
            MsgBox("This character already has " & sender.selecteditem & " equipped!", MsgBoxStyle.OkOnly, "Duplicate Passive Detected")
            sender.selectedindex = 0
        End If

        'Reset the JP costs
        Flush_Array()

        'Set new JP costs
        Populate_Array(cmb4Passive1.SelectedIndex, cmb4Passive2.SelectedIndex, cmb4Passive3.SelectedIndex, cmb4Passive4.SelectedIndex)

        'Add the new entry's JP cost, if any, to the running total
        Sum_Array()

        'Update new JP requirement to user
        lbl4JP.Text = intRunningTotal.ToString("N0")
    End Sub

    Private Sub mnuFileSave_Click(sender As Object, e As EventArgs) Handles mnuFileSave.Click
        'Set the extension(s) to be used
        SaveFileDialog1.Filter = "OctoPlan File (.opf)|*.opf"

        'See if file save was successful
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Declare a StreamWriter object
            Dim FS As IO.StreamWriter

            'Get the filename
            Dim FileName As String = SaveFileDialog1.FileName

            'Create the file
            FS = IO.File.CreateText(FileName)

            'Write relevant data to the file, starting with character 1
            FS.WriteLine(cmb1PartyMember.SelectedIndex)
            FS.WriteLine(cmb1SubJob.SelectedIndex)
            FS.WriteLine(cmb1Passive1.SelectedIndex)
            FS.WriteLine(cmb1Passive2.SelectedIndex)
            FS.WriteLine(cmb1Passive3.SelectedIndex)
            FS.WriteLine(cmb1Passive4.SelectedIndex)
            'character 2
            FS.WriteLine(cmb2PartyMember.SelectedIndex)
            FS.WriteLine(cmb2SubJob.SelectedIndex)
            FS.WriteLine(cmb2Passive1.SelectedIndex)
            FS.WriteLine(cmb2Passive2.SelectedIndex)
            FS.WriteLine(cmb2Passive3.SelectedIndex)
            FS.WriteLine(cmb2Passive4.SelectedIndex)
            'character 3
            FS.WriteLine(cmb3PartyMember.SelectedIndex)
            FS.WriteLine(cmb3SubJob.SelectedIndex)
            FS.WriteLine(cmb3Passive1.SelectedIndex)
            FS.WriteLine(cmb3Passive2.SelectedIndex)
            FS.WriteLine(cmb3Passive3.SelectedIndex)
            FS.WriteLine(cmb3Passive4.SelectedIndex)
            'character 4
            FS.WriteLine(cmb4PartyMember.SelectedIndex)
            FS.WriteLine(cmb4SubJob.SelectedIndex)
            FS.WriteLine(cmb4Passive1.SelectedIndex)
            FS.WriteLine(cmb4Passive2.SelectedIndex)
            FS.WriteLine(cmb4Passive3.SelectedIndex)
            FS.WriteLine(cmb4Passive4.SelectedIndex)

            'Close the file
            FS.Close()
        End If
    End Sub

    Private Sub mnuFileLoad_Click(sender As Object, e As EventArgs) Handles mnuFileLoad.Click
        OpenFileDialog1.Filter = "OctoPlan File (.opf)|*.opf"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim FileName As String = OpenFileDialog1.FileName
            Dim FL As IO.StreamReader
            Dim strInputArray(24) As String
            Dim intInputArray(24) As Integer



            FL = IO.File.OpenText(FileName)

            For intIncrementor As Integer = 1 To 24
                strInputArray(intIncrementor) = FL.ReadLine
                Integer.TryParse(strInputArray(intIncrementor), intInputArray(intIncrementor))
            Next

            FL.Close()

            'Set character 1 parameters
            cmb1PartyMember.SelectedIndex = intInputArray(1)
            cmb1SubJob.SelectedIndex = intInputArray(2)
            cmb1Passive1.SelectedIndex = intInputArray(3)
            cmb1Passive2.SelectedIndex = intInputArray(4)
            cmb1Passive3.SelectedIndex = intInputArray(5)
            cmb1Passive4.SelectedIndex = intInputArray(6)

            'character 2
            cmb2PartyMember.SelectedIndex = intInputArray(7)
            cmb2SubJob.SelectedIndex = intInputArray(8)
            cmb2Passive1.SelectedIndex = intInputArray(9)
            cmb2Passive2.SelectedIndex = intInputArray(10)
            cmb2Passive3.SelectedIndex = intInputArray(11)
            cmb2Passive4.SelectedIndex = intInputArray(12)

            'character 3
            cmb3PartyMember.SelectedIndex = intInputArray(13)
            cmb3SubJob.SelectedIndex = intInputArray(14)
            cmb3Passive1.SelectedIndex = intInputArray(15)
            cmb3Passive2.SelectedIndex = intInputArray(16)
            cmb3Passive3.SelectedIndex = intInputArray(17)
            cmb3Passive4.SelectedIndex = intInputArray(18)

            'character 4
            cmb4PartyMember.SelectedIndex = intInputArray(19)
            cmb4SubJob.SelectedIndex = intInputArray(20)
            cmb4Passive1.SelectedIndex = intInputArray(21)
            cmb4Passive2.SelectedIndex = intInputArray(22)
            cmb4Passive3.SelectedIndex = intInputArray(23)
            cmb4Passive4.SelectedIndex = intInputArray(24)
        End If
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        AboutBox1.Show()
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        Dim strAnswer As String = MsgBox("Exit OctoPlan? (ALL UNSAVED PROGRESS WILL BE LOST!)", MsgBoxStyle.YesNo, "Annoying yet obligatory confirmation box")

        If strAnswer = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub
End Class
