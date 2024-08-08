Attribute VB_Name = "ModFacciones"
Option Explicit

Public ArmaduraImperial1              As Integer
Public ArmaduraImperial2              As Integer
Public ArmaduraImperial3              As Integer
Public TunicaMagoImperial             As Integer
Public TunicaMagoImperialEnanos       As Integer
Public ArmaduraCaos1                  As Integer
Public ArmaduraCaos2                  As Integer
Public ArmaduraCaos3                  As Integer
Public TunicaMagoCaos                 As Integer
Public TunicaMagoCaosEnanos           As Integer
Public VestimentaImperialHumano       As Integer
Public VestimentaImperialEnano        As Integer
Public TunicaConspicuaHumano          As Integer
Public TunicaConspicuaEnano           As Integer
Public ArmaduraNobilisimaHumano       As Integer
Public ArmaduraNobilisimaEnano        As Integer
Public ArmaduraGranSacerdote          As Integer
Public VestimentaLegionHumano         As Integer
Public VestimentaLegionEnano          As Integer
Public TunicaLobregaHumano            As Integer
Public TunicaLobregaEnano             As Integer
Public TunicaEgregiaHumano            As Integer
Public TunicaEgregiaEnano             As Integer
Public SacerdoteDemoniaco             As Integer
Public Const NUM_RANGOS_FACCION       As Integer = 15
Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras
Public RecompensaFacciones(NUM_RANGOS_FACCION)         As Long

Public Enum eTipoDefArmors
    ieBaja
    ieMedia
    ieAlta
End Enum

Public Type tFaccionArmaduras
    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer
End Type

Private Function GetArmourAmount(ByVal Rango As Integer, ByVal TipoDef As eTipoDefArmors) As Integer
    Select Case TipoDef
        Case eTipoDefArmors.ieBaja
            GetArmourAmount = 20 / (Rango + 1)
            
        Case eTipoDefArmors.ieMedia
            GetArmourAmount = Rango * 2 / MaximoInt((Rango - 4), 1)
            
        Case eTipoDefArmors.ieAlta
            GetArmourAmount = Rango * 1.35
    End Select
End Function

Private Sub GiveFactionArmours(ByVal Userindex As Integer, ByVal IsCaos As Boolean)
    Dim ObjArmour As obj
    Dim Rango     As Integer
    With UserList(Userindex)
        Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Caos(eTipoDefArmors.ieBaja)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Armada(eTipoDefArmors.ieBaja)
        End If
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Caos(eTipoDefArmors.ieMedia)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Armada(eTipoDefArmors.ieMedia)
        End If
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieAlta)
        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Caos(eTipoDefArmors.ieAlta)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.Clase, .raza).Armada(eTipoDefArmors.ieAlta)
        End If
        If Not MeterItemEnInventario(Userindex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If
    End With
End Sub

Public Sub GiveExpReward(ByVal Userindex As Integer, ByVal Rango As Long)
    Dim GivenExp As Long
    With UserList(Userindex)
        GivenExp = RecompensaFacciones(Rango)
        .Stats.Exp = .Stats.Exp + GivenExp
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        Call WriteConsoleMsg(Userindex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
        Call CheckUserLevel(Userindex)
    End With
End Sub

Public Sub EnlistarArmadaReal(ByVal Userindex As Integer)
    With UserList(Userindex)
        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(Userindex, "Ya perteneces a las tropas reales!!! Ve a combatir criminales.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(Userindex, "Maldito insolente!!! Vete de aqui seguidor de las sombras.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "No se permiten criminales en el ejercito real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.CriminalesMatados < 30 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & .Faccion.CriminalesMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.CiudadanosMatados > 0 Then
            Call WriteChatOverHead(Userindex, "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.Reenlistadas > 4 Then
            Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas reales demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Reputacion.NobleRep < 1000000 Then
            Call WriteChatOverHead(Userindex, "Necesitas ser aun mas noble para integrar el ejercito real, solo tienes " & .Reputacion.NobleRep & "/1.000.000 puntos de nobleza", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(Userindex, "Perteneces a un clan neutro, sal de el si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
        .Faccion.ArmadaReal = 1
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        Call WriteChatOverHead(Userindex, "Bienvenido al ejercito real!!! Aqui tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargare de recompensarte.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        If .Faccion.RecibioArmaduraReal = 0 Then
            Call GiveFactionArmours(Userindex, False)
            Call GiveExpReward(Userindex, 0)
            .Faccion.RecibioArmaduraReal = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = DatePart("yyyy", Date) & "-" & DatePart("m", Date) & "-" & DatePart("d", Date)
            .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
            .Faccion.RecibioExpInicialReal = 1
            .Faccion.RecompensasReal = 0
            .Faccion.NextRecompensa = 70
        End If
        If .flags.Navegando Then Call RefreshCharStatus(Userindex)
        Call LogEjercitoReal(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)
    End With
End Sub

Public Sub RecompensaArmadaReal(ByVal Userindex As Integer)
    Dim Crimis    As Long
    Dim Lvl       As Byte
    Dim NextRecom As Long
    Dim Nobleza   As Long
    With UserList(Userindex)
        Lvl = .Stats.ELV
        Crimis = .Faccion.CriminalesMatados
        NextRecom = .Faccion.NextRecompensa
        Nobleza = .Reputacion.NobleRep
        If Crimis < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Crimis & " criminales mas para recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        Select Case NextRecom
            Case 70:
                .Faccion.RecompensasReal = 1
                .Faccion.NextRecompensa = 130
        
            Case 130:
                .Faccion.RecompensasReal = 2
                .Faccion.NextRecompensa = 210
        
            Case 210:
                .Faccion.RecompensasReal = 3
                .Faccion.NextRecompensa = 320
        
            Case 320:
                .Faccion.RecompensasReal = 4
                .Faccion.NextRecompensa = 460
        
            Case 460:
                .Faccion.RecompensasReal = 5
                .Faccion.NextRecompensa = 640
        
            Case 640:
                If Lvl < 27 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 6
                .Faccion.NextRecompensa = 870
        
            Case 870:
                .Faccion.RecompensasReal = 7
                .Faccion.NextRecompensa = 1160
        
            Case 1160:
                .Faccion.RecompensasReal = 8
                .Faccion.NextRecompensa = 2000
        
            Case 2000:
                If Lvl < 30 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 30 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 9
                .Faccion.NextRecompensa = 2500
        
            Case 2500:
                If Nobleza < 2000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 2000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 10
                .Faccion.NextRecompensa = 3000
        
            Case 3000:
                If Nobleza < 3000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 3000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 11
                .Faccion.NextRecompensa = 3500
        
            Case 3500:
                If Lvl < 35 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 35 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 4000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 4000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 12
                .Faccion.NextRecompensa = 4000
        
            Case 4000:
                If Lvl < 36 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 5000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 5000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 13
                .Faccion.NextRecompensa = 5000
        
            Case 5000:
                If Lvl < 37 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 37 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 6000000 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 6000000 - Nobleza & " puntos de nobleza para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 14
                .Faccion.NextRecompensa = 10000
        
            Case 10000:
                Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue asi. Ya no tengo mas recompensa para darte que mi agradecimiento. Felicidades!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
        
            Case Else:
                Exit Sub
        End Select
    
        Call WriteChatOverHead(Userindex, "Aqui tienes tu recompensa " & TituloReal(Userindex) & "!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Call GiveFactionArmours(Userindex, False)
        Call GiveExpReward(Userindex, .Faccion.RecompensasReal)
    End With
End Sub

Public Sub ExpulsarFaccionReal(ByVal Userindex As Integer, Optional Expulsado As Boolean = True)
    With UserList(Userindex)
        .Faccion.ArmadaReal = 0
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        If .Invent.ArmourEqpObjIndex <> 0 Then
            If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
        End If
        If .Invent.EscudoEqpObjIndex <> 0 Then
            If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
        End If
        If .flags.Navegando Then Call RefreshCharStatus(Userindex)
    End With
End Sub

Public Sub ExpulsarFaccionCaos(ByVal Userindex As Integer, Optional Expulsado As Boolean = True)
    With UserList(Userindex)
        .Faccion.FuerzasCaos = 0
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        If .Invent.ArmourEqpObjIndex <> 0 Then
            If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
        End If
        If .Invent.EscudoEqpObjIndex <> 0 Then
            If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
        End If
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
    End With
End Sub

Public Function TituloReal(ByVal Userindex As Integer) As String
    Select Case UserList(Userindex).Faccion.RecompensasReal
        Case 0
            TituloReal = "Aprendiz"

        Case 1
            TituloReal = "Escudero"

        Case 2
            TituloReal = "Soldado"

        Case 3
            TituloReal = "Sargento"

        Case 4
            TituloReal = "Teniente"

        Case 5
            TituloReal = "Comandante"

        Case 6
            TituloReal = "Capitan"

        Case 7
            TituloReal = "Senescal"

        Case 8
            TituloReal = "Mariscal"

        Case 9
            TituloReal = "Condestable"

        Case 10
            TituloReal = "Ejecutor Imperial"

        Case 11
            TituloReal = "Protector del Reino"

        Case 12
            TituloReal = "Avatar de la Justicia"

        Case 13
            TituloReal = "Guardian del Bien"

        Case Else
            TituloReal = "Campeon de la Luz"
    End Select
End Function

Public Sub EnlistarCaos(ByVal Userindex As Integer)
    With UserList(Userindex)
        If Not criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "Largate de aqui, bufon!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(Userindex, "Ya perteneces a la legion oscura!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(Userindex, "Las sombras reinaran en Argentum. Fuera de aqui insecto real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.RecibioExpInicialReal = 1 Then
            Call WriteChatOverHead(Userindex, "No permitire que ningun insecto real ingrese a mis tropas.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If Not criminal(Userindex) Then
            Call WriteChatOverHead(Userindex, "Ja ja ja!! Tu no eres bienvenido aqui asqueroso ciudadano.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Faccion.CiudadanosMatados < 70 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 70 ciudadanos, solo has matado " & .Faccion.CiudadanosMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(Userindex, "Perteneces a un clan neutro, sal de el si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
        If .Faccion.Reenlistadas > 4 Then
            If .Faccion.Reenlistadas = 200 Then
                Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldia has atacado a mi ejercito. Vete de aqui!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
                Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas oscuras demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            End If
            Exit Sub
        End If
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.FuerzasCaos = 1
        Call WriteChatOverHead(Userindex, "Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre ciudadana y real, y seras recompensado, lo prometo.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        If .Faccion.RecibioArmaduraCaos = 0 Then
            Call GiveFactionArmours(Userindex, True)
            Call GiveExpReward(Userindex, 0)
            .Faccion.RecibioArmaduraCaos = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = DatePart("yyyy", Date) & "-" & DatePart("m", Date) & "-" & DatePart("d", Date)
            .Faccion.RecibioExpInicialCaos = 1
            .Faccion.RecompensasCaos = 0
            .Faccion.NextRecompensa = 160
        End If
        If .flags.Navegando Then Call RefreshCharStatus(Userindex)
        Call LogEjercitoCaos(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)
    End With
End Sub

Public Sub RecompensaCaos(ByVal Userindex As Integer)
    Dim Ciudas    As Long
    Dim Lvl       As Byte
    Dim NextRecom As Long
    With UserList(Userindex)
        Lvl = .Stats.ELV
        Ciudas = .Faccion.CiudadanosMatados
        NextRecom = .Faccion.NextRecompensa
        If Ciudas < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Ciudas & " cuidadanos mas para recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        Select Case NextRecom
            Case 160:
                .Faccion.RecompensasCaos = 1
                .Faccion.NextRecompensa = 300
        
            Case 300:
                .Faccion.RecompensasCaos = 2
                .Faccion.NextRecompensa = 490
        
            Case 490:
                .Faccion.RecompensasCaos = 3
                .Faccion.NextRecompensa = 740
        
            Case 740:
                .Faccion.RecompensasCaos = 4
                .Faccion.NextRecompensa = 1100
        
            Case 1100:
                .Faccion.RecompensasCaos = 5
                .Faccion.NextRecompensa = 1500
        
            Case 1500:
                If Lvl < 27 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 6
                .Faccion.NextRecompensa = 2010
        
            Case 2010:
                .Faccion.RecompensasCaos = 7
                .Faccion.NextRecompensa = 2700
        
            Case 2700:
                .Faccion.RecompensasCaos = 8
                .Faccion.NextRecompensa = 4600
        
            Case 4600:
                If Lvl < 30 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 9
                .Faccion.NextRecompensa = 5800
        
            Case 5800:
                If Lvl < 31 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 31 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 10
                .Faccion.NextRecompensa = 6990
        
            Case 6990:
                If Lvl < 33 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 33 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 11
                .Faccion.NextRecompensa = 8100
        
            Case 8100:
                If Lvl < 35 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 35 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 12
                .Faccion.NextRecompensa = 9300
        
            Case 9300:
                If Lvl < 36 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 36 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 13
                .Faccion.NextRecompensa = 11500
        
            Case 11500:
                If Lvl < 37 Then
                    Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 14
                .Faccion.NextRecompensa = 23000
        
            Case 23000:
                Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos . Tu unica recompensa sera la sangre derramada. Continua asi!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
        
            Case Else:
                Exit Sub
        End Select
        Call WriteChatOverHead(Userindex, "Bien hecho " & TituloCaos(Userindex) & ", aqui tienes tu recompensa!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
        Call GiveFactionArmours(Userindex, True)
        Call GiveExpReward(Userindex, .Faccion.RecompensasCaos)
    End With
End Sub

Public Function TituloCaos(ByVal Userindex As Integer) As String
    Select Case UserList(Userindex).Faccion.RecompensasCaos
        Case 0
            TituloCaos = "Acolito"

        Case 1
            TituloCaos = "Alma Corrupta"

        Case 2
            TituloCaos = "Paria"

        Case 3
            TituloCaos = "Condenado"

        Case 4
            TituloCaos = "Esbirro"

        Case 5
            TituloCaos = "Sanguinario"

        Case 6
            TituloCaos = "Corruptor"

        Case 7
            TituloCaos = "Heraldo Impio"

        Case 8
            TituloCaos = "Caballero de la Oscuridad"

        Case 9
            TituloCaos = "Senor del Miedo"

        Case 10
            TituloCaos = "Ejecutor Infernal"

        Case 11
            TituloCaos = "Protector del Averno"

        Case 12
            TituloCaos = "Avatar de la Destruccion"

        Case 13
            TituloCaos = "Guardian del Mal"

        Case Else
            TituloCaos = "Campeon de la Oscuridad"
    End Select
End Function
