Attribute VB_Name = "mdlCOmercioConUsuario"
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 10000
Public Const MAX_OFFER_SLOTS    As Long = 40
Public Const GOLD_OFFER_SLOT    As Integer = MAX_OFFER_SLOTS + 1

Public Type tCOmercioUsuario
    DestUsu As Integer
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer
    GoldAmount As Long
    cant(1 To MAX_OFFER_SLOTS) As Long
    Acepto As Boolean
    Confirmo As Boolean
End Type

Private Type tOfferItem
    ObjIndex As Integer
    Amount As Long
End Type

Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
    On Error GoTo ErrorHandler
    If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
        If UserList(Origen).flags.Comerciando Or UserList(Destino).flags.Comerciando Then
            Call WriteConsoleMsg(Origen, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(Destino, "No puedes comerciar en este momento", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
        Call UpdateUserInv(True, Origen, 0)
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = True
        Call UpdateUserInv(True, Destino, 0)
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = True
    Else
        Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

Public Sub EnviarOferta(ByVal Userindex As Integer, ByVal OfferSlot As Byte)
    Dim ObjIndex       As Integer
    Dim ObjAmount      As Long
    Dim OtherUserIndex As Integer
    OtherUserIndex = UserList(Userindex).ComUsu.DestUsu
    With UserList(OtherUserIndex)
        If OfferSlot = GOLD_OFFER_SLOT Then
            ObjIndex = iORO
            ObjAmount = .ComUsu.GoldAmount
        Else
            ObjIndex = .ComUsu.Objeto(OfferSlot)
            ObjAmount = .ComUsu.cant(OfferSlot)
        End If
    End With
    Call WriteChangeUserTradeSlot(Userindex, OfferSlot, ObjIndex, ObjAmount)
End Sub

Public Sub FinComerciarUsu(ByVal Userindex As Integer)
    Dim i As Long
    With UserList(Userindex)
        If .ComUsu.DestUsu > 0 Then
            Call WriteUserCommerceEnd(Userindex)
        End If
        .ComUsu.Acepto = False
        .ComUsu.Confirmo = False
        .ComUsu.DestUsu = 0
        For i = 1 To MAX_OFFER_SLOTS
            .ComUsu.cant(i) = 0
            .ComUsu.Objeto(i) = 0
        Next i
        .ComUsu.GoldAmount = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
End Sub

Public Sub AceptarComercioUsu(ByVal Userindex As Integer)
    Dim TradingObj    As obj
    Dim OtroUserIndex As Integer
    Dim OfferSlot     As Integer
    UserList(Userindex).ComUsu.Acepto = True
    OtroUserIndex = UserList(Userindex).ComUsu.DestUsu
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Exit Sub
    End If
    If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
        Call FinComerciarUsu(Userindex)
        Exit Sub
    End If
    If Not HasOfferedItems(Userindex) Then
        Call WriteConsoleMsg(Userindex, "El comercio se cancelo porque no posees los items que ofertaste!!!", FontTypeNames.FONTTYPE_WARNING)
        Call WriteConsoleMsg(OtroUserIndex, "El comercio se cancelo porque " & UserList(Userindex).Name & " no posee los items que oferto!!!", FontTypeNames.FONTTYPE_WARNING)
        Call FinComerciarUsu(Userindex)
        Call FinComerciarUsu(OtroUserIndex)
        Exit Sub
    ElseIf Not HasOfferedItems(OtroUserIndex) Then
        Call WriteConsoleMsg(Userindex, "El comercio se cancelo porque " & UserList(OtroUserIndex).Name & " no posee los items que oferto!!!", FontTypeNames.FONTTYPE_WARNING)
        Call WriteConsoleMsg(OtroUserIndex, "El comercio se cancelo porque no posees los items que ofertaste!!!", FontTypeNames.FONTTYPE_WARNING)
        Call FinComerciarUsu(Userindex)
        Call FinComerciarUsu(OtroUserIndex)
        Exit Sub
    End If
    For OfferSlot = 1 To MAX_OFFER_SLOTS + 1
        With UserList(Userindex)
            If OfferSlot = GOLD_OFFER_SLOT Then
                .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.Name & " solto oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                Call WriteUpdateUserStats(Userindex)
                UserList(OtroUserIndex).Stats.Gld = UserList(OtroUserIndex).Stats.Gld + .ComUsu.GoldAmount
                Call WriteUpdateUserStats(OtroUserIndex)
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
                End If
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, Userindex)
                If ObjData(TradingObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " le paso en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                End If
                If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(UserList(OtroUserIndex).Name & " le paso en comercio seguro a " & .Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
        With UserList(OtroUserIndex)
            If OfferSlot = GOLD_OFFER_SLOT Then
                .Stats.Gld = .Stats.Gld - .ComUsu.GoldAmount
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(.Name & " solto oro en comercio seguro con " & UserList(Userindex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                Call WriteUpdateUserStats(OtroUserIndex)
                UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + .ComUsu.GoldAmount
                If .ComUsu.GoldAmount > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(Userindex).Name & " recibio oro en comercio seguro con " & .Name & ". Cantidad: " & .ComUsu.GoldAmount)
                Call WriteUpdateUserStats(Userindex)
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                If Not MeterItemEnInventario(Userindex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, TradingObj)
                End If
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, OtroUserIndex)
                If ObjData(TradingObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " le paso en comercio seguro a " & UserList(Userindex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                End If
                If TradingObj.Amount > MAX_OBJ_LOGUEABLE Then
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le paso en comercio seguro a " & UserList(Userindex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
    Next OfferSlot
    Call FinComerciarUsu(Userindex)
    Call FinComerciarUsu(OtroUserIndex)
End Sub

Public Sub AgregarOferta(ByVal Userindex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
    If PuedeSeguirComerciando(Userindex) Then
        With UserList(Userindex).ComUsu
            If Not .Confirmo Then
                If IsGold Then
                    .GoldAmount = .GoldAmount + Amount
                    If .GoldAmount < 0 Then .GoldAmount = 0
                Else
                    If ObjIndex > 0 Then .Objeto(OfferSlot) = ObjIndex
                    .cant(OfferSlot) = .cant(OfferSlot) + Amount
                    If .cant(OfferSlot) <= 0 Then
                        .Objeto(OfferSlot) = 0
                        .cant(OfferSlot) = 0
                    End If
                End If
            End If
        End With
    End If
End Sub

Public Function PuedeSeguirComerciando(ByVal Userindex As Integer) As Boolean
    Dim OtroUserIndex    As Integer
    Dim ComercioInvalido As Boolean
    With UserList(Userindex)
        If .ComUsu.DestUsu <= 0 Or .ComUsu.DestUsu > MaxUsers Then
            ComercioInvalido = True
        End If
        OtroUserIndex = .ComUsu.DestUsu
        If Not ComercioInvalido Then
            If UserList(OtroUserIndex).flags.UserLogged = False Or .flags.UserLogged = False Then
                ComercioInvalido = True
            End If
        End If
        If Not ComercioInvalido Then
            If UserList(OtroUserIndex).ComUsu.DestUsu <> Userindex Then
                ComercioInvalido = True
            End If
        End If
        If Not ComercioInvalido Then
            If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        If Not ComercioInvalido Then
            If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        If Not ComercioInvalido Then
            If UserList(OtroUserIndex).flags.Muerto = 1 Then
                ComercioInvalido = True
            End If
        End If
        If ComercioInvalido = True Then
            Call FinComerciarUsu(Userindex)
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call FinComerciarUsu(OtroUserIndex)
            End If
            Exit Function
        End If
    End With
    PuedeSeguirComerciando = True
End Function

Private Function HasOfferedItems(ByVal Userindex As Integer) As Boolean
    Dim OfferedItems(MAX_OFFER_SLOTS - 1) As tOfferItem
    Dim Slot                              As Long
    Dim SlotAux                           As Long
    Dim SlotCount                         As Long
    Dim ObjIndex                          As Integer
    With UserList(Userindex).ComUsu
        For Slot = 1 To MAX_OFFER_SLOTS
            ObjIndex = .Objeto(Slot)
            If ObjIndex > 0 Then
                For SlotAux = 0 To SlotCount - 1
                    If ObjIndex = OfferedItems(SlotAux).ObjIndex Then
                        OfferedItems(SlotAux).Amount = OfferedItems(SlotAux).Amount + .cant(Slot)
                        Exit For
                    End If
                Next SlotAux
                If SlotAux = SlotCount Then
                    OfferedItems(SlotCount).ObjIndex = ObjIndex
                    OfferedItems(SlotCount).Amount = .cant(Slot)
                    SlotCount = SlotCount + 1
                End If
            End If
        Next Slot
        For Slot = 0 To SlotCount - 1
            If Not HasEnoughItems(Userindex, OfferedItems(Slot).ObjIndex, OfferedItems(Slot).Amount) Then Exit Function
        Next Slot
        If UserList(Userindex).Stats.Gld < .GoldAmount Then Exit Function
    End With
    HasOfferedItems = True
End Function
