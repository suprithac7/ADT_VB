Option Strict Off
Option Explicit On
Module VBIB32
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 32-bit Visual Basic Language Interface
    ' Version 1.81
    ' Copyright 2001 National Instruments Corporation.
    ' All Rights Reserved.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   This module contains the subroutine declarations,
    '   function declarations and constants required to use
    '   the National Instruments GPIB Dynamic Link Library
    '   (DLL) for controlling IEEE-488 instrumentation.  This
    '   file must be 'added' to your Visual Basic project
    '   (by choosing Add File from the File menu or pressing
    '   CTRL+F12) so that you can access the NI-488.2
    '   subroutines and functions.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   NI-488.2 DLL entry function declarations

    Declare Function ibask32 Lib "Gpib-32.dll" Alias "ibask" (ByVal ud As Integer, ByVal opt As Integer, ByRef value As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibbna32 Lib "Gpib-32.dll" Alias "ibbnaA" (ByVal ud As Integer, ByVal sstr As String) As Integer
    Declare Function ibcac32 Lib "Gpib-32.dll" Alias "ibcac" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibclr32 Lib "Gpib-32.dll" Alias "ibclr" (ByVal ud As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibcmd32 Lib "Gpib-32.dll" Alias "ibcmd" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibcmda32 Lib "Gpib-32.dll" Alias "ibcmda" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    Declare Function ibconfig32 Lib "Gpib-32.dll" Alias "ibconfig" (ByVal ud As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
    Declare Function ibdev32 Lib "Gpib-32.dll" Alias "ibdev" (ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
    Declare Function ibdma32 Lib "Gpib-32.dll" Alias "ibdma" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibeos32 Lib "Gpib-32.dll" Alias "ibeos" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibeot32 Lib "Gpib-32.dll" Alias "ibeot" (ByVal ud As Integer, ByVal v As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibfind32 Lib "Gpib-32.dll" Alias "ibfindA" (ByVal sstr As String) As Integer
    Declare Function ibgts32 Lib "Gpib-32.dll" Alias "ibgts" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibist32 Lib "Gpib-32.dll" Alias "ibist" (ByVal ud As Integer, ByVal v As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function iblck32 Lib "Gpib-32.dll" Alias "iblck" (ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Integer, ByVal arg1 As String) As Integer
    Declare Function iblines32 Lib "Gpib-32.dll" Alias "iblines" (ByVal ud As Integer, ByRef v As Integer) As Integer
    Declare Function ibln32 Lib "Gpib-32.dll" Alias "ibln" (ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ByRef ln As Integer) As Integer
    Declare Function ibloc32 Lib "Gpib-32.dll" Alias "ibloc" (ByVal ud As Integer) As Integer
    Declare Function iblock32 Lib "Gpib-32.dll" Alias "iblock" (ByVal ud As Integer) As Integer
    Declare Function ibonl32 Lib "Gpib-32.dll" Alias "ibonl" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibpad32 Lib "Gpib-32.dll" Alias "ibpad" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibpct32 Lib "Gpib-32.dll" Alias "ibpct" (ByVal ud As Integer) As Integer
    Declare Function ibppc32 Lib "Gpib-32.dll" Alias "ibppc" (ByVal ud As Integer, ByVal v As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibrd32 Lib "Gpib-32.dll" Alias "ibrd" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibrda32 Lib "Gpib-32.dll" Alias "ibrda" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibrdf32 Lib "Gpib-32.dll" Alias "ibrdfA" (ByVal ud As Integer, ByVal sstr As String) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibrpp32 Lib "Gpib-32.dll" Alias "ibrpp" (ByVal ud As Integer, ByVal sstr As String) As Integer
    Declare Function ibrsc32 Lib "Gpib-32.dll" Alias "ibrsc" (ByVal ud As Integer, ByVal v As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibrsp32 Lib "Gpib-32.dll" Alias "ibrsp" (ByVal ud As Integer, ByVal sstr As String) As Integer
    Declare Function ibrsv32 Lib "Gpib-32.dll" Alias "ibrsv" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibsad32 Lib "Gpib-32.dll" Alias "ibsad" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibsic32 Lib "Gpib-32.dll" Alias "ibsic" (ByVal ud As Integer) As Integer
    Declare Function ibsre32 Lib "Gpib-32.dll" Alias "ibsre" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibstop32 Lib "Gpib-32.dll" Alias "ibstop" (ByVal ud As Integer) As Integer
    Declare Function ibtmo32 Lib "Gpib-32.dll" Alias "ibtmo" (ByVal ud As Integer, ByVal v As Integer) As Integer
    Declare Function ibtrg32 Lib "Gpib-32.dll" Alias "ibtrg" (ByVal ud As Integer) As Integer
    Declare Function ibunlock32 Lib "Gpib-32.dll" Alias "ibunlock" (ByVal ud As Integer) As Integer
    Declare Function ibwait32 Lib "Gpib-32.dll" Alias "ibwait" (ByVal ud As Integer, ByVal mask As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibwrt32 Lib "Gpib-32.dll" Alias "ibwrt" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibwrta32 Lib "Gpib-32.dll" Alias "ibwrta" (ByVal ud As Integer, ByVal sstr As String, ByVal cnt As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function ibwrtf32 Lib "Gpib-32.dll" Alias "ibwrtfA" (ByVal ud As Integer, ByVal sstr As String) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub AllSpoll32 Lib "Gpib-32.dll" Alias "AllSpoll" (ByVal boardID As Integer, ByVal arg1 As String, ByVal arg2 As String)
    Declare Sub DevClear32 Lib "Gpib-32.dll" Alias "DevClear" (ByVal boardID As Integer, ByVal v As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub DevClearList32 Lib "Gpib-32.dll" Alias "DevClearList" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub EnableLocal32 Lib "Gpib-32.dll" Alias "EnableLocal" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub EnableRemote32 Lib "Gpib-32.dll" Alias "EnableRemote" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub FindLstn32 Lib "Gpib-32.dll" Alias "FindLstn" (ByVal boardID As Integer, ByVal arg1 As String, ByVal arg2 As String, ByVal limit As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub FindRQS32 Lib "Gpib-32.dll" Alias "FindRQS" (ByVal boardID As Integer, ByVal arg1 As String, ByRef result As Integer)
    Declare Sub PassControl32 Lib "Gpib-32.dll" Alias "PassControl" (ByVal boardID As Integer, ByVal addr As Integer)
    Declare Sub PPoll32 Lib "Gpib-32.dll" Alias "PPoll" (ByVal boardID As Integer, ByRef result As Integer)
    Declare Sub PPollConfig32 Lib "Gpib-32.dll" Alias "PPollConfig" (ByVal boardID As Integer, ByVal addr As Integer, ByVal line As Integer, ByVal sense As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub PPollUnconfig32 Lib "Gpib-32.dll" Alias "PPollUnconfig" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub RcvRespMsg32 Lib "Gpib-32.dll" Alias "RcvRespMsg" (ByVal boardID As Integer, ByVal arg1 As String, ByVal cnt As Integer, ByVal term As Integer)
    Declare Sub ReadStatusByte32 Lib "Gpib-32.dll" Alias "ReadStatusByte" (ByVal boardID As Integer, ByVal addr As Integer, ByRef result As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub Receive32 Lib "Gpib-32.dll" Alias "Receive" (ByVal boardID As Integer, ByVal addr As Integer, ByVal arg1 As String, ByVal cnt As Integer, ByVal term As Integer)
    Declare Sub ReceiveSetup32 Lib "Gpib-32.dll" Alias "ReceiveSetup" (ByVal boardID As Integer, ByVal addr As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub ResetSys32 Lib "Gpib-32.dll" Alias "ResetSys" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub Send32 Lib "Gpib-32.dll" Alias "Send" (ByVal boardID As Integer, ByVal addr As Integer, ByVal sstr As String, ByVal cnt As Integer, ByVal term As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub SendCmds32 Lib "Gpib-32.dll" Alias "SendCmds" (ByVal boardID As Integer, ByVal sstr As String, ByVal cnt As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub SendDataBytes32 Lib "Gpib-32.dll" Alias "SendDataBytes" (ByVal boardID As Integer, ByVal sstr As String, ByVal cnt As Integer, ByVal term As Integer)
    Declare Sub SendIFC32 Lib "Gpib-32.dll" Alias "SendIFC" (ByVal boardID As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub SendList32 Lib "Gpib-32.dll" Alias "SendList" (ByVal boardID As Integer, ByVal arg1 As String, ByVal arg2 As String, ByVal cnt As Integer, ByVal term As Integer)
    Declare Sub SendLLO32 Lib "Gpib-32.dll" Alias "SendLLO" (ByVal boardID As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub SendSetup32 Lib "Gpib-32.dll" Alias "SendSetup" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub SetRWLS32 Lib "Gpib-32.dll" Alias "SetRWLS" (ByVal boardID As Integer, ByVal arg1 As String)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub TestSys32 Lib "Gpib-32.dll" Alias "TestSys" (ByVal boardID As Integer, ByVal arg1 As String, ByVal arg2 As String)
    Declare Sub Trigger32 Lib "Gpib-32.dll" Alias "Trigger" (ByVal boardID As Integer, ByVal addr As Integer)
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Sub TriggerList32 Lib "Gpib-32.dll" Alias "TriggerList" (ByVal boardID As Integer, ByVal arg1 As String)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   DLL entry function declarations needed for GPIB global variables

    Declare Function RegisterGpibGlobalsForThread Lib "Gpib-32.dll" (ByRef Longibsta As Integer, ByRef Longiberr As Integer, ByRef Longibcnt As Integer, ByRef ibcntl As Integer) As Integer
    Declare Function UnregisterGpibGlobalsForThread Lib "Gpib-32.dll" () As Integer
    Declare Function ThreadIbsta32 Lib "Gpib-32.dll" Alias "ThreadIbsta" () As Integer
    Declare Function ThreadIbcnt32 Lib "Gpib-32.dll" Alias "ThreadIbcnt" () As Integer
    Declare Function ThreadIbcntl32 Lib "Gpib-32.dll" Alias "ThreadIbcntl" () As Integer
    Declare Function ThreadIberr32 Lib "Gpib-32.dll" Alias "ThreadIberr" () As Integer

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   DLL entry function declarations needed for GPIBnotify OLE control

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   DLL entry function declarations needed for GPIB-ENET functions

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Declare Function iblockx32 Lib "Gpib-32.dll" Alias "iblockxA" (ByVal ud As Integer, ByVal LockWaitTime As Integer, ByRef arg1 As String) As Integer
    Declare Function ibunlockx32 Lib "Gpib-32.dll" Alias "ibunlockx" (ByVal ud As Integer) As Integer


    Sub AllSpoll(ByVal boardID As Short, ByRef addrs() As Short, ByRef results() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call AllSpoll32(boardID, addrs(0), results(0))

        Call copy_ibvars()
    End Sub

    Sub copy_ibvars()
        ibsta = ConvertLongToInt(Longibsta)
        iberr = CShort(Longiberr)
        ibcnt = ConvertLongToInt(ibcntl)
    End Sub

    Sub DevClear(ByVal boardID As Short, ByVal addr As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call DevClear32(boardID, addr)

        Call copy_ibvars()
    End Sub

    Sub DevClearList(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call DevClearList32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub EnableLocal(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call EnableLocal32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub EnableRemote(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call EnableRemote32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub FindLstn(ByVal boardID As Short, ByRef addrs() As Short, ByRef results() As Short, ByVal limit As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call FindLstn32(boardID, addrs(0), results(0), limit)

        Call copy_ibvars()
    End Sub

    Sub FindRQS(ByVal boardID As Short, ByRef addrs() As Short, ByRef result As Short)
        Dim tmpresult As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call FindRQS32(boardID, addrs(0), tmpresult)

        result = ConvertLongToInt(tmpresult)

        Call copy_ibvars()
    End Sub

    Sub ibask(ByVal ud As Short, ByVal opt As Short, ByRef rval As Short)
        Dim tmprval As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibask32(ud, opt, tmprval)

        rval = ConvertLongToInt(tmprval)

        Call copy_ibvars()
    End Sub

    Sub ibbna(ByVal ud As Short, ByVal udname As String)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibbna32(ud, udname)

        Call copy_ibvars()
    End Sub

    Sub ibcac(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibcac32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibclr(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibclr32(ud)

        Call copy_ibvars()
    End Sub

    Sub ibcmd(ByVal ud As Short, ByVal buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibcmd32(ud, buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibcmda(ByVal ud As Short, ByVal buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibcmd32(ud, buf, cnt)

        ' When Visual Basic remapping buffer problem solved, then use:
        '    call ibcmda32(ud, ByVal buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibconfig(ByVal bdid As Short, ByVal opt As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibconfig32(bdid, opt, v)

        Call copy_ibvars()
    End Sub

    Sub ibdev(ByVal bdid As Short, ByVal pad As Short, ByVal sad As Short, ByVal tmo As Short, ByVal eot As Short, ByVal eos As Short, ByRef ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ud = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

        Call copy_ibvars()
    End Sub

    Sub ibdma(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibdma32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibeos(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibeos32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibeot(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibeot32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibfind(ByVal udname As String, ByRef ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ud = ConvertLongToInt(ibfind32(udname))

        Call copy_ibvars()
    End Sub

    Sub ibgts(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibgts32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibist(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibist32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub iblines(ByVal ud As Short, ByRef lines As Short)
        Dim tmplines As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call iblines32(ud, tmplines)

        lines = ConvertLongToInt(tmplines)

        Call copy_ibvars()
    End Sub

    Sub ibln(ByVal ud As Short, ByVal pad As Short, ByVal sad As Short, ByRef ln As Short)
        Dim tmpln As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibln32(ud, pad, sad, tmpln)

        ln = ConvertLongToInt(tmpln)

        Call copy_ibvars()
    End Sub

    Sub ibloc(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibloc32(ud)

        Call copy_ibvars()
    End Sub

    Sub iblck(ByVal ud As Short, ByVal v As Short, ByVal LockWaitTime As Integer)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call iblck32(ud, v, LockWaitTime, 0)

        Call copy_ibvars()
    End Sub

    Sub ibonl(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibonl32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibpad(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibpad32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibpct(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibpct32(ud)

        Call copy_ibvars()
    End Sub

    Sub ibppc(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibppc32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibrd(ByVal ud As Short, ByRef buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibrd32(ud, buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibrda(ByVal ud As Short, ByRef buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibrd32(ud, buf, cnt)

        ' When Visual Basic remapping buffer problem solved, use this:
        '    Call ibrda32(ud, ByVal buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibrdf(ByVal ud As Short, ByVal filename As String)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrdf32(ud, filename)

        Call copy_ibvars()
    End Sub

    Sub ibrdi(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrd32(ud, ibuf(0), cnt)

        Call copy_ibvars()
    End Sub

    Sub ibrdia(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrd32(ud, ibuf(0), cnt)

        ' When Visual Basic remapping buffer problem is solved, then use:
        '    Call ibrda32(u, ibuf(0), cnt)

        Call copy_ibvars()
    End Sub

    Sub ibrpp(ByVal ud As Short, ByRef ppr As Short)
        'Static tmp_str As New VB6.FixedLengthString(2)
        <VBFixedString(2)> Static tmp_str As String
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrpp32(ud, tmp_str)

        ppr = Asc(tmp_str)

        Call copy_ibvars()
    End Sub

    Sub ibrsc(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrsc32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibrsp(ByVal ud As Short, ByRef spr As Short)
        'Static tmp_str As New VB6.FixedLengthString(2)
        <VBFixedString(2)> Static tmp_str As String
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL
        Call ibrsp32(ud, tmp_str)

        spr = Asc(tmp_str)

        Call copy_ibvars()
    End Sub

    Sub ibrsv(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibrsv32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibsad(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibsad32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibsic(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibsic32(ud)

        Call copy_ibvars()
    End Sub

    Sub ibsre(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibsre32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibstop(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibstop32(ud)

        Call copy_ibvars()
    End Sub

    Sub ibtmo(ByVal ud As Short, ByVal v As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibtmo32(ud, v)

        Call copy_ibvars()
    End Sub

    Sub ibtrg(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call 32-bit DLL.
        Call ibtrg32(ud)

        Call copy_ibvars()
    End Sub

    Sub ibwait(ByVal ud As Short, ByVal mask As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibwait32(ud, mask)

        Call copy_ibvars()
    End Sub

    Sub ibwrt(ByVal ud As Short, ByVal buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibwrt32(ud, buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibwrta(ByVal ud As Short, ByVal buf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call ibwrt32(ud, buf, cnt)

        ' When Visual Basic remapping buffer problem is solved, use this:
        '    Call ibwrta32(ud, ByVal buf, cnt)

        Call copy_ibvars()
    End Sub

    Sub ibwrtf(ByVal ud As Short, ByVal filename As String)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibwrtf32(ud, filename)

        Call copy_ibvars()
    End Sub

    Sub ibwrti(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibwrt32(ud, ibuf(0), cnt)

        Call copy_ibvars()
    End Sub

    Sub ibwrtia(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibwrt32(ud, ibuf(0), cnt)

        ' When Visual Basic remapping buffer problem is solved, use this:
        '    Call ibwrta32(ud, ibuf(0), cnt)

        Call copy_ibvars()
    End Sub

    Function ilask(ByVal ud As Short, ByVal opt As Short, ByRef rval As Short) As Short
        Dim tmprval As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilask = ConvertLongToInt(ibask32(ud, opt, tmprval))

        rval = ConvertLongToInt(tmprval)

        Call copy_ibvars()
    End Function

    Function ilbna(ByVal ud As Short, ByVal udname As String) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilbna = ConvertLongToInt(ibbna32(ud, udname))

        Call copy_ibvars()
    End Function

    Function ilcac(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilcac = ConvertLongToInt(ibcac32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilclr(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilclr = ConvertLongToInt(ibclr32(ud))

        Call copy_ibvars()
    End Function

    Function ilcmd(ByVal ud As Short, ByVal buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilcmd = ConvertLongToInt(ibcmd32(ud, buf, cnt))

        Call copy_ibvars()
    End Function

    Function ilcmda(ByVal ud As Short, ByVal buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilcmda = ConvertLongToInt(ibcmd32(ud, buf, cnt))

        ' When Visual Basic remapping buffer problem is solved, use this:
        '    ilcmda = ConvertLongToInt(ibcmda32(ud, ByVal buf, cnt))

        Call copy_ibvars()
    End Function

    Function ilconfig(ByVal bdid As Short, ByVal opt As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, v))

        Call copy_ibvars()
    End Function

    Function ildev(ByVal bdid As Short, ByVal pad As Short, ByVal sad As Short, ByVal tmo As Short, ByVal eot As Short, ByVal eos As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

        Call copy_ibvars()
    End Function

    Function ildma(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ildma = ConvertLongToInt(ibdma32(ud, v))

        Call copy_ibvars()
    End Function

    Function ileos(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ileos = ConvertLongToInt(ibeos32(ud, v))

        Call copy_ibvars()
    End Function

    Function ileot(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ileot = ConvertLongToInt(ibeot32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilfind(ByVal udname As String) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilfind = ConvertLongToInt(ibfind32(udname))

        Call copy_ibvars()
    End Function

    Function ilgts(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilgts = ConvertLongToInt(ibgts32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilist(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilist = ConvertLongToInt(ibist32(ud, v))

        Call copy_ibvars()
    End Function

    Function illck(ByVal ud As Short, ByVal v As Short, ByVal LockWaitTime As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illck = ConvertLongToInt(iblck32(ud, v, LockWaitTime, 0))

        Call copy_ibvars()
    End Function

    Function illines(ByVal ud As Short, ByRef lines As Short) As Short
        Dim tmplines As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illines = ConvertLongToInt(iblines32(ud, tmplines))

        lines = ConvertLongToInt(tmplines)

        Call copy_ibvars()
    End Function

    Function illn(ByVal ud As Short, ByVal pad As Short, ByVal sad As Short, ByRef ln As Short) As Short
        Dim tmpln As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illn = ConvertLongToInt(ibln32(ud, pad, sad, tmpln))

        ln = ConvertLongToInt(tmpln)

        Call copy_ibvars()
    End Function

    Function illoc(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illoc = ConvertLongToInt(ibloc32(ud))

        Call copy_ibvars()
    End Function

    Function ilonl(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilonl = ConvertLongToInt(ibonl32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilpad(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilpad = ConvertLongToInt(ibpad32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilpct(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilpct = ConvertLongToInt(ibpct32(ud))

        Call copy_ibvars()
    End Function

    Function ilppc(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilppc = ConvertLongToInt(ibppc32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilrd(ByVal ud As Short, ByRef buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrd = ConvertLongToInt(ibrd32(ud, buf, cnt))

        Call copy_ibvars()
    End Function

    Function ilrda(ByVal ud As Short, ByRef buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrda = ConvertLongToInt(ibrd32(ud, buf, cnt))

        ' When Visual Basic remapping buffer problem solved, use this:
        '    ilrda = ConvertLongToInt(ibrda32(ud, ByVal buf, cnt))

        Call copy_ibvars()
    End Function

    Function ilrdf(ByVal ud As Short, ByVal filename As String) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrdf = ConvertLongToInt(ibrdf32(ud, filename))

        Call copy_ibvars()
    End Function

    Function ilrdi(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrdi = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

        Call copy_ibvars()
    End Function

    Function ilrdia(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrdia = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

        ' When Visual Basic remapping buffer problem solved, use this:
        '    ilrdia = ConvertLongToInt(ibrda32(ud, ibuf(0), cnt))

        Call copy_ibvars()
    End Function

    Function ilrpp(ByVal ud As Short, ByRef ppr As Short) As Short
        'Static tmp_str As New VB6.FixedLengthString(2)
        <VBFixedString(2)> Static tmp_str As String
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrpp = ConvertLongToInt(ibrpp32(ud, tmp_str))

        ppr = Asc(tmp_str)

        Call copy_ibvars()
    End Function

    Function ilrsc(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        ilrsc = ConvertLongToInt(ibrsc32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilrsp(ByVal ud As Short, ByRef spr As Short) As Short
        'Static tmp_str As New VB6.FixedLengthString(2)
        <VBFixedString(2)> Static tmp_str As String
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL
        ilrsp = ConvertLongToInt(ibrsp32(ud, tmp_str))

        spr = Asc(tmp_str)

        Call copy_ibvars()
    End Function

    Function ilrsv(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilrsv = ConvertLongToInt(ibrsv32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilsad(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        ilsad = ConvertLongToInt(ibsad32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilsic(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        ilsic = ConvertLongToInt(ibsic32(ud))

        Call copy_ibvars()
    End Function

    Function ilsre(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        ilsre = ConvertLongToInt(ibsre32(ud, v))

        Call copy_ibvars()
    End Function

    Function ilstop(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        ilstop = ConvertLongToInt(ibstop32(ud))

        Call copy_ibvars()
    End Function

    Function iltmo(ByVal ud As Short, ByVal v As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        '  Call the 32-bit DLL.
        iltmo = ConvertLongToInt(ibtmo32(ud, v))

        Call copy_ibvars()
    End Function

    Function iltrg(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call 32-bit DLL.
        iltrg = ConvertLongToInt(ibtrg32(ud))

        Call copy_ibvars()
    End Function

    Function ilwait(ByVal ud As Short, ByVal mask As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwait = ConvertLongToInt(ibwait32(ud, mask))

        Call copy_ibvars()
    End Function

    Function ilwrt(ByVal ud As Short, ByVal buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwrt = ConvertLongToInt(ibwrt32(ud, buf, cnt))

        Call copy_ibvars()
    End Function

    Function ilwrta(ByVal ud As Short, ByVal buf As String, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwrta = ConvertLongToInt(ibwrt32(ud, buf, cnt))

        ' When the Visual Basic remapping solved, use this:
        '    ilwrta = ConvertLongToInt(ibwrta32(ud, ByVal buf, cnt))

        Call copy_ibvars()

    End Function

    Function ilwrtf(ByVal ud As Short, ByVal filename As String) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwrtf = ConvertLongToInt(ibwrtf32(ud, filename))

        Call copy_ibvars()
    End Function

    Function ilwrti(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwrti = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

        Call copy_ibvars()
    End Function

    Function ilwrtia(ByVal ud As Short, ByRef ibuf() As Short, ByVal cnt As Integer) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilwrtia = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

        ' When Visual Basic remapping buffer problem solved, use this:
        '    ilwrtia = ConvertLongToInt(ibwrta32(ud, ibuf(0), cnt))

        Call copy_ibvars()
    End Function

    Sub PassControl(ByVal boardID As Short, ByVal addr As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call PassControl32(boardID, addr)

        Call copy_ibvars()
    End Sub

    Sub Ppoll(ByVal boardID As Short, ByRef result As Short)
        Dim tmpresult As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call PPoll32(boardID, tmpresult)

        result = ConvertLongToInt(tmpresult)

        Call copy_ibvars()
    End Sub

    Sub PpollConfig(ByVal boardID As Short, ByVal addr As Short, ByVal lline As Short, ByVal sense As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call PPollConfig32(boardID, addr, lline, sense)

        Call copy_ibvars()
    End Sub

    Sub PpollUnconfig(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call PPollUnconfig32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub RcvRespMsg(ByVal boardID As Short, ByRef buf As String, ByVal term As Short)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call RcvRespMsg32(boardID, buf, cnt, term)

        Call copy_ibvars()
    End Sub

    Sub ReadStatusByte(ByVal boardID As Short, ByVal addr As Short, ByRef result As Short)
        Dim tmpresult As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ReadStatusByte32(boardID, addr, tmpresult)

        result = ConvertLongToInt(tmpresult)

        Call copy_ibvars()
    End Sub

    Sub Receive(ByVal boardID As Short, ByVal addr As Short, ByRef buf As String, ByVal term As Short)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call Receive32(boardID, addr, buf, cnt, term)

        Call copy_ibvars()
    End Sub

    Sub ReceiveSetup(ByVal boardID As Short, ByVal addr As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ReceiveSetup32(boardID, addr)

        Call copy_ibvars()
    End Sub

    Sub ResetSys(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ResetSys32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub Send(ByVal boardID As Short, ByVal addr As Short, ByVal buf As String, ByVal term As Short)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call Send32(boardID, addr, buf, cnt, term)

        Call copy_ibvars()
    End Sub

    Sub SendCmds(ByVal boardID As Short, ByVal cmdbuf As String)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(cmdbuf))

        ' Call the 32-bit DLL.
        Call SendCmds32(boardID, cmdbuf, cnt)

        Call copy_ibvars()
    End Sub

    Sub SendDataBytes(ByVal boardID As Short, ByVal buf As String, ByVal term As Short)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call SendDataBytes32(boardID, buf, cnt, term)

        Call copy_ibvars()
    End Sub

    Sub SendIFC(ByVal boardID As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call SendIFC32(boardID)

        Call copy_ibvars()
    End Sub

    Sub SendList(ByVal boardID As Short, ByRef addr() As Short, ByVal buf As String, ByVal term As Short)
        Dim cnt As Integer

        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        cnt = CInt(Len(buf))

        ' Call the 32-bit DLL.
        Call SendList32(boardID, addr(0), buf, cnt, term)

        Call copy_ibvars()
    End Sub

    Sub SendLLO(ByVal boardID As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call SendLLO32(boardID)

        Call copy_ibvars()
    End Sub

    Sub SendSetup(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call SendSetup32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub SetRWLS(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call SetRWLS32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub TestSRQ(ByVal boardID As Short, ByRef result As Short)
        Call ibwait(boardID, 0)

        If ibsta And &H1000 Then
            result = 1
        Else
            result = 0
        End If

    End Sub

    Sub TestSys(ByVal boardID As Short, ByRef addrs() As Short, ByRef results() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call TestSys32(boardID, addrs(0), results(0))

        Call copy_ibvars()
    End Sub

    Sub Trigger(ByVal boardID As Short, ByVal addr As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call Trigger32(boardID, addr)

        Call copy_ibvars()
    End Sub

    Sub TriggerList(ByVal boardID As Short, ByRef addrs() As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call TriggerList32(boardID, addrs(0))

        Call copy_ibvars()
    End Sub

    Sub WaitSRQ(ByVal boardID As Short, ByRef result As Short)
        Call ibwait(boardID, &H5000)

        If ibsta And &H1000 Then
            result = 1
        Else
            result = 0
        End If
    End Sub


    Public Function ConvertLongToInt(ByRef LongNumb As Long) As Integer

        If (LongNumb And &H8000) = 0 Then
            ConvertLongToInt = LongNumb And &HFFFF
        Else
            ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF)
        End If

    End Function

    Public Sub RegisterGPIBGlobals()
        Dim rc As Integer

        rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
        If (rc = 0) Then
            GPIBglobalsRegistered = 1
        ElseIf (rc = 1) Then
            rc = UnregisterGpibGlobalsForThread
            rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
            GPIBglobalsRegistered = 1
        ElseIf (rc = 2) Then
            rc = UnregisterGpibGlobalsForThread
            ibsta = &H8000
            iberr = EDVR
            ibcntl = &HDEAD37F0
        ElseIf (rc = 3) Then
            rc = UnregisterGpibGlobalsForThread
            ibsta = &H8000
            iberr = EDVR
            ibcntl = &HDEAD37F0
        Else
            ibsta = &H8000
            iberr = EDVR
            ibcntl = &HDEAD37F0
        End If
    End Sub

    Public Sub UnregisterGPIBGlobals()
        Dim rc As Integer

        rc = UnregisterGpibGlobalsForThread
        GPIBglobalsRegistered = 0

    End Sub



    Public Function ThreadIbsta() As Short
        ' Call the 32-bit DLL.
        ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
    End Function

    Public Function ThreadIberr() As Short
        ' Call the 32-bit DLL.
        ThreadIberr = ConvertLongToInt(ThreadIberr32())
    End Function

    Public Function ThreadIbcnt() As Short
        ' Call the 32-bit DLL.
        ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
    End Function

    Public Function ThreadIbcntl() As Integer
        ' Call the 32-bit DLL.
        ThreadIbcntl = ThreadIbcntl32()
    End Function

    Public Function illock(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illock = ConvertLongToInt(iblock32(ud))

        Call copy_ibvars()
    End Function

    Public Function ilunlock(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilunlock = ConvertLongToInt(ibunlock32(ud))

        Call copy_ibvars()
    End Function

    Public Sub iblock(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call iblock32(ud)

        Call copy_ibvars()
    End Sub

    Public Sub ibunlock(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibunlock32(ud)

        Call copy_ibvars()
    End Sub

    Public Function illockx(ByVal ud As Short, ByVal LockWaitTime As Short, ByVal buf As String) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        illockx = ConvertLongToInt(iblockx32(ud, LockWaitTime, buf))

        Call copy_ibvars()
    End Function

    Public Function ilunlockx(ByVal ud As Short) As Short
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        ilunlockx = ConvertLongToInt(ibunlockx32(ud))

        Call copy_ibvars()
    End Function

    Public Sub iblockx(ByVal ud As Short, ByVal LockWaitTime As Short, ByVal buf As String)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call iblockx32(ud, LockWaitTime, buf)

        Call copy_ibvars()
    End Sub

    Public Sub ibunlockx(ByVal ud As Short)
        ' Check to see if GPIB Global variables are registered
        If (GPIBglobalsRegistered = 0) Then
            Call RegisterGPIBGlobals()
        End If

        ' Call the 32-bit DLL.
        Call ibunlockx32(ud)

        Call copy_ibvars()
    End Sub
End Module