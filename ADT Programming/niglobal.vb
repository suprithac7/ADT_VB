Option Strict Off
Option Explicit On
Module NIGLOBAL
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' 32-bit Visual Basic Language Interface
	' Version 1.9
	' Copyright 2001 National Instruments Corporation.
	' All Rights Reserved.
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'   This module contains the variable  declarations,
	'   constant definitions, and type information that
	'   is recognized by the entire application.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
    Public ibsta As Double
	Public iberr As Short
	Public ibcnt As Short
	Public ibcntl As Integer
	
	' Needed to register for GPIB Global Thread.
	Public Longibsta As Integer
	Public Longiberr As Integer
	Public Longibcnt As Integer
	Public GPIBglobalsRegistered As Short
	
	Public buf As String
	
	Public bytebuf() As Byte
	
	Public Const UNL As Integer = &H3F ' GPIB unlisten command
	Public Const UNT As Integer = &H5F ' GPIB untalk command
	Public Const GTL As Integer = &H1 ' GPIB go to local
	Public Const SDC As Integer = &H4 ' GPIB selected device clear
	Public Const PPC As Integer = &H5 ' GPIB parallel poll configure
	Public Const GGET As Integer = &H8 ' GPIB group execute trigger
	Public Const TCT As Integer = &H9 ' GPIB take control
	Public Const LLO As Integer = &H11 ' GPIB local lock out
	Public Const DCL As Integer = &H14 ' GPIB device clear
	Public Const PPU As Integer = &H15 ' GPIB parallel poll unconfigure
	Public Const SPE As Integer = &H18 ' GPIB serial poll enable
	Public Const SPD As Integer = &H19 ' GPIB serial poll disable
	Public Const PPE As Integer = &H60 ' GPIB parallel poll enable
	Public Const PPD As Integer = &H70 ' GPIB parallel poll disable
	
	' GPIB status bit vector :
	'       status variable ibsta and wait mask
	
	Public Const EERR As Integer = &H8000 ' Error detected
	Public Const TIMO As Integer = &H4000 ' Timeout
	Public Const EEND As Integer = &H2000 ' EOI or EOS detected
	Public Const SRQI As Integer = &H1000 ' SRQ detected by CIC
	Public Const RQS As Integer = &H800 ' Device requesting service
	Public Const CMPL As Integer = &H100 ' I/O completed
	Public Const LOK As Integer = &H80 ' Local lockout state
	Public Const RREM As Integer = &H40 ' Remote state
	Public Const CIC As Integer = &H20 ' Controller-in-Charge
	Public Const AATN As Integer = &H10 ' Attention asserted
	Public Const TACS As Integer = &H8 ' Talker active
	Public Const LACS As Integer = &H4 ' Listener active
	Public Const DTAS As Integer = &H2 ' Device trigger state
	Public Const DCAS As Integer = &H1 ' Device clear state
	
	' Error messages returned in global variable iberr
	
	Public Const EDVR As Short = 0 ' System error
	Public Const ECIC As Short = 1 ' Function requires GPIB board to be CIC
	Public Const ENOL As Short = 2 ' Write function detected no listeners
	Public Const EADR As Short = 3 ' Interface board not addressed correctly
	Public Const EARG As Short = 4 ' Invalid argument to function call
	Public Const ESAC As Short = 5 ' Function requires GPIB board to be SAC
	Public Const EABO As Short = 6 ' I/O operation aborted
	Public Const ENEB As Short = 7 ' Non-existent interface board
	Public Const EDMA As Short = 8 ' DMA Error
	Public Const EOIP As Short = 10 ' I/O operation started before previous
	' operation completed
	Public Const ECAP As Short = 11 ' No capability for intended operation
	Public Const EFSO As Short = 12 ' File system operation error
	Public Const EBUS As Short = 14 ' Command error during device call
	Public Const ESTB As Short = 15 ' Serial poll status byte lost
	Public Const ESRQ As Short = 16 ' SRQ remains asserted
	Public Const ETAB As Short = 20 ' The return buffer is full
	Public Const ELCK As Short = 21 ' Address or board is locked
	Public Const EARM As Short = 22 ' ibnotify's asynchronous event
	' notification failed to rearm
	Public Const EHDL As Short = 23 ' Invalid GPIB handle type for function
	Public Const WCFG As Short = 24 ' Configuration warning
	Public Const ECFG As Short = 24 ' Configuration warning
	Public Const EWIP As Short = 26 ' Wait already in progress on input ud
	Public Const ERST As Short = 27 ' The event notification was cancelled   
	' due to a reset of the interface
	Public Const EPWR As Short = 28 ' The interface lost power                           
	
	
	' EOS mode bits
	
	Public Const BIN As Integer = &H1000 ' Eight bit compare
	Public Const XEOS As Integer = &H800 ' Send EOI with EOS byte
	Public Const REOS As Integer = &H400 ' Terminate read on EOS
	
	' Timeout values and meanings
	
	Public Const TNONE As Short = 0 ' Infinite timeout (disabled)
	Public Const T10us As Short = 1 ' Timeout of 10 us (ideal)
	Public Const T30us As Short = 2 ' Timeout of 30 us (ideal)
	Public Const T100us As Short = 3 ' Timeout of 100 us (ideal)
	Public Const T300us As Short = 4 ' Timeout of 300 us (ideal)
	Public Const T1ms As Short = 5 ' Timeout of 1 ms (ideal)
	Public Const T3ms As Short = 6 ' Timeout of 3 ms (ideal)
	Public Const T10ms As Short = 7 ' Timeout of 10 ms (ideal)
	Public Const T30ms As Short = 8 ' Timeout of 30 ms (ideal)
	Public Const T100ms As Short = 9 ' Timeout of 100 ms (ideal)
	Public Const T300ms As Short = 10 ' Timeout of 300 ms (ideal)
	Public Const T1s As Short = 11 ' Timeout of 1 s (ideal)
	Public Const T3s As Short = 12 ' Timeout of 3 s (ideal)
	Public Const T10s As Short = 13 ' Timeout of 10 s (ideal)
	Public Const T30s As Short = 14 ' Timeout of 30 s (ideal)
	Public Const T100s As Short = 15 ' Timeout of 100 s (ideal)
	Public Const T300s As Short = 16 ' Timeout of 300 s (ideal)
	Public Const T1000s As Short = 17 ' Timeout of 1000 s (maximum)
	
	' IBLN constants
	
	Public Const ALL_SAD As Short = -1
	Public Const NO_SAD As Short = 0
	
	' The following constants are used for the second parameter of the
	' ibconfig function.  They are the "option" selection codes.
	
	Public Const IbcPAD As Integer = &H1 ' Primary Address
	Public Const IbcSAD As Integer = &H2 ' Secondary Address
	Public Const IbcTMO As Integer = &H3 ' Timeout Value
	Public Const IbcEOT As Integer = &H4 ' Send EOI with last data byte?
	Public Const IbcPPC As Integer = &H5 ' Parallel Poll Configure
	Public Const IbcREADDR As Integer = &H6 ' Repeat Addressing
	Public Const IbcAUTOPOLL As Integer = &H7 ' Disable Auto Serial Polling
	Public Const IbcCICPROT As Integer = &H8 ' Use the CIC Protocol?
	Public Const IbcIRQ As Integer = &H9 ' Use PIO for I/O
	Public Const IbcSC As Integer = &HA ' Board is System Controller.
	Public Const IbcSRE As Integer = &HB ' Assert SRE on device calls?
	Public Const IbcEOSrd As Integer = &HC ' Terminate reads on EOS.
	Public Const IbcEOSwrt As Integer = &HD ' Send EOI with EOS character.
	Public Const IbcEOScmp As Integer = &HE ' Use 7 or 8-bit EOS compare.
	Public Const IbcEOSchar As Integer = &HF ' The EOS character.
	Public Const IbcPP2 As Integer = &H10 ' Use Parallel Poll Mode 2.
	Public Const IbcTIMING As Integer = &H11 ' NORMAL, HIGH, or VERY_HIGH timing.
	Public Const IbcDMA As Integer = &H12 ' Use DMA for I/O.
	Public Const IbcReadAdjust As Integer = &H13 ' Swap bytes during an ibrd.
	Public Const IbcWriteAdjust As Integer = &H14 ' Swap bytes during an ibwrt.
	Public Const IbcSendLLO As Integer = &H17 ' Enable/disable the sending of LLO.
	Public Const IbcSPollTime As Integer = &H18 ' Set the timeout value for serial polls.
	Public Const IbcPPollTime As Integer = &H19 ' Set the parallel poll length period
	Public Const IbcEndBitIsNormal As Integer = &H1A ' Remove EOS from END bit of IBSTA.
	Public Const IbcUnAddr As Integer = &H1B ' Enable/disable device unaddressing.
	Public Const IbcSignalNumber As Integer = &H1C ' Set UNIX signal number - unsupported
	Public Const IbcBlockIfLocked As Integer = &H1D ' Enable/disable blocking for locked boards/devices
	Public Const IbcHSCableLength As Integer = &H1F ' Enable/disable high-speed handshaking.
	Public Const IbcIst As Integer = &H20 ' Set the IST bit
	Public Const IbcRsv As Integer = &H21 ' Set the RSV bit
	Public Const IbcLON As Integer = &H22 ' Enable listen only mode.
	
	
	'   Constants that can be used (in addition to the ibconfig constants)
	'   when calling the IBASK function.
	
	Public Const IbaPAD As Integer = &H1 ' Primary Address
	Public Const IbaSAD As Integer = &H2 ' Secondary Address
	Public Const IbaTMO As Integer = &H3 ' Timeout Value
	Public Const IbaEOT As Integer = &H4 ' Send EOI with last data byte?
	Public Const IbaPPC As Integer = &H5 ' Parallel Poll Configure
	Public Const IbaREADDR As Integer = &H6 ' Repeat Addressing
	Public Const IbaAUTOPOLL As Integer = &H7 ' Disable Auto Serial Polling
	Public Const IbaCICPROT As Integer = &H8 ' Use the CIC Protocol?
	Public Const IbaIRQ As Integer = &H9 ' Use PIO for I/O
	Public Const IbaSC As Integer = &HA ' Board is System Controller.
	Public Const IbaSRE As Integer = &HB ' Assert SRE on device calls?
	Public Const IbaEOSrd As Integer = &HC ' Terminate reads on EOS.
	Public Const IbaEOSwrt As Integer = &HD ' Send EOI with EOS character.
	Public Const IbaEOScmp As Integer = &HE ' Use 7 or 8-bit EOS compare.
	Public Const IbaEOSchar As Integer = &HF ' The EOS character.
	Public Const IbaPP2 As Integer = &H10 ' Use Parallel Poll Mode 2.
	Public Const IbaTIMING As Integer = &H11 ' NORMAL, HIGH, or VERY_HIGH timing.
	Public Const IbaDMA As Integer = &H12 ' Use DMA for I/O.
	Public Const IbaReadAdjust As Integer = &H13 ' Swap bytes during an ibrd.
	Public Const IbaWriteAdjust As Integer = &H14 ' Swap bytes during an ibwrt.
	Public Const IbaSendLLO As Integer = &H17 ' Enable/disable the sending of LLO.
	Public Const IbaSPollTime As Integer = &H18 ' Set the timeout value for serial polls.
	Public Const IbaPPollTime As Integer = &H19 ' Set the parallel poll length period
	Public Const IbaEndBitIsNormal As Integer = &H1A ' Remove EOS from END bit of IBSTA.
	Public Const IbaUnAddr As Integer = &H1B ' Enable/disable device unaddressing.
	Public Const IbaSignalNumber As Integer = &H1C ' Set UNIX signal number - unsupported
	Public Const IbaBlockIfLocked As Integer = &H1D ' Enable/disable blocking for locked boards/devices.
	Public Const IbaHSCableLength As Integer = &H1F ' Enable/disable high-speed handshaking.
	Public Const IbaIst As Integer = &H20 ' Set the IST bit
	Public Const IbaRsv As Integer = &H21 ' Set the RSV bit
	Public Const IbaLON As Integer = &H22 ' Enable listen only mode.
	Public Const IbaBNA As Integer = &H200 ' A device's access board.
	Public Const IbaBaseAddr As Integer = &H201 ' A GPIB board's base I/O address.
	Public Const IbaDmaChannel As Integer = &H202 ' A GPIB board's DMA channel.
	Public Const IbaIrqLevel As Integer = &H203 ' A GPIB board's IRQ level.
	Public Const IbaBaud As Integer = &H204 ' Baud rate used to communicate to CT box.
	Public Const IbaParity As Integer = &H205 ' Parity setting for CT box.
	Public Const IbaStopBits As Integer = &H206 ' Stop bits used for communicating to CT.
	Public Const IbaDataBits As Integer = &H207 ' Data bits used for communicating to CT.
	Public Const IbaComPort As Integer = &H208 ' System COM port used for CT box.
	Public Const IbaComIrqLevel As Integer = &H209 ' System COM port's interrupt level.
	Public Const IbaComPortBase As Integer = &H20A ' System COM port's base I/O address.
	Public Const IbaSingleCycleDma As Integer = &H20B ' Does the board use single cycle DMA?
	Public Const IbaSlotNumber As Integer = &H20C ' Board's slot number.
	Public Const IbaLPTNumber As Integer = &H20D ' Parallel port number
	Public Const IbaLPTType As Integer = &H20E ' Parallel port protocol
	
	' These are the values used by the 488.2 Send command
	
    Public Const NULLend As Integer = &H0 ' Do nothing at the end of a transfer
	Public Const NLend As Integer = &H1 ' Send NL with EOI after a transfer
	Public Const DABend As Integer = &H2 ' Send EOI with the last DAB
	
	' This value is useds by the 488.2 Receive command
	
	Public Const STOPend As Integer = &H100 ' Stop the read on EOI
	
	' The following values are used by the iblines function.  The integer
	' returned by iblines contains:
	'       The lower byte will contain a "monitor" bit mask.  If a bit
	'               is set (1) in this mask, then the corresponding line
	'               can be monitored by the driver.  If the bit is clear (0),
	'               then the line cannot be monitored.
	'       The upper byte will contain the status of the bus lines.
	'               Each bit corresponds to a certain bus line, and has
	'               a corresponding "monitor" bit in the lower byte.
	
	Public Const ValidEOI As Integer = &H80
	Public Const ValidATN As Integer = &H40
	Public Const ValidSRQ As Integer = &H20
	Public Const ValidREN As Integer = &H10
	Public Const ValidIFC As Integer = &H8
	Public Const ValidNRFD As Integer = &H4
	Public Const ValidNDAC As Integer = &H2
	Public Const ValidDAV As Integer = &H1
	Public Const BusEOI As Integer = &H8000
	Public Const BusATN As Integer = &H4000
	Public Const BusSRQ As Integer = &H2000
	Public Const BusREN As Integer = &H1000
	Public Const BusIFC As Integer = &H800
	Public Const BusNRFD As Integer = &H400
	Public Const BusNDAC As Integer = &H200
	Public Const BusDAV As Integer = &H100
	
	' This value is used to terminate an address list.  It should be
	' assigned to the last entry.  (488.2)
	
	Public Const NOADDR As Integer = &HFFFF
	
	' This value is defined for when the GPIBnotify Callback
	' has failed to rearm.
	
	Public Const IBNOTIFY_REARM_FAILED As Integer = &HE00A003F
	
	' These constants are for use with iblockx/ibunlockx
	' functions for GPIB-ENET
	
	Public Const TIMMEDIATE As Short = -1
	Public Const TINFINITE As Short = -2
	Public Const MAX_LOCKSHARENAME_LENGTH As Short = 64
End Module