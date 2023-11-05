!include LogicLib.nsh
!include FileFunc.nsh
!include WordFunc.nsh
!include WinMessages.nsh
!include WinCore.nsh
!include WMI.nsh
!include StrRep.nsh

Var RAM_info 
Var RAM_size
Var GraphicsCard_info
Var Arch_info

!define /IfNDef LVM_GETITEMCOUNT 0x1004
!define /IfNDef LVM_GETITEMTEXTA 0x102D
!define /IfNDef LVM_GETITEMTEXTW 0x1073
!if "${NSIS_CHAR_SIZE}" > 1
	!define /IfNDef LVM_GETITEMTEXT ${LVM_GETITEMTEXTW}
!else
	!define /IfNDef LVM_GETITEMTEXT ${LVM_GETITEMTEXTA}
!endif

!ifndef SEE_MASK_NOCLOSEPROCESS
	!define SEE_MASK_NOCLOSEPROCESS 0x00000040
!endif

#Launch mail client and set recipent to #################
Function SupportMail_Link
	Pop $0
	
	#Just-in-time system information here, system information that requires WMI API calls are made in .onInit function
	#Registry accessible information
	ReadRegStr $3 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" ProductName ;OS
	ReadRegStr $R3 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" InstallationType ;client or server install
	ReadRegStr $R4 HKLM "HARDWARE\DESCRIPTION\System\BIOS" SystemFamily ;specific model
	ReadRegStr $R5 HKLM "HARDWARE\DESCRIPTION\System\BIOS" SystemManufacturer ;model maker
	ReadRegStr $R6 HKLM "HARDWARE\DESCRIPTION\System\CentralProcessor\0" ProcessorNameString ;CPU
	
	ExecShell "" "mailto:#############'&body=RAM: $RAM_info $RAM_size GB %0D%0AGraphics Card: $GraphicsCard_info %0D%0AOS :$3 %0D%0ACPU: $R6 %0D%0AInstall Type: $R3 %0D%0AModel: $R5, $R4 %0D%0AOS Type: $Arch_info %0D%0APlease add message below: %0D%0A"
FunctionEnd

#Support Functions for WMI information
#Getting Graphics Device
Function GraphicsCardInfo
	StrCpy $GraphicsCard_info $R2
FunctionEnd

#OSType 32/64 bit
Function ArchType
	StrCpy $Arch_info $R2
FunctionEnd

#Converting RAM from KB to GB
Function RAMSizeConvert
	System::Int64Op $R2 / 1048576
	Pop $R2
	StrCpy $RAM_size $R2
	StrCpy $R2 ""
FunctionEnd

#Getting RAM specifications
Function RAMInfoAggregate
	Pop $R2
	StrCpy $RAM_info "$RAM_info $R2"
FunctionEnd

Function .onInit
	#WIN32 API calls for system information
	#Windows Management Instrumentation (WMI) accessible information
	${WMIGetInfo} root\CIMV2 Win32_VideoController Description GraphicsCardInfo
	${WMIGetInfo} root\CIMV2 Win32_OperatingSystem OSArchitecture ArchType
	${WMIGetInfo} root\CIMV2 Win32_PhysicalMemory Manufacturer RAMInfoAggregate
	${WMIGetInfo} root\CIMV2 Win32_PhysicalMemory PartNumber RAMInfoAggregate
	${WMIGetInfo} root\CIMV2 Win32_PhysicalMemoryArray MaxCapacity RAMSizeConvert
FunctionEnd

Section "SysInfo get"
	Call SupportMail_Link
SectionEnd