<div align="center">

## Looking inside the executable \- a VB guide to the portable executable format


</div>

### Description

Shows how the executable is laid out so that you can browse it's contents...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Advanced
**User Rating**    |4.9 (79 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-looking-inside-the-executable-a-vb-guide-to-the-portable-executable-format__1-45316/archive/master.zip)





### Source Code

<font size=2>
<h2>Inside the executable: The Portable Executable Format</h2>
<p>The source code can be found at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=41711&lngWId=1</p>
<p>The <b>Portable Executable Format</b> is the data structure that describes how the various parts of a Win32
executable file are held together. It allows the operating system to load the executable and to locate the dynamically
linked libraries required to run that executable and to navigate the code,data and resource sections compiled into that
executable.</p>
<h3>Getting over DOS</h3>
<p>The PE Format was created for Windows but Microsoft had to make sure that running such an executable in DOS would
yield a meaningful error message and exit. To this end the very first bit of a windows executable file is actually a DOS
executable (sometimes known as the <b>stub</b>) which writes "This program requires Windows" or similar then exits.</p>
<p>The format of the DOS stub is:</p>
<p>
<code><pre>
Private Type IMAGE_DOS_HEADER
 e_magic As Integer ''\\ Magic number
 e_cblp As Integer ''\\ Bytes on last page of file
 e_cp As Integer  ''\\ Pages in file
 e_crlc As Integer ''\\ Relocations
 e_cparhdr As Integer ''\\ Size of header in paragraphs
 e_minalloc As Integer ''\\ Minimum extra paragraphs needed
 e_maxalloc As Integer ''\\ Maximum extra paragraphs needed
 e_ss As Integer ''\\ Initial (relative) SS value
 e_sp As Integer ''\\ Initial SP value
 e_csum As Integer ''\\ Checksum
 e_ip As Integer ''\\ Initial IP value
 e_cs As Integer ''\\ Initial (relative) CS value
 e_lfarlc As Integer ''\\ File address of relocation table
 e_ovno As Integer ''\\ Overlay number
 e_res(0 To 3) As Integer ''\\ Reserved words
 e_oemid As Integer ''\\ OEM identifier (for e_oeminfo)
 e_oeminfo As Integer ''\\ OEM information; e_oemid specific
 e_res2(0 To 9) As Integer ''\\ Reserved words
 e_lfanew As Long ''\\ File address of new exe header
End Type
</pre></code>
</p>
<p>The only field of this structure that is of interest to Windows is <b>e_lfanew</b> which is the file pointer to the new
Windows executable header. To skip over the DOS part of the program, set the file pointer to the value held in this field:</p>
<p>
<code><pre>
Private Sub SkipDOSStub(ByVal hfile As Long)
Dim BytesRead As Long
'\\ Go to start of file...
Call SetFilePointer(hfile, 0, 0, FILE_BEGIN)
If Err.LastDllError Then
 Debug.Print LastSystemError
End If
Dim stub As IMAGE_DOS_HEADER
Call ReadFileLong(hfile, VarPtr(stub), Len(stub), BytesRead, ByVal 0&)
Call SetFilePointer(hfile, stub.e_lfanew, 0, FILE_BEGIN)
End Sub
</pre></code>
<p>
<h3>The NT header</h3>
<p>The NT header holds the information needed by the windows program loader to load the program. It consists of the PE File signature
followed by an <b>IMAGE_FILE_HEADER</b> and <b>IMAGE_OPTIONAL_HEADER</b> records.</p>
<p>For applications designed to run under Windows (i.e. not OS/2 or VxD files) the four bytes of the <b>PE File signature</b> should equal &h4550.
The other defined signatures are:</p>
<p>
<code><pre>
Public Enum ImageSignatureTypes
 IMAGE_DOS_SIGNATURE = &H5A4D  ''\\ MZ
 IMAGE_OS2_SIGNATURE = &H454E  ''\\ NE
 IMAGE_OS2_SIGNATURE_LE = &H454C ''\\ LE
 IMAGE_VXD_SIGNATURE = &H454C  ''\\ LE
 IMAGE_NT_SIGNATURE = &H4550  ''\\ PE00
End Enum
</pre></code>
</p>
<p>Following the PE file signature is the <b>IMAGE_NT_HEADERS</b> structure that stores information about the target environment of the executable.
The structure is:</p>
<p>
<code><pre>
Private Type IMAGE_FILE_HEADER
 Machine As Integer
 NumberOfSections As Integer
 TimeDateStamp As Long
 PointerToSymbolTable As Long
 NumberOfSymbols As Long
 SizeOfOptionalHeader As Integer
 Characteristics As Integer
End Type
</pre></code>
</p>
<p>The <b>Machine</b> member describes what target CPU the executable was compiled for. It can be one of:</p>
<p>
<code><pre>
Public Enum ImageMachineTypes
 IMAGE_FILE_MACHINE_I386 = &H14C ''\\ Intel 386.
 IMAGE_FILE_MACHINE_R3000 = &H162 ''\\ MIPS little-endian,= &H160 big-endian
 IMAGE_FILE_MACHINE_R4000 = &H166 ''\\ MIPS little-endian
 IMAGE_FILE_MACHINE_R10000 = &H168 ''\\ MIPS little-endian
 IMAGE_FILE_MACHINE_WCEMIPSV2 = &H169 ''\\ MIPS little-endian WCE v2
 IMAGE_FILE_MACHINE_ALPHA = &H184  ''\\ Alpha_AXP
 IMAGE_FILE_MACHINE_POWERPC = &H1F0 ''\\ IBM PowerPC Little-Endian
 IMAGE_FILE_MACHINE_SH3 = &H1A2 ''\\ SH3 little-endian
 IMAGE_FILE_MACHINE_SH3E = &H1A4 ''\\ SH3E little-endian
 IMAGE_FILE_MACHINE_SH4 = &H1A6 ''\\ SH4 little-endian
 IMAGE_FILE_MACHINE_ARM = &H1C0 ''\\ ARM Little-Endian
 IMAGE_FILE_MACHINE_IA64 = &H200 ''\\ Intel 64
End Enum
</pre></code>
</p>
<p>The <b>SizeOfOptionalHeader</b> member indicates the size (in bytes) of the <b>IMAGE_OPTIONAL_HEADER</b> structure that immediatley follows it.
 In practice this structure is not optional so that is a bit of a misnomer. This structure is defined as:</p>
<p>
<code><pre>
Private Type IMAGE_OPTIONAL_HEADER
 Magic As Integer
 MajorLinkerVersion As Byte
 MinorLinkerVersion As Byte
 SizeOfCode As Long
 SizeOfInitializedData As Long
 SizeOfUninitializedData As Long
 AddressOfEntryPoint As Long
 BaseOfCode As Long
 BaseOfData As Long
End Type
</pre></code>
</p>
<p> and this in turn is immediately followed by the <b>IMAGE_OPTIONAL_HEADER_NT</b> structure:</p>
<p>
<code><pre>
Private Type IMAGE_OPTIONAL_HEADER_NT
 ImageBase As Long
 SectionAlignment As Long
 FileAlignment As Long
 MajorOperatingSystemVersion As Integer
 MinorOperatingSystemVersion As Integer
 MajorImageVersion As Integer
 MinorImageVersion As Integer
 MajorSubsystemVersion As Integer
 MinorSubsystemVersion As Integer
 Win32VersionValue As Long
 SizeOfImage As Long
 SizeOfHeaders As Long
 CheckSum As Long
 Subsystem As Integer
 DllCharacteristics As Integer
 SizeOfStackReserve As Long
 SizeOfStackCommit As Long
 SizeOfHeapReserve As Long
 SizeOfHeapCommit As Long
 LoaderFlags As Long
 NumberOfRvaAndSizes As Long
 DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type
</pre></code>
</p>
<p>The most useful field of this structure (to my purposes, anyhow) are the 16 <b>IMAGE_DATA_DIRECTORY</b> entries. These describe whereabouts
(if at all) the particular sections of the executable are located. The structure is defined thus:</p>
<p>
<code><pre>
Private Type IMAGE_DATA_DIRECTORY
 VirtualAddress As Long
 Size As Long
End Type
</pre></code>
</p>
<p>And the directories are held in order thus:</p>
<p>
<code><pre>
Public Enum ImageDataDirectoryIndexes
 IMAGE_DIRECTORY_ENTRY_EXPORT = 0 ''\\ Export Directory
 IMAGE_DIRECTORY_ENTRY_IMPORT = 1 ''\\ Import Directory
 IMAGE_DIRECTORY_ENTRY_RESOURCE = 2 ''\\ Resource Directory
 IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3 ''\\ Exception Directory
 IMAGE_DIRECTORY_ENTRY_SECURITY = 4 ''\\ Security Directory
 IMAGE_DIRECTORY_ENTRY_BASERELOC = 5 ''\\ Base Relocation Table
 IMAGE_DIRECTORY_ENTRY_DEBUG = 6 ''\\ Debug Directory
 IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7 ''\\ Architecture Specific Data
 IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8 ''\\ RVA of GP
 IMAGE_DIRECTORY_ENTRY_TLS = 9 ''\\ TLS Directory
 IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10 ''\\ Load Configuration Directory
 IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11 ''\\ Bound Import Directory in headers
 IMAGE_DIRECTORY_ENTRY_IAT = 12 ''\\ Import Address Table
 IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT = 13 ''\\ Delay Load Import Descriptors
End Enum
</pre></code>
</p>
<p>Note that is an executable does not contain one of the sections (as is often the case) there will be an IMAGE_DATA_DIRECTORY for it but the
address and size will both be zero.</p>
<h2>The image data directories</h2>
<h3>The exports directory</h3>
<p>The exports directory holds details of the functions exported by this executable. For example, if you were to look in the exports directory
of the MSVBVM50.dll it would list all the functions it exports that make up the visual basic 5 runtime environment.</p>
<p>This directory consists of some info to tell you how many exported functions there are followed by three parallel arrays which give you the
address, name and ordinal of the functions respectively. The structure is defined thus:
</p>
<p>
<code><pre>
Private Type IMAGE_EXPORT_DIRECTORY
 Characteristics As Long
 TimeDateStamp As Long
 MajorVersion As Integer
 MinorVersion As Integer
 lpName As Long
 Base As Long
 NumberOfFunctions As Long
 NumberOfNames As Long
 lpAddressOfFunctions As Long '\\ Three parrallel arrays...(LONG)
 lpAddressOfNames As Long  '\\ (LONG)
 lpAddressOfNameOrdinals As Long '\\ (INTEGER)
End Type
</pre></code>
</p>
<p>And you can read this info from the executable thus:</p>
<p>
<code><pre>
Private Sub ProcessExportTable(ExportDirectory As IMAGE_DATA_DIRECTORY)
Dim deThis As IMAGE_EXPORT_DIRECTORY
Dim lBytesWritten As Long
Dim lpAddress As Long
Dim nFunction As Long
If ExportDirectory.VirtualAddress > 0 And ExportDirectory.Size > 0 Then
 '\\ Get the true address from the RVA
 lpAddress = AbsoluteAddress(ExportDirectory.VirtualAddress)
 '\\ Copy the image_export_directory structure...
 Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(deThis), Len(deThis), lBytesWritten)
 With deThis
  If .lpName <> 0 Then
   image.Name = StringFromOutOfProcessPointer(DebugProcess.Handle, image.AbsoluteAddress(.lpName), 32, False)
  End If
  If .NumberOfFunctions > 0 Then
   For nFunction = 1 To .NumberOfFunctions
    lpAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(.lpAddressOfNames) + ((nFunction - 1) * 4))
    fExport.Name = StringFromOutOfProcessPointer(DebugProcess.Handle, image.AbsoluteAddress(lpAddress), 64, False)
    fExport.Ordinal = .Base + IntegerFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(.lpAddressOfNameOrdinals) + ((nFunction - 1) * 2))
    fExport.ProcAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(.lpAddressOfFunctions) + ((nFunction - 1) * 4))
   Next nFunction
  End If
 End With
End If
End Sub
</pre></code>
</p>
<h3>The imports directory</h3>
<p>The imports directory lists the dynamic link libraries that this executable depends on and which functions it imports from that dynamic link library.
It consists of an array of <b>IMAGE_IMPORT_DESCRIPTOR</b> structures terminated by an instance of this structure where the <b>lpName</b> parameter is zero.
The structure is defined as:
</p>
<p>
<code><pre>
Private Type IMAGE_IMPORT_DESCRIPTOR
 lpImportByName As Long ''\\ 0 for terminating null import descriptor
 TimeDateStamp As Long ''\\ 0 if not bound,
       ''\\ -1 if bound, and real date\time stamp
       ''\\ in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
       ''\\ O.W. date/time stamp of DLL bound to (Old BIND)
 ForwarderChain As Long ''\\ -1 if no forwarders
 lpName As Long
 lpFirstThunk As Long ''\\ RVA to IAT (if bound this IAT has actual addresses)
End Type
</pre></code>
</p>
<p>And you can walk the import directory thus:</p>
<p>
<code><pre>
Private Sub ProcessImportTable(ImportDirectory As IMAGE_DATA_DIRECTORY)
Dim lpAddress As Long
Dim diThis As IMAGE_IMPORT_DESCRIPTOR
Dim byteswritten As Long
Dim sName As String
Dim lpNextName As Long
Dim lpNextThunk As Long
Dim lImportEntryIndex As Long
Dim nOrdinal As Integer
Dim lpFuncAddress As Long
'\\ If the image has an imports section...
If ImportDirectory.VirtualAddress > 0 And ImportDirectory.Size > 0 Then
 '\\ Get the true address from the RVA
 lpAddress = AbsoluteAddress(ImportDirectory.VirtualAddress)
 Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(diThis), Len(diThis), byteswritten)
 While diThis.lpName <> 0
  '\\ Process this import directory entry
  sName = StringFromOutOfProcessPointer(DebugProcess.Handle, image.AbsoluteAddress(diThis.lpName), 32, False)
  '\\ Process the import file's functions list
  If diThis.lpImportByName <> 0 Then
   lpNextName = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(diThis.lpImportByName))
   lpNextThunk = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(diThis.lpFirstThunk))
   While (lpNextName <> 0) And (lpNextThunk <> 0)
    '\\ get the function address
    lpFuncAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, lpNextThunk)
    nOrdinal = IntegerFromOutOfprocessPointer(DebugProcess.Handle, lpNextName)
    '\\ Skip the two-byte ordinal hint
    lpNextName = lpNextName + 2
    '\\ Get this function's name
    sName = StringFromOutOfProcessPointer(DebugProcess.Handle, image.AbsoluteAddress(lpNextName), 64, False)
    If Trim$(sName) <> "" Then
     '\\ Get the next imported function...
     lImportEntryIndex = lImportEntryIndex + 1
     lpNextName = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(diThis.lpImportByName + (lImportEntryIndex * 4)))
     lpNextThunk = LongFromOutOfprocessPointer(DebugProcess.Handle, image.AbsoluteAddress(diThis.lpFirstThunk + (lImportEntryIndex * 4)))
    Else
     lpNextName = 0
    End If
   Wend
  End If
  '\\ And get the next one
  lpAddress = lpAddress + Len(diThis)
  Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(diThis), Len(diThis), byteswritten)
 Wend
End If
End Sub
</pre></code>
</p>
<h3>The resource directory</h3>
<p>The structure of the resource director is somewhat more involved. It consists of a root directory (defined by the structure
<b>IMAGE_RESOURCE_DIRECTORY</b> immediately followed by a number of resource directory entries (defined by the structure <b>
IMAGE_RESOURCE_DIRECTORY_ENTRY</b>). These are defined thus:<p>
<p>
<code><pre>
Private Type IMAGE_RESOURCE_DIRECTORY
 Characteristics As Long '\\Seems to be always zero?
 TimeDateStamp As Long
 MajorVersion As Integer
 MinorVersion As Integer
 NumberOfNamedEntries As Integer
 NumberOfIdEntries As Integer
End Type
Private Type IMAGE_RESOURCE_DIRECTORY_ENTRY
 dwName As Long
 dwDataOffset As Long
 CodePage As Long
 Reserved As Long
End Type
</pre></code>
</p>
<p>Each resource directory entry can either point to the actual resource data or to another layer of resource directory entries. If the highest bit of
<b>dwDataOffset</b> is set then this points to a directory otherwise it points to the resource data.</p>
<h2>How is this information useful?</h2>
<p>Once you know how an executable is put together you can use this information to peer into its workings. You can view the resources compiled into
it, the dlls it depends on and the actual functions it imports from them. More importantly you can attach to the executable as a debugger and track down any of those really troublesome general protection faults. The next article will describe how to attach a debugger and use the PE file format.</p>
</font>

