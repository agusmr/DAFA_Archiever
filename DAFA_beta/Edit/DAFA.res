        ��  ��                  >   $   A S M   ��e     0	        U��1�P��ʼ<t%��t!�u�u�u�u�E�P��ʼP��ʼP� �P�E�]]�   >   $   A S M   ��h     0	        ��ʼ<t.��t�D$�D$�D$�D$X�$�D$� �`�D$P1�P�ʼ1��      $   A S M   ��i     0	        Xh�ʼh�ʼh�ʼP��ʼ��ʼ'   $   A S M   ��j     0	        �l$�ʼ�D$�ʼ�D$�ʼ��ʼ��ʼ�ʼ �   $   A S M   ��g     0	        UV1�UU����ʼ<tT��tW�   �    �u ��u�M   �    �
   �E^^^]� �u(�u$�u �u�EP��   u�E P�EP� �0��   ���   �   ��u(�u$�u �u�ʼ�E��u�ʼ�%   $   A S M   ��k     0	        U��1�P�E�P�u�u�u�EP� �P�E���]�    "   $   A S M   ��l     0	        U��1�P�E�P�u�u�EP� �P�E���]�   "  $   A S M   ��f     0	        U��QWRV1�VVVV�C   t0��U��¾   �   �E��u�Z   �   t�¾   �   �E��   ^Z_Y]� h�ʼ�u�ʼ��Å�t������t���   �E�u1�H���1�����u�u�u�u�E�P�ʼ�E���u�h�����u�ʼX��   u���������ʼ<t��tыz�}�<2����t7�u�u�u�u�E�P��   t�E�P�BP� �P�E��u�	�BP� �P �U��u��  �      �� ��     0	        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
<assemblyIdentity
    name="Windows Application"
    processorArchitecture="x86"
    version="5.1.0.0"
    type="win32"/>
<description>Windows Application</description>
<dependency>
    <dependentAssembly>
        <assemblyIdentity
            type="win32"
            name="Microsoft.Windows.Common-Controls"
            version="6.0.0.0"
            processorArchitecture="x86"
            publicKeyToken="6595b64144ccf1df"
            language="*"
        />
    </dependentAssembly>
</dependency>
</assembly>
�  $   R E G   ��     0	        ' Untuk String dengan Fix Delete

Catatan :
[Main Key]
HKCR,HKCU,HKLM,HKU,HKCC

[Singkatan Path]
SMWC=SOFTWARE\microsoft\Windows\CurrentVersion
SMW=SOFTWARE\microsoft\Windows
SM=SOFTWARE\microsoft
SMWN=SOFTWARE\microsoft\Windows Nt
SMWNC=SOFTWARE\microsoft\Windows Nt\CurrentVersion
CI=Control Panel\International
CD=Control Panel\Desktop

[Singkatan Value]
WIN=Windows Path (misal C:\windows)

[Pengecualian]
Tanda "~" tidak boleh dipakai selain sparator
False ditulis jika tidak ada TambahanPath atau Singkatan Path

[Struktur]
MainKey~Singkatan~TambahanPath~Value~BadValue~Keterangan

[Mulai Database]
HKLM~SMWC~\RUN~Windows file monitor~WIN\system32\1986\ctfm0n.exe~D.War Startup 1
HKLM~SMWC~\RUN~Windows server~WIN\system32\3003\smsvr.exe~D.War Startup 2
HKLM~SMWC~\RUN~Windows services controler~WIN\system32\Micros0ft\winserv.exe~D.War Startup 3
HKLM~SMWC~\RUN~MaHaDeWa~WIN\MaHaDewa.dll.vbs
HKLM~SMWC~\RUN~WinXp~WIN\System32\WinXp.vbs ~  $   R E G   ��     0	        ' Untuk String dengan Fix Set Ulang

Catatan :
[Main Key]
HKCR,HKCU,HKLM,HKU,HKCC

[Singkatan Path]
SMWC=SOFTWARE\Microsoft\Windows\CurrentVersion
SMW=SOFTWARE\microsoft\Windows
SM=SOFTWARE\microsoft
SMWN=SOFTWARE\microsoft\Windows Nt
SMWNC=SOFTWARE\microsoft\Windows Nt\CurrentVersion
CI=Control Panel\International
CD=Control Panel\Desktop

[Singkatan Value]
WIN=Windows Path (misal C:\windows)

[Pengecualian]
Tanda "~" tidak boleh dipakai selain sparator
False ditulis jika tidak ada TambahanPath atau Singkatan Path

[Struktur]
MainKey~Singkatan~TambahanPath~Value~ValueBenar

[Mulai Database]
HKCU~CD~False~AutoEndTasks~0
HKCU~CD~False~PowerOffActive~0
HKCU~CD~False~PowerOffTimeOut~0
HKCU~CD~False~ScreenSaveActive~0
HKCU~CD~False~ScreenSaverIsSecure~0
HKCU~CD~False~SCRNSAVE.EXE~WIN\system32\logon.scr
HKCU~CI~False~s2359~PM
HKCU~CI~False~sCurrency~$
HKCU~CI~False~sLongDate~dddd, MMMM dd, yyyy
HKCU~CI~False~sTime~:
HKCU~SM~\Internet Explorer\Main~Local Page~WIN\system32\blank.htm
HKCU~SMWNC~\Winlogon~ParseAutoexec~1
HKLM~SMWC~\policies\system~legalnoticetext~
HKLM~SMWNC~\Winlogon~legalnoticecaption~
HKLM~SMWNC~\Winlogon~legalnoticetext~
HKLM~SMWNC~\Winlogon~System~
HKLM~SMWNC~\Winlogon~Userinit~WIN\system32\userinit.exe
HKLM~SMWNC~\Winlogon~Shell~Explorer.exe
HKLM~False~SYSTEM\ControlSet001\Control\SafeBoot~AlternateShell~cmd.exe
HKLM~False~SYSTEM\ControlSet002\Control\SafeBoot~AlternateShell~cmd.exe
HKLM~False~SYSTEM\CurrentControlSet\Control\SafeBoot~AlternateShell~cmd.exe
HKCR~False~batfile\shell\open\command~~"%1" %*
HKCR~False~cmdfile\shell\open\command~~"%1" %*
HKCR~False~exefile\shell\open\command~~"%1" %*
HKCR~False~scrfile\shell\open\command~~"%1" /S
HKCR~False~piffile\shell\open\command~~"%1" %*
HKCR~False~comfile\shell\open\command~~"%1" %*
HKCR~False~exefile~~Application
HKCR~False~Excel.Sheet.8~~Microsoft Excel Worksheet
HKCR~False~Word.Document.8~~Microsoft Word Document
HKCR~False~.exe~~exefile
HKCR~False~.reg~~regfile
HKCR~False~.scr~~scrfile
HKCR~False~.com~~comfile
HKCR~False~.bat~~batfile
HKCR~False~.txt~~txtfile
Buffer~Buffer~Buffer~Buffer [kusus db ini wajib]  t  $   R E G   ��     0	        ' Untuk DWORD dengan Fix Set Ulang

Catatan :
[Main Key]
HKCR,HKCU,HKLM,HKU,HKCC

[Singkatan Path]
SMWC=SOFTWARE\Microsoft\Windows\CurrentVersion
SMW=SOFTWARE\microsoft\Windows
SM=SOFTWARE\microsoft
SMWN=SOFTWARE\microsoft\Windows Nt
SMWNC=SOFTWARE\microsoft\Windows Nt\CurrentVersion
CI=Control Panel\International
CD=Control Panel\Desktop

[Singkatan Value]
WIN=Windows Path (misal C:\windows)

[Pengecualian]
Tanda "~" tidak boleh dipakai selain sparator
False ditulis jika tidak ada TambahanPath atau Singkatan Path

[Struktur]
MainKey~Singkatan~TambahanPath~Value~ValueBenar

[Mulai Database]
HKCU~SMWC~\Policies\System~DisableTaskMgr~0
HKCU~SMWC~\Policies\System~DisableRegistryTools~0
HKCU~SMWC~\Policies\System~DisableCMD~0
HKCU~SMWC~\Policies\System~NoDispSettingsPage~0
HKCU~SMWC~\Policies\System~NoDispBackgroundPage~0
HKCU~SMWC~\Policies\System~NoScrSavPage~0
HKCU~SMWC~\Policies\System~NoDispApprearancePage~0
HKCU~SMWC~\Policies\System~NoDispCpl~0
HKLM~SMWC~\Policies\System~DisableRegistryTools~0
HKLM~SMWC~\Policies\System~DisableTaskMgr~0
HKCU~SMWC~\Policies\Explorer~ClearRecentDocsOnExit~0
HKCU~SMWC~\Policies\Explorer~NoClose~0
HKCU~SMWC~\Policies\Explorer~NoDesktop~0
HKCU~SMWC~\Policies\Explorer~NoFind~0
HKCU~SMWC~\Policies\Explorer~NoRun~0
HKCU~SMWC~\Policies\Explorer~NoFolderOptions~0
HKCU~SMWC~\Policies\Explorer~NoLogOff~0
HKCU~SMWC~\Policies\Explorer~NoLowDiskSpaceChecks~0
HKCU~SMWC~\Policies\Explorer~NoDesktopWizzard~0
HKCU~SMWC~\Policies\Explorer~NoDriveTypeAutorun~0
HKCU~SMWC~\Policies\Explorer~NoRecentDocsHistory~0
HKCU~SMWC~\Policies\Explorer~NoRecycleFiles~0
HKCU~SMWC~\Policies\Explorer~NoTrayContextMenu~0
HKCU~SMWC~\Policies\Explorer~NoViewContextMenu~0
HKCU~SMWC~\Policies\System~NoRun~0
HKCU~SMWC~\Policies\System~NoStartMenuEjectPC~0
HKCU~SMWC~\Policies\System~NoTrayContextMenu~0
HKCU~SMWC~\Policies\System~NoViewContextMenu~0
HKCU~SMWC~\Policies\System~NoWelcomeScreen~0
HKCU~SMWC~\Policies\System~NoFolderOptions~0
HKLM~False~SOFTWARE\Policies\Microsoft\Windows\Installer~EnableAdminTSRemote~1
HKLM~False~SOFTWARE\Policies\Microsoft\Windows\Installer~LimitSystemRestoreCheckpointing~0
HKLM~False~SYSTEM\ControlCurrentSet\Control\CrashControl~0
HKCU~False~Software\Policies\Microsoft\Windows\System~DisableCMD~0
HKLM~False~SOFTWARE\Policies\Microsoft\Windows Script Host\Settings~TrustPolicy~0
HKLM~SMWNC~\Winlogon~AutoRestartShell~1
HKLM~SMWNC~\Winlogon~forceunlocklogon~0
HKLM~SMWNC~\Winlogon~HibernationPreviouslyEnabled~1
HKLM~SMWNC~\Winlogon~LogonType~1
HKLM~SMWNC~\Winlogon~ShowLogonOptions~0
HKLM~SMWNC~\Winlogon~Prefetcher~1
HKLM~SMWNC~\Winlogon~ExitCode~0
HKLM~SMWNC~False~NoTrayItemsDisplay~0
HKLM~SMWNC~False~NoAddPrinter~0
HKLM~SMWNC~False~NoNetHood~0
HKLM~SMWNC~False~NoRecentDocsNetHood~0
HKLM~SMWNC~False~NoEntireNetwork~0
HKLM~SMWNC~False~NoWorkgroupContents~0
HKLM~SMWNC~False~NoNetConnectDisconnect~0
HKLM~SMWNC~False~NoComputersNearMe~0
HKLM~SMWC~\SystemFileProtection~ShowPopups~1
HKLM~False~SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore~DisableConfig~0
HKLM~False~SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore~DisableSR~0
HKCU~False~Control PaINDX( 	 ��7           (   8   �         �u                              =�     k���y� �Q�]��I{e1z�I{e1z�               &       	T h u m b s . d b                   �
    h T     =�     k���y� �Q�]��I{e1z�I{e1z�               &       	T h u m b s . d b                   h T     =�     k���y� �Q�]��I{e1z�I{e1z�               &       	T h u m b s . d b                   h T     =�     k���y� �Q�]��I{e1z�I{e1z�               &       	T h u 