;InnoSetupVersion=4.2.6

[Setup]
AppName=Multipot
AppVerName=Multipot v0.3
DefaultDirName=c:\iDefense\MultiPot\
DefaultGroupName=Multipot
OutputBaseFilename=Multipot_setup
OutputDir=./

[Files]
Source: ./dependancy\adoKit.dll; DestDir: {win}; Flags: uninsneveruninstall regserver ignoreversion
Source: ./honeypot.mdb; DestDir: c:\honeypot
Source: ./dependancy\vbDevKit.dll; DestDir: {win}; Flags: uninsneveruninstall regserver ignoreversion
Source: ./dependancy\MSWINSCK.OCX; DestDir: {sys}; Flags: uninsneveruninstall regserver onlyifdoesntexist
Source: ./dependancy\mscomctl.ocx; DestDir: {sys}; Flags: uninsneveruninstall regserver onlyifdoesntexist
Source: ./\multipot_help.chm; DestDir: {app}
Source: ./\multi_pot.exe; DestDir: {app}
Source: ./\source\CAntiHammer.cls; DestDir: {app}\source
Source: ./\source\CCmdEmulator.cls; DestDir: {app}\source
Source: ./\source\CFtpGet.frm; DestDir: {app}\source
Source: ./\source\CGenericURL.cls; DestDir: {app}\source
Source: ./\source\CHost.cls; DestDir: {app}\source
Source: ./\source\CLsassCmd.frm; DestDir: {app}\source
Source: ./\source\CLsassRecvFile.frm; DestDir: {app}\source
Source: ./\source\clsBagle.cls; DestDir: {app}\source
Source: ./\source\clsBagleDownloader.frm; DestDir: {app}\source
Source: ./\source\clsBagleFtpRecv.cls; DestDir: {app}\source
Source: ./\source\clsKuang.cls; DestDir: {app}\source
Source: ./\source\clsMyDoom.cls; DestDir: {app}\source
Source: ./\source\clsOptix.cls; DestDir: {app}\source
Source: ./\source\clsServer.frm; DestDir: {app}\source
Source: ./\source\clsSub7.cls; DestDir: {app}\source
Source: ./\source\clsUpload.cls; DestDir: {app}\source
Source: ./\source\CSc_Tftp.cls; DestDir: {app}\source
Source: ./\source\CTFTPClient.frm; DestDir: {app}\source
Source: ./\source\cVeritas.cls; DestDir: {app}\source
Source: ./\source\cVeritas_II.cls; DestDir: {app}\source
Source: ./\source\frmHexEdit.frm; DestDir: {app}\source
Source: ./\source\frmMain.frx; DestDir: {app}\source
Source: ./\source\frmScTest.frm; DestDir: {app}\source
Source: ./\source\frmSearch.frm; DestDir: {app}\source
Source: ./\source\frmStats.frm; DestDir: {app}\source
Source: ./\source\honeypot_2.ico; DestDir: {app}\source
Source: ./\source\MAINUI.frx; DestDir: {app}\source
Source: ./\source\Module1.bas; DestDir: {app}\source
Source: ./\source\PNPCmd.frm; DestDir: {app}\source
Source: ./\source\Project1.vbw; DestDir: {app}\source
Source: ./\source\VeritasCmd.frm; DestDir: {app}\source
Source: ./\source\cLsass.cls; DestDir: {app}\source
Source: ./\source\MAINUI.frm; DestDir: {app}\source
Source: ./\source\Project1.vbp; DestDir: {app}\source
Source: ./\public_sc_Analysis\recvfile\1126_extract.idb; DestDir: {app}\public_sc_Analysis\recvfile
Source: ./\public_sc_Analysis\recvfile\1126_extract.dat; DestDir: {app}\public_sc_Analysis\recvfile
Source: ./\public_sc_Analysis\recvfile\1126_raw_capture.dat; DestDir: {app}\public_sc_Analysis\recvfile
Source: ./\public_sc_Analysis\recvcmd\raw_capture.dat; DestDir: {app}\public_sc_Analysis\recvcmd
Source: ./\public_sc_Analysis\recvcmd\extract.dat; DestDir: {app}\public_sc_Analysis\recvcmd
Source: ./\public_sc_Analysis\recvcmd\disasm.idb; DestDir: {app}\public_sc_Analysis\recvcmd
Source: ./\public_sc_Analysis\tftp\dcom_tftp.dat; DestDir: {app}\public_sc_Analysis\tftp
Source: ./\public_sc_Analysis\tftp\dcom_tftp.idb; DestDir: {app}\public_sc_Analysis\tftp
Source: ./\public_sc_Analysis\gereric_url\veritas_tftp\721206063.dat; DestDir: {app}\public_sc_Analysis\gereric_url\veritas_tftp
Source: ./\public_sc_Analysis\gereric_url\veritas_tftp\721206063.idb; DestDir: {app}\public_sc_Analysis\gereric_url\veritas_tftp
Source: ./\public_sc_Analysis\gereric_url\lsass_ftp_echo\237424482.dat; DestDir: {app}\public_sc_Analysis\gereric_url\lsass_ftp_echo
Source: ./\public_sc_Analysis\gereric_url\http_download\487301940.dat; DestDir: {app}\public_sc_Analysis\gereric_url\http_download
Source: ./\public_sc_Analysis\gereric_url\http_download\879818985.dat; DestDir: {app}\public_sc_Analysis\gereric_url\http_download
Source: ./\public_sc_Analysis\gereric_url\http_download\879_.dat; DestDir: {app}\public_sc_Analysis\gereric_url\http_download
Source: ./\public_sc_Analysis\gereric_url\http_download\233_korgo.dat; DestDir: {app}\public_sc_Analysis\gereric_url\http_download
Source: ./\public_sc_Analysis\gereric_url\Readme.txt; DestDir: {app}\public_sc_Analysis\gereric_url
Source: ./\public_sc_Analysis\veritas_connect_back_shell\veritas.sc; DestDir: {app}\public_sc_Analysis\veritas_connect_back_shell
Source: ./\public_sc_Analysis\veritas_connect_back_shell\veritas.sc.idb; DestDir: {app}\public_sc_Analysis\veritas_connect_back_shell
Source: ./\public_sc_Analysis\pnp\pnp_shell_1\pnp.dat; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_1
Source: ./\public_sc_Analysis\pnp\pnp_shell_1\pnp.sc.dmp; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_1
Source: ./\public_sc_Analysis\pnp\pnp_shell_1\pnp.sc.idb; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_1
Source: ./\public_sc_Analysis\pnp\pnp_shell_2\pnp_shell2.dat; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_2
Source: ./\public_sc_Analysis\pnp\pnp_shell_2\pnp_shell2.sc; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_2
Source: ./\public_sc_Analysis\pnp\pnp_shell_2\pnp_shell2.sc.dmp; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_2
Source: ./\public_sc_Analysis\pnp\pnp_shell_2\pnp_shell2.sc.idb; DestDir: {app}\public_sc_Analysis\pnp\pnp_shell_2

[Dirs]
Name: c:\honeypot\
Name: {app}\source
Name: {app}\public_sc_Analysis
Name: {app}\public_sc_Analysis\recvfile
Name: {app}\public_sc_Analysis\recvcmd
Name: {app}\public_sc_Analysis\tftp
Name: {app}\public_sc_Analysis\gereric_url
Name: {app}\public_sc_Analysis\gereric_url\veritas_tftp
Name: {app}\public_sc_Analysis\gereric_url\lsass_ftp_echo
Name: {app}\public_sc_Analysis\gereric_url\http_download
Name: {app}\public_sc_Analysis\veritas_connect_back_shell
Name: {app}\public_sc_Analysis\pnp
Name: {app}\public_sc_Analysis\pnp\pnp_shell_1
Name: {app}\public_sc_Analysis\pnp\pnp_shell_2

[Run]
Filename: {app}\multipot_help.chm; WorkingDir: {app}; StatusMsg: View Help File; Flags: shellexec postinstall

[Icons]
Name: {group}\Multipot.exe; Filename: {app}\multi_pot.exe; WorkingDir: {app}
Name: {group}\Multipot_help; Filename: {app}\multipot_help.chm
Name: {group}\Source\Multi_Pot.vbp; Filename: {app}\source\Project1.vbp
Name: {group}\Uninstall; Filename: {app}\unins000.exe

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...
