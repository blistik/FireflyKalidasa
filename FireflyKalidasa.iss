; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Firefly The PC Game"
#define MyAppVersion "1.8.5"
#define MyAppPublisher "VeeBee-er"
#define MyAppURL "https://boardgamegeek.com/thread/2996155/firefly-windows-pc-game"
#define MyAppExeName "FireflyKalidasa.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{EB56A73F-D1D8-44DE-896B-FCF27F5E7436}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=C:\Games\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputBaseFilename=FireflyKalidasaSetupV1.8.5
SetupIconFile=D:\Progs\GitHub\FireflyKalidasa\serenity.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "D:\Progs\GitHub\FireflyKalidasa\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\SFTTREEX.OCX"; DestDir: "{sys}"; Flags: 32bit regserver sharedfile 
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\MSCOMCTL.OCX"; DestDir: "{sys}"; Flags: 32bit regserver sharedfile
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\LaVolpeAlphaImg2.ocx"; DestDir: "{sys}"; Flags: 32bit regserver sharedfile
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\XDockFloat.dll"; DestDir: "{sys}"; Flags: 32bit regserver sharedfile

Source: "D:\Progs\GitHub\FireflyKalidasa\bin\SHOWG.TTF";  DestDir: "{autofonts}"; FontInstall: "Showcard Gothic"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\BRITANIC.TTF";  DestDir: "{autofonts}"; FontInstall: "Britannic Bold"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\CIND.otf";  DestDir: "{autofonts}"; FontInstall: "Cyberpunk Is Not Dead"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "D:\Progs\GitHub\FireflyKalidasa\bin\Forque.ttf";  DestDir: "{autofonts}"; FontInstall: "FORQUE"; Flags: onlyifdoesntexist uninsneveruninstall

Source: "D:\Progs\GitHub\FireflyKalidasa\FireFlyAIBot.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\VerseMapTool.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\FireflyKalidasa.mdb"; DestDir: "{app}"; Flags: ignoreversion; Permissions: users-modify
Source: "D:\Progs\GitHub\FireflyKalidasa\Firefly_rulebook.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\FireflyBlueSun_rulebook.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\FireflyKalidasa_rulebook.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\FireflyForPC.pdf"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\amnon.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Atherton.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken5.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\AToken6.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\badger.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Bester.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Billy.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\KalidasaBoard.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Bree.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Bridgit.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\browncoat.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Burgess.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\BurgessOrig.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\CargoBay.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Corbin.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Cortland.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Crow.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\CryBaby.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\delta.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Doralee.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Dress.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\ellie.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Emma.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Explosives.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Fendris.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\flyingmule.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Gran.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Grange.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\harken.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\harrow.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Helen.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Hotspot.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Inara.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Jayne.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\JaynesHat.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Jed.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Jesse.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\jethro.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Kaylee.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\KReprogammer.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\logo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Lucy.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Lund.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Malcolm.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\MaqueTiles.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Marco.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Monty.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\MrUniverse.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Murphy.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Nandi.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\niska.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\OriginalBoard.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\OverChg.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Parasol.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Patience.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Quarters.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\River.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken5.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RToken6.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Saffron.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Shepherd.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Sheydra.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Simon.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Skunk.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Skyhook.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmAtherton.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmBurgess.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmCorbin.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmMalcolm.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmMarco.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmMonty.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmMurphy.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmNandi.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmWomack.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Stark.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Stitch.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Tracey.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\timebomb.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Two-Fry.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Wash.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Womack.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Yolonda.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Zoe.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\wright.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmWright.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\boardgame.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\fantymingo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\inarasbow.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\higgins.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\4WDmule.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Vera.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\FireflyBlue.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\FireflyYellow.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\FireflyOrange.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\FireflyGreen.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Cruiser.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion           
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Corvette.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Cutter.gif"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\barkeepbob.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\barkeepbex.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\fess.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\4range.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\6range.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\ambo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\baton.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\billiards.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\BurgessLaser.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\clips.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\creds.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\crowsknife.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\dc6range.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\dc6rangeII.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\docs.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\flac.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\foam.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Glucklich.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\grenade.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\guild.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\hfdocs.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\horses.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\ident.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\ImproHack.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Intel.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Jacket.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\jaynesPistol.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\kit.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\knife.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\lovebot.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\malspistol.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\medbay.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\necktie.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\oatbar.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\pistol.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Radion.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\RimNav.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SecInterPad.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\shipupgd.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\shirt.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\sniper.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SpaceJeep.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\sword.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\washcharts.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\YolondaPistol.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\zoerifle.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\suit1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\suit2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\suit3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\suit4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D5.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\D6.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\haven1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\haven2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\haven3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\haven4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\fight.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\tech.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\nego.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Foreman.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\draper.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\roberta.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\medic.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\busker.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\middleman.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\maddi.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\max.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\maz.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\moo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\mic.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\moi.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\grimey.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\gunhand.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\fandancer.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\salesman.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\recycler.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smharken.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smlogo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smharrow.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmMrUniverse.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smniska.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmPatience.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smhiggins.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smfantymingo.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smamnon.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Smbadger.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\Jubal.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\CrewTemplate.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\GearBlank.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\MisbehaveTemplate.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\smsuit1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\smsuit2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\smsuit3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\smsuit4.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\skill1.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\skill2.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\skill3.bmp"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\bonnet.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmSimon.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmJayne.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmRiver.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmZoe.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmShepherd.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmKaylee.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmInara.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\pictures\SmWash.jpg"; DestDir: "{app}\pictures"; Flags: ignoreversion

Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Alert.WAV"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Beep.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Burn.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Cruiser.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\mosey.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\msg.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\no.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Reaver.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\Win.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\yourgo.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\gear.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\reload.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion
Source: "D:\Progs\GitHub\FireflyKalidasa\sounds\clack.wav"; DestDir: "{app}\sounds"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Firefly AI Bot"; Filename: "{app}\FireFlyAIBot.exe"
Name: "{group}\Firefly Rulebook"; Filename: "{app}\Firefly_rulebook.pdf"
Name: "{group}\BlueSun Rulebook"; Filename: "{app}\FireflyBlueSun_rulebook.pdf"
Name: "{group}\Kalidasa Rulebook"; Filename: "{app}\FireflyKalidasa_rulebook.pdf"
Name: "{group}\Firefly-for-PC Guide"; Filename: "{app}\FireflyForPC.pdf"

Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

