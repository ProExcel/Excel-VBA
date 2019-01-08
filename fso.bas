Sub ModelFSO()

    Dim objSlozka As Object
    Dim objPodslozka As Object
    Dim objSlozkaNova As Object
    Dim objSlozkaPodslozky As Object
    Dim objSoubor As Object

    Dim objFSO As Object
    'vytvoreni objektu FSO a jeho prirazeni do objektove promenne
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    'alternativne:
    'Tools / References / Microsoft Scripting Runtime
    'Dim objFSO As New FileSystemObject

    'aktualni slozka bez lomitka na konci
    'strAktualniSlozka = objFSO.GetAbsolutePathName("")
    'alternativne (pro zmenu ChDir)
    'strAktualniSlozka = CurDir
    'vetsinou vraci odkaz do slozky Dokumenty

    'cesta k tomuto sesitu
    strSlozka = ThisWorkbook.Path

    'sesit
    strSesit = ThisWorkbook.Name

    'cesta a sesit
    strSlozkaSesit = ThisWorkbook.FullName

    'nadrizena slozka bez lomitka na konci
    strSesitRodicovskaSlozka = objFSO.GetParentFolderName(strSlozkaSesit)

    'jednotka (pismeno disku) bez lomitka na konci
    strSesitJednotka = objFSO.GetDriveName(strSlozkaSesit)

    'nazev souboru, pripona souboru
    'existence neni overovana
    strSouborNazevBezPripony = objFSO.GetBaseName(strSlozkaSesit)
    strSouborPriponaBezTecky = objFSO.GetExtensionName(strSlozkaSesit)
    strSouborNazevSPriponou = objFSO.GetFileName(strSlozkaSesit)

    'alternativne
    'nazev souboru s priponou
    'overovana existence
    'strSouborNazevSPriponou = Dir("C:\Windows\System32\msconfig.exe")
    'mozne uzit zastupne znaky * a ?, vracen prvni vyhovujici soubor
    'strSouborNazevSPriponou = Dir("C:\Windows\System32\System32\ms*.exe")

    'existence jednotky
    blnJednotkaExistuje = objFSO.DriveExists("D:")

    'existence slozky
    blnSlozkaExistuje = objFSO.FolderExists("C:\Windows\")

    'alternativne
    blnSlozkaExistuje = Len(Dir("C:\Windows\", vbDirectory)) > 0

    'existence souboru
    blnSouborExistuje = objFSO.FileExists("C:\pagefile.sys")

    'alternativne, pouze pro bezne viditelny soubor
    've vychozim stavu parametr vbNormal
    'False
    blnSouborExistuje = Len(Dir("C:\pagefile.sys")) > 0

    'atributy -ahs
    'True
    blnSouborExistuje = Len(Dir("C:\pagefile.sys", vbArchive + vbHidden + _
        vbSystem)) > 0

    'vytvoreni neexistujicich(!) slozek Slozka1 a Slozka2
    Set objSlozka = objFSO.CreateFolder(strSlozka & "\Slozka1")
    Set objPodslozka = objFSO.CreateFolder(strSlozka & "\Slozka1\Slozka2")

    'alternativne
    'MkDir strSlozka & "\Slozka1"
    'MkDir strSlozka & "\Slozka1\Slozka2"

    'nebo
    'Set objSlozka = objFSO.CreateFolder(strSlozka & "\Slozka1")
    'Set objSlozkaPodslozky = objSlozka.SubFolders
    'Set objSlozkaNova = objSlozkaPodslozky.Add("Slozka2")

    'v jednom kroku
    'nelze pouzit CreateFolder ani MkDir pod VBA
    'prikazova radka
    'Shell ("cmd /c mkdir """ & strSlozka & "\Slozka1\Slozka2" & """")

    'vytvoreni (neexistujiciho) souboru Soubor1.txt
    'CreateTextFile(filename[, overwrite[, unicode]])
    'kodovani UCS-2 Little Endian
    'objSlozka.CreateTextFile "Soubor1.txt", True, True

    'alternativne
    'kodovani UTF-8 bez BOM
    objFSO.CreateTextFile (strSlozka & "\Slozka1\Soubor1.txt")

    'otevreni (vytvoreni) souboru pro zapis
    Set objSoubor = objFSO.OpenTextFile(strSlozka & "\Slozka1\Soubor1.txt", 2, _
        True, -2)
    '1 nebo ForReading ...pristup pro cteni
    '2 nebo ForWriting ... pristup pro zapis
    '8 nebo ForAppending ...zapis na konec souboru

    'True ...vytvori soubor, pokud neexistuje

    '-2 nebo TristateUseDefault ...vychozi format
    '-1 nebo TristateTrue...Unicode
    ' 0 nebo TristateFalse ...ASCII

    With objSoubor
        'prvni radek
        .WriteLine "Prvni poznamka"
        'dva prazdne radky
        .WriteBlankLines 2
        'dalsi radek
        .WriteLine "Druha poznamka"
        'uzavreni souboru
        .Close
    End With

    'otevreni (vytvoreni) souboru pro pridani obsahu na konec
    Set objSoubor = objFSO.OpenTextFile(strSlozka & "\Slozka1\Soubor1.txt", 8, _
        True, -2)

    With objSoubor
        .WriteLine "Treti poznamka"
        .Close
    End With

    'otevreni (vytvořeni) souboru pro cteni
    Set objSoubor = objFSO.OpenTextFile(strSlozka & "\Slozka1\Soubor1.txt", 1, _
        True, -2)

    'cteni ze souboru
    'Read(n) ... nacteni n znaku
    'ReadLine ... nacteni jedne radky
    'ReadAll ... nacteni celeho obsahu
    'Skip(n), SkipLine ... preskoceni n znaku ci celeho radku
    Do While Not (objSoubor.AtEndOfStream)
        strText = strText & objSoubor.ReadLine & vbCrLf
    Loop

    'zobrazeni obsahu souboru
    MsgBox strText

    'uzavreni souboru
    objSoubor.Close

    'cteni jinak
    'Set objSoubor = objFSO.GetFile(strSlozka & "\Slozka1\Soubor1.txt")
    'Set objStream = objSoubor.OpenAsTextStream()
    'zobrazeni obsahu souboru
    'MsgBox objStream.ReadAll
    'uzavreni souboru
    'objStream.Close

    'vymazani obsahu souboru
    'Shell Environ$("comspec") & " /c type > " & strSlozka & _
        "\Slozka1\Soubor1.txt", vbHide

    'odstraneni souboru
    'True znaci vcetne souboru ke cteni
    'objFSO.DeleteFile strSlozka & "\Slozka1\Soubor1.txt", True
    'nebo
    'objSoubor.Delete , True

    'alternativne
    'mozne pouzit zastupne znaky *, ?
    'Kill strSlozka & "Slozka1\Soubor1.txt"

    'otevreni textoveho souboru v Poznamkovem bloku
    Shell "Notepad " & strSlozka & "\Slozka1\Soubor1.txt", 1

    'tisk textoveho souboru s pomoci Poznamkoveho bloku
    'Shell "Notepad /p " & strSlozka & "\Slozka1\Soubor1.txt", 0

    'kopie jednoho souboru
    's pripadnym vytvorenim cilove slozky, prejmenovanim
    'objFSO.CopyFile strSlozka & "\Slozka1\Soubor1.txt", strSlozka & _
        "\Slozka1\Slozka2\"
    objFSO.CopyFile strSlozka & "\Slozka1\Soubor1.txt", strSlozka & _
        "\Slozka1\Slozka2\Soubor2.txt"

    'nebo
    'Set objSoubor = objFSO.GetFile(strSlozka & "\Slozka1\Soubor1.txt")
    'objSoubor.Copy strSlozka & "\Slozka1\Slozka2\"

    'prejmenovani souboru
    Set objSoubor = objFSO.GetFile(strSlozka & "\Slozka1\Slozka2\Soubor1.txt")
    objSoubor.Name = "Soubor2.txt"
    'pouziva se take metoda MoveFile
    'objFSO.MoveFile strSlozka & "\Slozka1\Slozka2\Soubor1.txt", strSlozka & _
        "\Slozka1\Slozka2\Soubor2.txt"

    'kopirovani kompletniho obsahu slozky Slozka1 primo na disk D
    'nevytvori se slozka Slozka1 na disku D
    'True ...dojde k prepisu pripadnych souboru
    'objFSO.CopyFolder strSlozka & "\Slozka1", "D:", True
    'nebo napr. objSlozka.Copy "D:", True

    'viz predchozi, navic se vytvori slozka Slozka1 na disku D
    'objFSO.CopyFolder strSlozka & "\Slozka1", "D:\", True

    'kopirovani podslozek slozky Slozka1 primo na disk D
    'objFSO.CopyFolder strSlozka & "\Slozka1\*", "D:", True

    'kopirovani podslozek ze slozky Slozka1 i s jejich
    'obsahem do jiz existujici slozky Slozka1 na disku D
    'objFSO.CopyFolder strSlozka & "\Slozka1\*", "D:\Slozka1\", True

    'kopirovani vsech souboru slozky Slozka1 do jiz
    'existujici slozky Slozka1 na disku D
    'objFSO.CopyFile strSlozka & "\Slozka1\*.*", "D:\Slozka1\", True

    'presun slozky
    'MoveFolder doprovazi prilis otazniku
    'je lepsi pouzivat Copy a Delete
    'objFSO.MoveFolder strSlozka & "\Slozka1", "D:"

    'prejmenovani slozky
    'Set objSlozka = objFSO.GetFolder("D:\Slozka1")
    'objSlozka.Name = strSlozka & "\Slozka3"

    'pripadne MoveFolder
    'objFSO.MoveFolder "D:\Slozka1", "D:\Slozka3"

    'alternativne
    'Name "D:\Slozka1" As "D:\Slozka3"

    'odstraneni slozky
    'True znaci vcetne slozek ke cteni
    objFSO.DeleteFolder strSlozka & "\Slozka1", True
    'nebo napr. objSlozka.Delete, True

    'reset objektovych promennych
    Set objSoubor = Nothing
    Set objSlozka = Nothing
    Set objFSO = Nothing

End Sub
