# ePodaci harok XML generator pre Slovensku postu
# Verzia: 2.0

# Nastavenie kodowania na UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Konfiguracny subor pre ulozenie udajov odosielatela
$configPath = Join-Path $PSScriptRoot "odosielatel_config.xml"

function Save-OdosielatelConfig {
    param($config)
    
    $config | Export-Clixml -Path $configPath -Encoding UTF8
    Write-Host "Udaje odosielatela boli ulozene pre buduce pouzitie." -ForegroundColor Green
}

function Load-OdosielatelConfig {
    if (Test-Path $configPath) {
        try {
            return Import-Clixml -Path $configPath
        } catch {
            return $null
        }
    }
    return $null
}

function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$Default = "",
        [bool]$Required = $false
    )
    
    if ($Default) {
        $fullPrompt = "{0} [{1}]" -f $Prompt, $Default
    } else {
        $fullPrompt = $Prompt
    }
    
    do {
        $userValue = Read-Host $fullPrompt
        if ([string]::IsNullOrWhiteSpace($userValue) -and $Default) {
            $userValue = $Default
        }
        
        if ($Required -and [string]::IsNullOrWhiteSpace($userValue)) {
            Write-Host "Toto pole je povinne!" -ForegroundColor Red
        }
    } while ($Required -and [string]::IsNullOrWhiteSpace($userValue))
    
    return $userValue
}

function Get-DruhZasielky {
    Write-Host ""
    Write-Host "=== DRUH ZASIELKY ===" -ForegroundColor Cyan
    Write-Host "1  - Doporuceny list"
    Write-Host "2  - Poisteny list"
    Write-Host "3  - Uradna zasielka"
    Write-Host "4  - Balik"
    Write-Host "8  - Expres kurier"
    Write-Host "10 - EMS"
    Write-Host "11 - EPG - Obchodny balik"
    Write-Host "14 - Balik - zmluvni zakaznici"
    Write-Host "15 - Easy Expres 1"
    Write-Host "16 - Easy Expres 10"
    Write-Host "30 - List"
    Write-Host "33 - Balicek"
    
    $validValues = @(1,2,3,4,8,10,11,14,15,16,30,33)
    do {
        $druh = Read-Host "Vyberte druh zasielky"
        if ($druh -notin $validValues) {
            Write-Host "Neplatna volba!" -ForegroundColor Red
        }
    } while ($druh -notin $validValues)
    
    return $druh
}

function Get-SposobUhrady {
    Write-Host ""
    Write-Host "=== SPOSOB UHRADY ===" -ForegroundColor Cyan
    Write-Host "1 - Postovne uverovane"
    Write-Host "2 - Vyplatny stroj"
    Write-Host "3 - Platene prevodom"
    Write-Host "4 - Postove znamky"
    Write-Host "5 - Platene v hotovosti"
    Write-Host "7 - Vec postovej sluzby"
    Write-Host "8 - Faktura"
    Write-Host "9 - Online (platba kartou)"
    
    $validValues = @(1,2,3,4,5,7,8,9)
    do {
        $sposob = Read-Host "Vyberte sposob uhrady"
        if ($sposob -notin $validValues) {
            Write-Host "Neplatna volba!" -ForegroundColor Red
        }
    } while ($sposob -notin $validValues)
    
    return $sposob
}

function Get-Sluzby {
    Write-Host ""
    Write-Host "=== DOPLNKOVE SLUZBY ===" -ForegroundColor Cyan
    Write-Host "Dostupne sluzby (zadajte cisla oddelene ciarkou, alebo Enter pre preskocenie):"
    Write-Host "1  - D   (Dorucenka)"
    Write-Host "2  - DOH (Dorucit do 10:00)"
    Write-Host "3  - F   (Krehke)"
    Write-Host "4  - IOD (Info o doruceni)"
    Write-Host "5  - NDO (Nedoposielat)"
    Write-Host "6  - NEU (Neukladat)"
    Write-Host "7  - NEV (Nevratit)"
    Write-Host "8  - NSK (Neskladne)"
    Write-Host "9  - OD  (Opakovane dorucenie)"
    Write-Host "10 - OS  (Odpovedna sluzba)"
    Write-Host "11 - PR  (Na postu)"
    Write-Host "12 - PUZ (Podaj u kuriera)"
    Write-Host "13 - SV  (Splnomocnenie vylucene)"
    Write-Host "14 - SVD (Spatne vratenie potvr. dokumentacie)"
    Write-Host "15 - VR  (Do vlastnych ruk)"
    Write-Host "16 - VT  (Vymena tovaru)"
    
    $sluzbyCodes = @{
        1="D"; 2="DOH"; 3="F"; 4="IOD"; 5="NDO"; 6="NEU"; 7="NEV"; 8="NSK";
        9="OD"; 10="OS"; 11="PR"; 12="PUZ"; 13="SV"; 14="SVD"; 15="VR"; 16="VT"
    }
    
    $userInput = Read-Host "Vyberte sluzby (napr. 1,4,15)"
    $sluzby = @()
    
    if (![string]::IsNullOrWhiteSpace($userInput)) {
        $cisla = $userInput -split ','
        foreach ($cislo in $cisla) {
            $c = [int]$cislo.Trim()
            if ($sluzbyCodes.ContainsKey($c)) {
                $sluzby += $sluzbyCodes[$c]
            }
        }
    }
    
    return $sluzby
}

# Hlavny skript
Clear-Host
Write-Host "============================================" -ForegroundColor Green
Write-Host "  ePodaci harok XML generator" -ForegroundColor Green
Write-Host "  Slovenska posta" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

# Informacie o odosielatelovi
Write-Host ""
Write-Host "=== UDAJE ODOSIELATELA ===" -ForegroundColor Cyan

# Pokus o nacitanie ulozenych udajov
$savedConfig = Load-OdosielatelConfig

if ($savedConfig) {
    Write-Host "Nasli sa ulozene udaje odosielatela:" -ForegroundColor Yellow
    Write-Host "Meno: $($savedConfig.Meno)" -ForegroundColor Gray
    Write-Host "Organizacia: $($savedConfig.Organizacia)" -ForegroundColor Gray
    Write-Host "Adresa: $($savedConfig.Ulica), $($savedConfig.Mesto), $($savedConfig.PSC)" -ForegroundColor Gray
    $pouzitUlozene = Read-Host "Chcete pouzit tieto udaje? (a/n)"
    
    if ($pouzitUlozene -eq "a") {
        $odosielatelMeno = $savedConfig.Meno
        $odosielatelOrganizacia = $savedConfig.Organizacia
        $odosielatelUlica = $savedConfig.Ulica
        $odosielatelMesto = $savedConfig.Mesto
        $odosielatelPSC = $savedConfig.PSC
        $odosielatelKrajina = $savedConfig.Krajina
        $odosielatelTelefon = $savedConfig.Telefon
        $odosielatelEmail = $savedConfig.Email
        Write-Host "Udaje odosielatela boli nacitane." -ForegroundColor Green
    } else {
        # Zadanie novych udajov
        $odosielatelMeno = Get-UserInput "Meno odosielatela (tituly, meno, priezvisko)"
        $odosielatelOrganizacia = Get-UserInput "Organizacia"
        $odosielatelUlica = Get-UserInput "Ulica a cislo" -Required $true
        $odosielatelMesto = Get-UserInput "Mesto" -Required $true
        $odosielatelPSC = Get-UserInput "PSC" -Required $true
        $odosielatelKrajina = Get-UserInput "Krajina (ISO kod)" -Default "SK"
        $odosielatelTelefon = Get-UserInput "Telefon (0912345678)"
        $odosielatelEmail = Get-UserInput "Email" -Required $true
        
        # Ulozenie novych udajov
        $configToSave = @{
            Meno = $odosielatelMeno
            Organizacia = $odosielatelOrganizacia
            Ulica = $odosielatelUlica
            Mesto = $odosielatelMesto
            PSC = $odosielatelPSC
            Krajina = $odosielatelKrajina
            Telefon = $odosielatelTelefon
            Email = $odosielatelEmail
        }
        Save-OdosielatelConfig $configToSave
    }
} else {
    # Prve spustenie - zadanie udajov
    $odosielatelMeno = Get-UserInput "Meno odosielatela (tituly, meno, priezvisko)"
    $odosielatelOrganizacia = Get-UserInput "Organizacia"
    $odosielatelUlica = Get-UserInput "Ulica a cislo" -Required $true
    $odosielatelMesto = Get-UserInput "Mesto" -Required $true
    $odosielatelPSC = Get-UserInput "PSC" -Required $true
    $odosielatelKrajina = Get-UserInput "Krajina (ISO kod)" -Default "SK"
    $odosielatelTelefon = Get-UserInput "Telefon (0912345678)"
    $odosielatelEmail = Get-UserInput "Email" -Required $true
    
    # Ulozenie udajov
    $configToSave = @{
        Meno = $odosielatelMeno
        Organizacia = $odosielatelOrganizacia
        Ulica = $odosielatelUlica
        Mesto = $odosielatelMesto
        PSC = $odosielatelPSC
        Krajina = $odosielatelKrajina
        Telefon = $odosielatelTelefon
        Email = $odosielatelEmail
    }
    Save-OdosielatelConfig $configToSave
}

# Druh zasielky
$druhZasielky = Get-DruhZasielky

# Sposob uhrady
$sposobUhrady = Get-SposobUhrady

# Pocet zasielok
Write-Host ""
Write-Host "=== ZASIELKY ===" -ForegroundColor Cyan
[int]$pocetZasielok = Get-UserInput "Kolko zasielok chcete pridat?" -Default "1" -Required $true

# Aktualny datum
$datum = Get-Date -Format "yyyyMMdd"

# Kolekcia zasielok
$zasielky = @()

for ($i = 1; $i -le $pocetZasielok; $i++) {
    Write-Host ""
    Write-Host "--- ZASIELKA $i/$pocetZasielok ---" -ForegroundColor Yellow
    
    # Adresat
    Write-Host "Udaje adresata:" -ForegroundColor White
    $adresatMeno = Get-UserInput "Meno adresata"
    $adresatOrganizacia = Get-UserInput "Organizacia adresata"
    $adresatUlica = Get-UserInput "Ulica a cislo" -Required $true
    $adresatMesto = Get-UserInput "Mesto" -Required $true
    $adresatPSC = Get-UserInput "PSC" -Required $true
    $adresatKrajina = Get-UserInput "Krajina (ISO kod)" -Default "SK"
    $adresatTelefon = Get-UserInput "Telefon (nepovinne)"
    $adresatEmail = Get-UserInput "Email (nepovinne)"
    
    # Parametre zasielky
    Write-Host ""
    Write-Host "Parametre zasielky:" -ForegroundColor White
    $hmotnost = Get-UserInput "Hmotnost v kg (napr. 0.5)" -Default "0.0"
    $poznamka = Get-UserInput "Poznamka"
    
    # Dobierka
    $maDobierku = Read-Host "Ma zasielka dobierku? (a/n)"
    $cenaDobierky = ""
    $druhPPP = ""
    if ($maDobierku -eq "a") {
        $cenaDobierky = Get-UserInput "Suma dobierky (EUR)" -Required $true
        Write-Host "Sposob vyplaty dobierky:"
        Write-Host "5 - Bezdokladova dobierka na ucet"
        Write-Host "6 - Bezdokladova dobierka na adresu"
        $druhPPP = Get-UserInput "Vyberte" -Default "5"
    }
    
    # Poistenie
    $maPoistenie = Read-Host "Ma zasielka poistenie? (a/n)"
    $cenaPoistneho = ""
    if ($maPoistenie -eq "a") {
        $cenaPoistneho = Get-UserInput "Suma poistenia (EUR)" -Required $true
    }
    
    # Sluzby
    $sluzby = Get-Sluzby
    
    # Vytvorenie objektu zasielky
    $zasielka = @{
        AdresatMeno = $adresatMeno
        AdresatOrganizacia = $adresatOrganizacia
        AdresatUlica = $adresatUlica
        AdresatMesto = $adresatMesto
        AdresatPSC = $adresatPSC
        AdresatKrajina = $adresatKrajina
        AdresatTelefon = $adresatTelefon
        AdresatEmail = $adresatEmail
        Hmotnost = $hmotnost
        Poznamka = $poznamka
        CenaDobierky = $cenaDobierky
        DruhPPP = $druhPPP
        CenaPoistneho = $cenaPoistneho
        Sluzby = $sluzby
    }
    
    $zasielky += $zasielka
}

# Vytvorenie XML
Write-Host ""
Write-Host "=== GENEROVANIE XML ===" -ForegroundColor Cyan

# Zaciatok XML
$xmlLines = @()
$xmlLines += '<?xml version="1.0" encoding="UTF-8"?>'
$xmlLines += '<EPH verzia="3.0" xmlns="http://ekp.posta.sk/LOGIS/Formulare/Podaj_v03">'
$xmlLines += '  <InfoEPH>'
$xmlLines += '    <Mena>EUR</Mena>'
$xmlLines += '    <TypEPH>1</TypEPH>'
$xmlLines += '    <EPHID></EPHID>'
$xmlLines += "    <Datum>$datum</Datum>"
$xmlLines += "    <PocetZasielok>$pocetZasielok</PocetZasielok>"
$xmlLines += '    <Uhrada>'
$xmlLines += "      <SposobUhrady>$sposobUhrady</SposobUhrady>"
$xmlLines += '      <SumaUhrady>0.00</SumaUhrady>'
$xmlLines += '    </Uhrada>'
$xmlLines += "    <DruhZasielky>$druhZasielky</DruhZasielky>"
$xmlLines += '    <Odosielatel>'
$xmlLines += '      <OdosielatelID></OdosielatelID>'
$xmlLines += "      <Meno>$odosielatelMeno</Meno>"
$xmlLines += "      <Organizacia>$odosielatelOrganizacia</Organizacia>"
$xmlLines += "      <Ulica>$odosielatelUlica</Ulica>"
$xmlLines += "      <Mesto>$odosielatelMesto</Mesto>"
$xmlLines += "      <PSC>$odosielatelPSC</PSC>"
$xmlLines += "      <Krajina>$odosielatelKrajina</Krajina>"
$xmlLines += "      <Telefon>$odosielatelTelefon</Telefon>"
$xmlLines += "      <Email>$odosielatelEmail</Email>"
$xmlLines += '    </Odosielatel>'
$xmlLines += '  </InfoEPH>'
$xmlLines += '  <Zasielky>'

# Pridanie jednotlivych zasielok
foreach ($z in $zasielky) {
    $xmlLines += '    <Zasielka>'
    $xmlLines += '      <Adresat>'
    $xmlLines += "        <Meno>$($z.AdresatMeno)</Meno>"
    $xmlLines += "        <Organizacia>$($z.AdresatOrganizacia)</Organizacia>"
    $xmlLines += "        <Ulica>$($z.AdresatUlica)</Ulica>"
    $xmlLines += "        <Mesto>$($z.AdresatMesto)</Mesto>"
    $xmlLines += "        <PSC>$($z.AdresatPSC)</PSC>"
    $xmlLines += "        <Krajina>$($z.AdresatKrajina)</Krajina>"
    $xmlLines += "        <Telefon>$($z.AdresatTelefon)</Telefon>"
    $xmlLines += "        <Email>$($z.AdresatEmail)</Email>"
    $xmlLines += '      </Adresat>'
    $xmlLines += '      <Info>'
    $xmlLines += "        <Hmotnost>$($z.Hmotnost)</Hmotnost>"
    
    if ($z.CenaDobierky) {
        $xmlLines += "        <CenaDobierky>$($z.CenaDobierky)</CenaDobierky>"
        $xmlLines += "        <DruhPPP>$($z.DruhPPP)</DruhPPP>"
    }
    
    if ($z.CenaPoistneho) {
        $xmlLines += "        <CenaPoistneho>$($z.CenaPoistneho)</CenaPoistneho>"
    }
    
    if ($z.Poznamka) {
        $xmlLines += "        <Poznamka>$($z.Poznamka)</Poznamka>"
    }
    
    $xmlLines += '      </Info>'
    
    # Pridanie sluzieb
    if ($z.Sluzby.Count -gt 0) {
        $xmlLines += '      <PouziteSluzby>'
        foreach ($sluzba in $z.Sluzby) {
            $xmlLines += "        <Sluzba>$sluzba</Sluzba>"
        }
        $xmlLines += '      </PouziteSluzby>'
    }
    
    $xmlLines += '    </Zasielka>'
}

# Ukoncenie XML
$xmlLines += '  </Zasielky>'
$xmlLines += '</EPH>'

# Spojenie do jedneho retazca
$xml = $xmlLines -join "`n"

# Ulozenie suboru
$defaultFileName = "epodaci_harok_$datum.xml"
$fileName = Get-UserInput "Nazov vystupneho suboru" -Default $defaultFileName

# Ak nazov suboru neobsahuje cestu, pouzit aktualny priecinok
if (-not [System.IO.Path]::IsPathRooted($fileName)) {
    $fileName = Join-Path (Get-Location) $fileName
}

# Ulozenie s UTF-8 bez BOM
try {
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($fileName, $xml, $utf8WithoutBom)
    
    $fullPath = $fileName
    
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "XML subor bol uspesne vytvoreny!" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Cela cesta k suboru:" -ForegroundColor Cyan
    Write-Host $fullPath -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Subor mozete nahrat na www.posta.sk v sekcii ePodaci harok." -ForegroundColor Cyan
    
    # Ponuka na zobrazenie obsahu
    Write-Host ""
    $zobraz = Read-Host "Chcete zobrazit obsah XML? (a/n)"
    if ($zobraz -eq "a") {
        Write-Host ""
        Write-Host "--- OBSAH XML ---" -ForegroundColor Yellow
        Get-Content $fileName
    }
} catch {
    Write-Host ""
    Write-Host "CHYBA: Nepodarilo sa vytvorit XML subor!" -ForegroundColor Red
    Write-Host "Detaily: $($_.Exception.Message)" -ForegroundColor Red
}

# Cakanie na Enter pred zatvorenim
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "Stlacte Enter pre ukoncenie programu..." -ForegroundColor Cyan
Read-Host