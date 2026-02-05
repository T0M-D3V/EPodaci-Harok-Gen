# ePodacÌ h·rok XML gener·tor pre Slovensk˙ poötu
# Verzia: 1.0

# Nastavenie kÛdovania na UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-UserInput {
    param(
        [string]$Prompt,
        [string]$Default = "",
        [bool]$Required = $false
    )
    
    if ($Default) {
        $fullPrompt = "{0} [{1}]: " -f $Prompt, $Default
    } else {
        $fullPrompt = "{0}: " -f $Prompt
    }
    
    do {
        $input = Read-Host $fullPrompt
        if ([string]::IsNullOrWhiteSpace($input) -and $Default) {
            $input = $Default
        }
        
        if ($Required -and [string]::IsNullOrWhiteSpace($input)) {
            Write-Host "Toto pole je povinnÈ!" -ForegroundColor Red
        }
    } while ($Required -and [string]::IsNullOrWhiteSpace($input))
    
    return $input
}

function Get-DruhZasielky {
    Write-Host "`n=== DRUH Z¡SIELKY ===" -ForegroundColor Cyan
    Write-Host "1  - DoporuËen˝ list"
    Write-Host "2  - Poisten˝ list"
    Write-Host "3  - ⁄radn· z·sielka"
    Write-Host "4  - BalÌk"
    Write-Host "8  - Expres kuriÈr"
    Write-Host "10 - EMS"
    Write-Host "11 - EPG ñ Obchodn˝ balÌk"
    Write-Host "14 - BalÌk ñ zmluvnÌ z·kaznÌci"
    Write-Host "15 - Easy Expres 1"
    Write-Host "16 - Easy Expres 10"
    Write-Host "30 - List"
    Write-Host "33 - BalÌËek"
    
    $validValues = @(1,2,3,4,8,10,11,14,15,16,30,33)
    do {
        $druh = Read-Host "Vyberte druh z·sielky"
        if ($druh -notin $validValues) {
            Write-Host "Neplatn· voæba!" -ForegroundColor Red
        }
    } while ($druh -notin $validValues)
    
    return $druh
}

function Get-SposobUhrady {
    Write-Host "`n=== SP‘SOB ⁄HRADY ===" -ForegroundColor Cyan
    Write-Host "1 - PoötovnÈ ˙verovanÈ"
    Write-Host "2 - V˝platn˝ stroj"
    Write-Host "3 - PlatenÈ prevodom"
    Write-Host "4 - PoötovÈ zn·mky"
    Write-Host "5 - PlatenÈ v hotovosti"
    Write-Host "7 - Vec poötovej sluûby"
    Write-Host "8 - Fakt˙ra"
    Write-Host "9 - Online (platba kartou)"
    
    $validValues = @(1,2,3,4,5,7,8,9)
    do {
        $sposob = Read-Host "Vyberte spÙsob ˙hrady"
        if ($sposob -notin $validValues) {
            Write-Host "Neplatn· voæba!" -ForegroundColor Red
        }
    } while ($sposob -notin $validValues)
    
    return $sposob
}

function Get-Sluzby {
    Write-Host "`n=== DOPLNKOV… SLUéBY ===" -ForegroundColor Cyan
    Write-Host "DostupnÈ sluûby (zadajte ËÌsla oddelenÈ Ëiarkou, alebo Enter pre preskoËenie):"
    Write-Host "1  - D   (DoruËenka)"
    Write-Host "2  - DOH (DoruËiù do 10:00)"
    Write-Host "3  - F   (KrehkÈ)"
    Write-Host "4  - IOD (Info o doruËenÌ)"
    Write-Host "5  - NDO (Nedoposielaù)"
    Write-Host "6  - NEU (Neukladaù)"
    Write-Host "7  - NEV (Nevr·tiù)"
    Write-Host "8  - NSK (NeskladnÈ)"
    Write-Host "9  - OD  (OpakovanÈ doruËenie)"
    Write-Host "10 - OS  (Odpovedn· sluûba)"
    Write-Host "11 - PR  (Na poötu)"
    Write-Host "12 - PUZ (Podaj u kuriÈra)"
    Write-Host "13 - SV  (Splnomocnenie vyl˙ËenÈ)"
    Write-Host "14 - SVD (Sp‰tnÈ vr·tenie potvrdenej dokument·cie)"
    Write-Host "15 - VR  (Do vlastn˝ch r˙k)"
    Write-Host "16 - VT  (V˝mena tovaru)"
    
    $sluzbyCodes = @{
        1="D"; 2="DOH"; 3="F"; 4="IOD"; 5="NDO"; 6="NEU"; 7="NEV"; 8="NSK";
        9="OD"; 10="OS"; 11="PR"; 12="PUZ"; 13="SV"; 14="SVD"; 15="VR"; 16="VT"
    }
    
    $input = Read-Host "Vyberte sluûby (napr. 1,4,15)"
    $sluzby = @()
    
    if (![string]::IsNullOrWhiteSpace($input)) {
        $cisla = $input -split ','
        foreach ($cislo in $cisla) {
            $c = [int]$cislo.Trim()
            if ($sluzbyCodes.ContainsKey($c)) {
                $sluzby += $sluzbyCodes[$c]
            }
        }
    }
    
    return $sluzby
}

# Hlavn˝ skript
Clear-Host
Write-Host "============================================" -ForegroundColor Green
Write-Host "  ePodacÌ h·rok XML gener·tor" -ForegroundColor Green
Write-Host "  Slovensk· poöta" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green

# Inform·cie o odosielateæovi
Write-Host "`n=== ⁄DAJE ODOSIELATEºA ===" -ForegroundColor Cyan

$odosielatelMeno = Get-UserInput "Meno odosielateæa (tituly, meno, priezvisko)"
$odosielatelOrganizacia = Get-UserInput "Organiz·cia"
$odosielatelUlica = Get-UserInput "Ulica a ËÌslo" -Required $true
$odosielatelMesto = Get-UserInput "Mesto" -Required $true
$odosielatelPSC = Get-UserInput "PS»" -Required $true
$odosielatelKrajina = Get-UserInput "Krajina (ISO kÛd)" -Default "SK"
$odosielatelTelefon = Get-UserInput "TelefÛn (0912345678)"
$odosielatelEmail = Get-UserInput "Email" -Required $true

# Druh z·sielky
$druhZasielky = Get-DruhZasielky

# SpÙsob ˙hrady
$sposobUhrady = Get-SposobUhrady

# PoËet z·sielok
Write-Host "`n=== Z¡SIELKY ===" -ForegroundColor Cyan
[int]$pocetZasielok = Get-UserInput "Koæko z·sielok chcete pridaù?" -Default "1" -Required $true

# Aktu·lny d·tum
$datum = Get-Date -Format "yyyyMMdd"

# Kolekcia z·sielok
$zasielky = @()

for ($i = 1; $i -le $pocetZasielok; $i++) {
    Write-Host "`n--- Z¡SIELKA $i/$pocetZasielok ---" -ForegroundColor Yellow
    
    # Adres·t
    Write-Host "⁄daje adres·ta:" -ForegroundColor White
    $adresatMeno = Get-UserInput "Meno adres·ta"
    $adresatOrganizacia = Get-UserInput "Organiz·cia adres·ta"
    $adresatUlica = Get-UserInput "Ulica a ËÌslo" -Required $true
    $adresatMesto = Get-UserInput "Mesto" -Required $true
    $adresatPSC = Get-UserInput "PS»" -Required $true
    $adresatKrajina = Get-UserInput "Krajina (ISO kÛd)" -Default "SK"
    $adresatTelefon = Get-UserInput "TelefÛn"
    $adresatEmail = Get-UserInput "Email"
    
    # Parametre z·sielky
    Write-Host "`nParametre z·sielky:" -ForegroundColor White
    $hmotnost = Get-UserInput "Hmotnosù (kg, napr. 0.5)" -Default "0.0"
    $poznamka = Get-UserInput "Pozn·mka"
    
    # Dobierka
    $maDobierku = Read-Host "M· z·sielka dobierku? (a/n)"
    $cenaDobierky = ""
    $druhPPP = ""
    if ($maDobierku -eq "a") {
        $cenaDobierky = Get-UserInput "Suma dobierky (EUR)" -Required $true
        Write-Host "SpÙsob v˝platy dobierky:"
        Write-Host "5 - Bezdokladov· dobierka na ˙Ëet"
        Write-Host "6 - Bezdokladov· dobierka na adresu"
        $druhPPP = Get-UserInput "Vyberte" -Default "5"
    }
    
    # Poistenie
    $maPoistenie = Read-Host "M· z·sielka poistenie? (a/n)"
    $cenaPoistneho = ""
    if ($maPoistenie -eq "a") {
        $cenaPoistneho = Get-UserInput "Suma poistenia (EUR)" -Required $true
    }
    
    # Sluûby
    $sluzby = Get-Sluzby
    
    # Vytvorenie objektu z·sielky
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
Write-Host "`n=== GENEROVANIE XML ===" -ForegroundColor Cyan

# XML hlaviËka a namespace
$xml = @"
<?xml version="1.0" encoding="UTF-8"?>
<EPH verzia="3.0" xmlns="http://ekp.posta.sk/LOGIS/Formulare/Podaj_v03">
  <InfoEPH>
    <Mena>EUR</Mena>
    <TypEPH>1</TypEPH>
    <EPHID></EPHID>
    <Datum>$datum</Datum>
    <PocetZasielok>$pocetZasielok</PocetZasielok>
    <Uhrada>
      <SposobUhrady>$sposobUhrady</SposobUhrady>
      <SumaUhrady>0.00</SumaUhrady>
    </Uhrada>
    <DruhZasielky>$druhZasielky</DruhZasielky>
    <Odosielatel>
      <OdosielatelID></OdosielatelID>
      <Meno>$odosielatelMeno</Meno>
      <Organizacia>$odosielatelOrganizacia</Organizacia>
      <Ulica>$odosielatelUlica</Ulica>
      <Mesto>$odosielatelMesto</Mesto>
      <PSC>$odosielatelPSC</PSC>
      <Krajina>$odosielatelKrajina</Krajina>
      <Telefon>$odosielatelTelefon</Telefon>
      <Email>$odosielatelEmail</Email>
    </Odosielatel>
  </InfoEPH>
  <Zasielky>
"@

# Pridanie jednotliv˝ch z·sielok
foreach ($z in $zasielky) {
    $xml += @"

    <Zasielka>
      <Adresat>
        <Meno>$($z.AdresatMeno)</Meno>
        <Organizacia>$($z.AdresatOrganizacia)</Organizacia>
        <Ulica>$($z.AdresatUlica)</Ulica>
        <Mesto>$($z.AdresatMesto)</Mesto>
        <PSC>$($z.AdresatPSC)</PSC>
        <Krajina>$($z.AdresatKrajina)</Krajina>
        <Telefon>$($z.AdresatTelefon)</Telefon>
        <Email>$($z.AdresatEmail)</Email>
      </Adresat>
      <Info>
        <Hmotnost>$($z.Hmotnost)</Hmotnost>
"@

    if ($z.CenaDobierky) {
        $xml += "`n        <CenaDobierky>$($z.CenaDobierky)</CenaDobierky>"
        $xml += "`n        <DruhPPP>$($z.DruhPPP)</DruhPPP>"
    }
    
    if ($z.CenaPoistneho) {
        $xml += "`n        <CenaPoistneho>$($z.CenaPoistneho)</CenaPoistneho>"
    }
    
    if ($z.Poznamka) {
        $xml += "`n        <Poznamka>$($z.Poznamka)</Poznamka>"
    }
    
    $xml += @"

      </Info>
"@

    # Pridanie sluûieb
    if ($z.Sluzby.Count -gt 0) {
        $xml += "      <PouziteSluzby>`n"
        foreach ($sluzba in $z.Sluzby) {
            $xml += "        <Sluzba>$sluzba</Sluzba>`n"
        }
        $xml += "      </PouziteSluzby>`n"
    }
    
    $xml += "    </Zasielka>"
}

# UkonËenie XML
$xml += @"

  </Zasielky>
</EPH>
"@

# Uloûenie s˙boru
$defaultFileName = "epodaci_harok_$datum.xml"
$fileName = Get-UserInput "N·zov v˝stupnÈho s˙boru" -Default $defaultFileName

# Uloûenie s UTF-8 BOM (bez BOM mÙûe spÙsobiù problÈmy)
$utf8WithoutBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($fileName, $xml, $utf8WithoutBom)

Write-Host "`n============================================" -ForegroundColor Green
Write-Host "XML s˙bor bol ˙speöne vytvoren˝!" -ForegroundColor Green
Write-Host "Umiestnenie: $((Get-Item $fileName).FullName)" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "`nS˙bor mÙûete nahraù na www.posta.sk v sekcii ePodacÌ h·rok." -ForegroundColor Cyan
Write-Host ""

# Ponuka na zobrazenie obsahu
$zobraz = Read-Host "Chcete zobraziù obsah XML? (a/n)"
if ($zobraz -eq "a") {
    Write-Host "`n--- OBSAH XML ---" -ForegroundColor Yellow
    Get-Content $fileName
}