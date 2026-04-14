# Script d'analyse des transactions COMPLÈTES
# Génère 3 rapports : Items détaillés + Global avec Sellwand + Ventes par Joueur

param(
    [string]$DossierLogs = ".\Economie",
    [string]$RapportItems = ".\Rapport-Items.html",
    [string]$RapportGlobal = ".\Rapport-Global.html",
    [string]$RapportJoueurs = ".\Rapport-Joueurs.html"
)

# ===== FONCTION 1: Parser les transactions NORMALES =====
function Parse-LigneTransactionNormale {
    param([string]$ligne)
    
    if ($ligne -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] - (\S+) (achet.|vendu) (.+?) pour \$([0-9,]+\.\d{2})') {
        $date = [DateTime]::ParseExact($matches[1], "yyyy-MM-dd HH:mm:ss", $null)
        $joueur = $matches[2]
        $typeAction = $matches[3]
        $type = if ($typeAction -like "achet*") { "Achat" } else { "Vente" }
        $itemsRaw = $matches[4]
        $montant = [decimal]($matches[5] -replace ',', '')
        
        $itemsListe = @()
        $pattern = '(\d+)\s*x\s+([^<,]+?)(?=<|,|\s+pour|$)'
        $itemMatches = [regex]::Matches($itemsRaw, $pattern)
        
        if ($itemMatches.Count -gt 0) {
            foreach ($match in $itemMatches) {
                $quantite = [int]$match.Groups[1].Value
                $nomItem = $match.Groups[2].Value.Trim()
                
                if ($nomItem -and $nomItem.Length -gt 0) {
                    $itemsListe += [PSCustomObject]@{
                        Item = $nomItem
                        Quantite = $quantite
                    }
                }
            }
        }
        
        if ($itemsListe.Count -gt 0) {
            return [PSCustomObject]@{
                Date = $date
                Joueur = $joueur
                Type = $type
                Items = $itemsListe
                Montant = $montant
                Source = "Normal"
            }
        }
    }
    return $null
}

# ===== FONCTION 2: Parser les ventes SELLWAND (ancien format .txt) =====
function Parse-LigneSellwand {
    param([string]$ligne)
    
    if ($ligne -match '(\d{2})\.(\d{2})\.(\d{4}) @ (\d{2}):(\d{2}):(\d{2}) / \[SALE\] User: (\S+) .+ - Money: \$([0-9,]+(?:\.\d+)?) - Items: x([0-9,]+)') {
        $jour = $matches[1]
        $mois = $matches[2]
        $annee = $matches[3]
        $heure = $matches[4]
        $minute = $matches[5]
        $seconde = $matches[6]
        
        $dateStr = "$annee-$mois-$jour $heure`:$minute`:$seconde"
        $date = [DateTime]::ParseExact($dateStr, "yyyy-MM-dd HH:mm:ss", $null)
        
        $joueur = $matches[7]
        $montantStr = $matches[8] -replace ',', ''
        $montant = [decimal]$montantStr
        
        $quantiteStr = $matches[9] -replace ',', ''
        $quantite = [int]$quantiteStr
        
        return [PSCustomObject]@{
            Date = $date
            Joueur = $joueur
            Type = "Vente"
            Items = @()
            Montant = $montant
            Source = "Sellwand"
            QuantiteSellwand = $quantite
        }
    }
    return $null
}

# ===== FONCTION 3: Parser les ventes SELLWAND (nouveau format .log) =====
function Parse-LigneSellwandNew {
    param(
        [string]$ligne,
        [string]$dateFichier
    )
    
    if ($ligne -match '^\[(\d{2}:\d{2}:\d{2})\]\s+(?:\[ShulkerFix\]\s+)?\.?(\S+)\s+sold\s+([0-9,]+)x\s+items\s+\[(.+?)\](?:\s+from shulker boxes)?\s+and\s+earned\s+([0-9,.]+)\s+\(multiplier:\s+[0-9.]+(?:,\s+uses:\s+\d+)?\)') {
        $heure = $matches[1]
        $joueur = $matches[2]
        $quantiteTotaleStr = $matches[3] -replace ',', ''
        $quantiteTotale = [int]$quantiteTotaleStr
        $itemsRaw = $matches[4]
        $montantStr = $matches[5] -replace ',', ''
        $montant = [decimal]$montantStr
        
        $dateStr = "$dateFichier $heure"
        try {
            $date = [DateTime]::ParseExact($dateStr, "yyyy-MM-dd HH:mm:ss", $null)
        } catch {
            return $null
        }
        
        $itemsListe = @()
        $itemPattern = '(\d+)x\s+([A-Z_]+)'
        $itemMatches = [regex]::Matches($itemsRaw, $itemPattern)
        
        foreach ($match in $itemMatches) {
            $quantite = [int]$match.Groups[1].Value
            $nomItem = $match.Groups[2].Value.Trim()
            
            if ($nomItem -and $nomItem.Length -gt 0) {
                $itemsListe += [PSCustomObject]@{
                    Item = $nomItem
                    Quantite = $quantite
                }
            }
        }
        
        return [PSCustomObject]@{
            Date = $date
            Joueur = $joueur
            Type = "Vente"
            Items = $itemsListe
            Montant = $montant
            Source = "SellwandNew"
            QuantiteSellwand = $quantiteTotale
        }
    }
    return $null
}

# ===== FONCTION: Normaliser les noms d'items =====
# Convertit "SHULKER_BOX" et "Shulker Box" vers un format uniforme "Shulker Box"
function Normalize-ItemName {
    param([string]$nom)
    
    # Remplacer les underscores par des espaces
    $nom = $nom -replace '_', ' '
    
    # Convertir en Title Case (premiere lettre majuscule, reste minuscule pour chaque mot)
    $nom = (Get-Culture).TextInfo.ToTitleCase($nom.ToLower())
    
    return $nom
}

function Get-NomMoisFR {
    param([int]$mois)
    $moisFR = @{
        1 = "Janvier"; 2 = "Fevrier"; 3 = "Mars"; 4 = "Avril"
        5 = "Mai"; 6 = "Juin"; 7 = "Juillet"; 8 = "Aout"
        9 = "Septembre"; 10 = "Octobre"; 11 = "Novembre"; 12 = "Decembre"
    }
    return $moisFR[$mois]
}

Write-Host "=== ANALYSE DES TRANSACTIONS ===" -ForegroundColor Cyan
Write-Host ""

# ===== CHARGEMENT DES TRANSACTIONS NORMALES =====
Write-Host "1. Transactions normales..." -ForegroundColor Yellow

$fichiersNormaux = Get-ChildItem -Path $DossierLogs -Filter "transaction-log-*.txt" -ErrorAction SilentlyContinue

$transactionsNormales = @()
foreach ($fichier in $fichiersNormaux) {
    Write-Host "   - $($fichier.Name)" -ForegroundColor Gray
    $lignes = Get-Content $fichier.FullName -Encoding UTF8
    foreach ($ligne in $lignes) {
        $trans = Parse-LigneTransactionNormale -ligne $ligne
        if ($trans) { $transactionsNormales += $trans }
    }
}

Write-Host "   => $($transactionsNormales.Count) transactions" -ForegroundColor Green
Write-Host ""

# ===== CHARGEMENT DES VENTES SELLWAND =====
Write-Host "2. Ventes Sellwand..." -ForegroundColor Yellow

$dossiersSellwand = Get-ChildItem -Path $DossierLogs -Directory -Filter "Sellwand-*" -ErrorAction SilentlyContinue

$ventesSellwand = @()
foreach ($dossier in $dossiersSellwand) {
    Write-Host "   - Dossier: $($dossier.Name)" -ForegroundColor Gray
    
    # Ancien format (.txt) - DD.MM.YYYY.txt
    $fichiersTxt = Get-ChildItem -Path $dossier.FullName -Filter "*.txt" -ErrorAction SilentlyContinue
    foreach ($fichier in $fichiersTxt) {
        $lignes = Get-Content $fichier.FullName -Encoding UTF8
        foreach ($ligne in $lignes) {
            $vente = Parse-LigneSellwand -ligne $ligne
            if ($vente) { $ventesSellwand += $vente }
        }
    }
    
    # Nouveau format (.log) - YYYY-MM-DD.log
    $fichiersLog = Get-ChildItem -Path $dossier.FullName -Filter "*.log" -ErrorAction SilentlyContinue
    foreach ($fichier in $fichiersLog) {
        $dateFichier = [System.IO.Path]::GetFileNameWithoutExtension($fichier.Name)
        if ($dateFichier -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
        $lignes = Get-Content $fichier.FullName -Encoding UTF8
        foreach ($ligne in $lignes) {
            if ($ligne.Trim() -eq '') { continue }
            $vente = Parse-LigneSellwandNew -ligne $ligne -dateFichier $dateFichier
            if ($vente) { $ventesSellwand += $vente }
        }
    }
}

Write-Host "   => $($ventesSellwand.Count) ventes" -ForegroundColor Green
Write-Host ""

$dateGeneration = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'

# ===================================================
# RAPPORT 1 : ITEMS
# ===================================================

Write-Host "Generation Rapport ITEMS..." -ForegroundColor Cyan

$rapportItemsMensuel = @{}

# Inclure les transactions normales ET les ventes sellwand nouveau format (avec items)
$transactionsAvecItems = $transactionsNormales + ($ventesSellwand | Where-Object { $_.Items.Count -gt 0 })

foreach ($trans in $transactionsAvecItems) {
    $mois = $trans.Date.ToString("yyyy-MM")
    
    if (-not $rapportItemsMensuel.ContainsKey($mois)) {
        $rapportItemsMensuel[$mois] = @{
            DebutPeriode = $trans.Date
            FinPeriode = $trans.Date
            Items = @{}
        }
    }
    
    if ($trans.Date -lt $rapportItemsMensuel[$mois].DebutPeriode) {
        $rapportItemsMensuel[$mois].DebutPeriode = $trans.Date
    }
    if ($trans.Date -gt $rapportItemsMensuel[$mois].FinPeriode) {
        $rapportItemsMensuel[$mois].FinPeriode = $trans.Date
    }
    
    foreach ($item in $trans.Items) {
        $itemNom = Normalize-ItemName -nom $item.Item
        
        if (-not $rapportItemsMensuel[$mois].Items.ContainsKey($itemNom)) {
            $rapportItemsMensuel[$mois].Items[$itemNom] = @{
                Vente = 0
                Achat = 0
            }
        }
        
        $montantParItem = $trans.Montant / $trans.Items.Count
        
        if ($trans.Type -eq "Vente") {
            $rapportItemsMensuel[$mois].Items[$itemNom].Vente += $montantParItem
        } else {
            $rapportItemsMensuel[$mois].Items[$itemNom].Achat += $montantParItem
        }
    }
}

$sbItems = New-Object System.Text.StringBuilder

[void]$sbItems.AppendLine('<!DOCTYPE html>')
[void]$sbItems.AppendLine('<html lang="fr">')
[void]$sbItems.AppendLine('<head>')
[void]$sbItems.AppendLine('    <meta charset="UTF-8">')
[void]$sbItems.AppendLine('    <title>Rapport ITEMS</title>')
[void]$sbItems.AppendLine('    <style>')
[void]$sbItems.AppendLine('        * { margin: 0; padding: 0; box-sizing: border-box; }')
[void]$sbItems.AppendLine('        body { font-family: "Segoe UI", sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; }')
[void]$sbItems.AppendLine('        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }')
[void]$sbItems.AppendLine('        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; }')
[void]$sbItems.AppendLine('        .header h1 { font-size: 2.5em; margin-bottom: 10px; }')
[void]$sbItems.AppendLine('        .content { padding: 40px; }')
[void]$sbItems.AppendLine('        .section-title { font-size: 2em; color: #667eea; margin-bottom: 20px; border-bottom: 3px solid #667eea; padding-bottom: 10px; }')
[void]$sbItems.AppendLine('        .accordion { background: white; border: 2px solid #e0e0e0; border-radius: 10px; margin-bottom: 15px; overflow: hidden; }')
[void]$sbItems.AppendLine('        .accordion-header { display: flex; justify-content: space-between; align-items: center; padding: 20px; background: linear-gradient(135deg, #f8f9ff 0%, #e0e7ff 100%); cursor: pointer; transition: background 0.3s; }')
[void]$sbItems.AppendLine('        .accordion-header:hover { background: linear-gradient(135deg, #e0e7ff 0%, #d0d9ff 100%); }')
[void]$sbItems.AppendLine('        .accordion-title { font-size: 1.5em; color: #764ba2; font-weight: 700; }')
[void]$sbItems.AppendLine('        .accordion-stats { display: flex; gap: 20px; font-size: 0.9em; }')
[void]$sbItems.AppendLine('        .stat-badge { padding: 8px 15px; border-radius: 20px; font-weight: 600; }')
[void]$sbItems.AppendLine('        .badge-vente { background: #d1fae5; color: #065f46; }')
[void]$sbItems.AppendLine('        .badge-achat { background: #fee2e2; color: #991b1b; }')
[void]$sbItems.AppendLine('        .toggle-icon { font-size: 1.5em; color: #667eea; transition: transform 0.3s; }')
[void]$sbItems.AppendLine('        .toggle-icon.active { transform: rotate(180deg); }')
[void]$sbItems.AppendLine('        .accordion-content { max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; }')
[void]$sbItems.AppendLine('        .accordion-content.active { max-height: 5000px; }')
[void]$sbItems.AppendLine('        .accordion-content-inner { padding: 20px; }')
[void]$sbItems.AppendLine('        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; box-shadow: 0 2px 15px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }')
[void]$sbItems.AppendLine('        thead { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }')
[void]$sbItems.AppendLine('        th { padding: 15px; text-align: left; font-weight: 600; }')
[void]$sbItems.AppendLine('        td { padding: 12px 15px; border-bottom: 1px solid #e0e0e0; }')
[void]$sbItems.AppendLine('        tbody tr:hover { background-color: #f8f9ff; }')
[void]$sbItems.AppendLine('        .montant-vente { color: #10b981; font-weight: 600; }')
[void]$sbItems.AppendLine('        .montant-achat { color: #ef4444; font-weight: 600; }')
[void]$sbItems.AppendLine('        .total-row { background-color: #f8f9ff !important; font-weight: 700; border-top: 2px solid #667eea; }')
[void]$sbItems.AppendLine('        .periode { display: inline-block; background: #667eea; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; margin-left: 10px; }')
[void]$sbItems.AppendLine('        .grid-2cols { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }')
[void]$sbItems.AppendLine('        .expand-all-btn { background: #667eea; color: white; border: none; padding: 12px 25px; border-radius: 8px; cursor: pointer; font-size: 1em; margin-bottom: 20px; }')
[void]$sbItems.AppendLine('        .expand-all-btn:hover { background: #5568d3; }')
[void]$sbItems.AppendLine('    </style>')
[void]$sbItems.AppendLine('</head>')
[void]$sbItems.AppendLine('<body>')
[void]$sbItems.AppendLine('    <div class="container">')
[void]$sbItems.AppendLine('        <div class="header">')
[void]$sbItems.AppendLine('            <h1>Rapport ITEMS</h1>')
[void]$sbItems.AppendLine("            <p>Genere le $dateGeneration</p>")
[void]$sbItems.AppendLine('        </div>')
[void]$sbItems.AppendLine('        <div class="content">')
[void]$sbItems.AppendLine('            <h2 class="section-title">Top Items par Mois</h2>')
[void]$sbItems.AppendLine('            <button class="expand-all-btn" onclick="toggleAllAccordions()">Tout Deplier</button>')

$moisTries = $rapportItemsMensuel.Keys | Sort-Object
$moisIndex = 0

foreach ($mois in $moisTries) {
    $debut = $rapportItemsMensuel[$mois].DebutPeriode.ToString("dd/MM/yyyy")
    $fin = $rapportItemsMensuel[$mois].FinPeriode.ToString("dd/MM/yyyy")
    $moisNum = [int]$mois.Substring(5)
    $moisNom = Get-NomMoisFR -mois $moisNum
    $annee = $mois.Substring(0, 4)
    
    $totalVenteMois = 0
    $totalAchatMois = 0
    foreach ($item in $rapportItemsMensuel[$mois].Items.Keys) {
        $totalVenteMois += $rapportItemsMensuel[$mois].Items[$item].Vente
        $totalAchatMois += $rapportItemsMensuel[$mois].Items[$item].Achat
    }
    
    $uniqueId = "mois_$moisIndex"
    $moisIndex++
    
    [void]$sbItems.AppendLine("            <div class='accordion'>")
    [void]$sbItems.AppendLine("                <div class='accordion-header' onclick='toggleAccordion(`"$uniqueId`")'>")
    [void]$sbItems.AppendLine("                    <div class='accordion-title'>$moisNom $annee <span class='periode'>$debut au $fin</span></div>")
    [void]$sbItems.AppendLine("                    <div style='display: flex; align-items: center; gap: 15px;'>")
    [void]$sbItems.AppendLine("                        <div class='accordion-stats'>")
    [void]$sbItems.AppendLine("                            <span class='stat-badge badge-vente'>Ventes: $" + $totalVenteMois.ToString('N2') + "</span>")
    [void]$sbItems.AppendLine("                            <span class='stat-badge badge-achat'>Achats: $" + $totalAchatMois.ToString('N2') + "</span>")
    [void]$sbItems.AppendLine("                        </div>")
    [void]$sbItems.AppendLine("                        <span class='toggle-icon' id='icon_$uniqueId'>&#9660;</span>")
    [void]$sbItems.AppendLine("                    </div>")
    [void]$sbItems.AppendLine("                </div>")
    [void]$sbItems.AppendLine("                <div class='accordion-content' id='$uniqueId'>")
    [void]$sbItems.AppendLine("                    <div class='accordion-content-inner'>")
    
    # Ventes
    $itemsVente = @()
    foreach ($item in $rapportItemsMensuel[$mois].Items.Keys) {
        $vente = $rapportItemsMensuel[$mois].Items[$item].Vente
        if ($vente -gt 0) {
            $itemsVente += [PSCustomObject]@{ Item = $item; Montant = $vente }
        }
    }
    $itemsVente = $itemsVente | Sort-Object -Property Montant -Descending
    
    # Achats
    $itemsAchat = @()
    foreach ($item in $rapportItemsMensuel[$mois].Items.Keys) {
        $achat = $rapportItemsMensuel[$mois].Items[$item].Achat
        if ($achat -gt 0) {
            $itemsAchat += [PSCustomObject]@{ Item = $item; Montant = $achat }
        }
    }
    $itemsAchat = $itemsAchat | Sort-Object -Property Montant -Descending
    
    [void]$sbItems.AppendLine("                        <div class='grid-2cols'>")
    [void]$sbItems.AppendLine("                            <div>")
    [void]$sbItems.AppendLine("                                <h4 style='color: #10b981;'>Top Ventes</h4>")
    [void]$sbItems.AppendLine("                                <table><thead><tr><th>Rang</th><th>Item</th><th>Montant</th></tr></thead><tbody>")
    
    $rang = 1
    foreach ($item in $itemsVente) {
        [void]$sbItems.AppendLine("                                    <tr><td><strong>#$rang</strong></td><td>$($item.Item)</td><td class='montant-vente'>$" + $item.Montant.ToString('N2') + "</td></tr>")
        $rang++
    }
    
    if ($itemsVente.Count -gt 0) {
        [void]$sbItems.AppendLine("                                    <tr class='total-row'><td colspan='2'><strong>TOTAL</strong></td><td class='montant-vente'>$" + $totalVenteMois.ToString('N2') + "</td></tr>")
    }
    
    [void]$sbItems.AppendLine("                                </tbody></table>")
    [void]$sbItems.AppendLine("                            </div>")
    
    [void]$sbItems.AppendLine("                            <div>")
    [void]$sbItems.AppendLine("                                <h4 style='color: #ef4444;'>Top Achats</h4>")
    [void]$sbItems.AppendLine("                                <table><thead><tr><th>Rang</th><th>Item</th><th>Montant</th></tr></thead><tbody>")
    
    $rang = 1
    foreach ($item in $itemsAchat) {
        [void]$sbItems.AppendLine("                                    <tr><td><strong>#$rang</strong></td><td>$($item.Item)</td><td class='montant-achat'>$" + $item.Montant.ToString('N2') + "</td></tr>")
        $rang++
    }
    
    if ($itemsAchat.Count -gt 0) {
        [void]$sbItems.AppendLine("                                    <tr class='total-row'><td colspan='2'><strong>TOTAL</strong></td><td class='montant-achat'>$" + $totalAchatMois.ToString('N2') + "</td></tr>")
    }
    
    [void]$sbItems.AppendLine("                                </tbody></table>")
    [void]$sbItems.AppendLine("                            </div>")
    [void]$sbItems.AppendLine("                        </div>")
    [void]$sbItems.AppendLine("                    </div>")
    [void]$sbItems.AppendLine("                </div>")
    [void]$sbItems.AppendLine("            </div>")
}

[void]$sbItems.AppendLine('        </div>')
[void]$sbItems.AppendLine('    </div>')
[void]$sbItems.AppendLine('    <script>')
[void]$sbItems.AppendLine('        function toggleAccordion(id) {')
[void]$sbItems.AppendLine('            var content = document.getElementById(id);')
[void]$sbItems.AppendLine('            var icon = document.getElementById("icon_" + id);')
[void]$sbItems.AppendLine('            content.classList.toggle("active");')
[void]$sbItems.AppendLine('            icon.classList.toggle("active");')
[void]$sbItems.AppendLine('        }')
[void]$sbItems.AppendLine('        function toggleAllAccordions() {')
[void]$sbItems.AppendLine('            var contents = document.querySelectorAll(".accordion-content");')
[void]$sbItems.AppendLine('            var icons = document.querySelectorAll(".toggle-icon");')
[void]$sbItems.AppendLine('            var btn = document.querySelector(".expand-all-btn");')
[void]$sbItems.AppendLine('            var isExpanding = btn.textContent.includes("Deplier");')
[void]$sbItems.AppendLine('            contents.forEach(function(el) { if (isExpanding) el.classList.add("active"); else el.classList.remove("active"); });')
[void]$sbItems.AppendLine('            icons.forEach(function(el) { if (isExpanding) el.classList.add("active"); else el.classList.remove("active"); });')
[void]$sbItems.AppendLine('            btn.textContent = isExpanding ? "Tout Replier" : "Tout Deplier";')
[void]$sbItems.AppendLine('        }')
[void]$sbItems.AppendLine('    </script>')
[void]$sbItems.AppendLine('</body>')
[void]$sbItems.AppendLine('</html>')

$sbItems.ToString() | Out-File -FilePath $RapportItems -Encoding UTF8
Write-Host "  => $RapportItems" -ForegroundColor Green

# ===================================================
# RAPPORT 2 : GLOBAL
# ===================================================

Write-Host "Generation Rapport GLOBAL..." -ForegroundColor Cyan

$toutesTransactions = $transactionsNormales + $ventesSellwand
$rapportGlobalMensuel = @{}

foreach ($trans in $toutesTransactions) {
    $mois = $trans.Date.ToString("yyyy-MM")
    
    if (-not $rapportGlobalMensuel.ContainsKey($mois)) {
        $rapportGlobalMensuel[$mois] = @{
            DebutPeriode = $trans.Date
            FinPeriode = $trans.Date
            TotalVente = 0
            TotalAchat = 0
            Joueurs = @{}
        }
    }
    
    if ($trans.Date -lt $rapportGlobalMensuel[$mois].DebutPeriode) {
        $rapportGlobalMensuel[$mois].DebutPeriode = $trans.Date
    }
    if ($trans.Date -gt $rapportGlobalMensuel[$mois].FinPeriode) {
        $rapportGlobalMensuel[$mois].FinPeriode = $trans.Date
    }
    
    if ($trans.Type -eq "Vente") {
        $rapportGlobalMensuel[$mois].TotalVente += $trans.Montant
    } else {
        $rapportGlobalMensuel[$mois].TotalAchat += $trans.Montant
    }
    
    if (-not $rapportGlobalMensuel[$mois].Joueurs.ContainsKey($trans.Joueur)) {
        $rapportGlobalMensuel[$mois].Joueurs[$trans.Joueur] = @{ Vente = 0; Achat = 0 }
    }
    
    if ($trans.Type -eq "Vente") {
        $rapportGlobalMensuel[$mois].Joueurs[$trans.Joueur].Vente += $trans.Montant
    } else {
        $rapportGlobalMensuel[$mois].Joueurs[$trans.Joueur].Achat += $trans.Montant
    }
}

$sbGlobal = New-Object System.Text.StringBuilder

[void]$sbGlobal.AppendLine('<!DOCTYPE html>')
[void]$sbGlobal.AppendLine('<html lang="fr">')
[void]$sbGlobal.AppendLine('<head>')
[void]$sbGlobal.AppendLine('    <meta charset="UTF-8">')
[void]$sbGlobal.AppendLine('    <title>Rapport GLOBAL</title>')
[void]$sbGlobal.AppendLine('    <style>')
[void]$sbGlobal.AppendLine('        * { margin: 0; padding: 0; box-sizing: border-box; }')
[void]$sbGlobal.AppendLine('        body { font-family: "Segoe UI", sans-serif; background: linear-gradient(135deg, #10b981 0%, #059669 100%); padding: 20px; }')
[void]$sbGlobal.AppendLine('        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }')
[void]$sbGlobal.AppendLine('        .header { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; padding: 40px; text-align: center; }')
[void]$sbGlobal.AppendLine('        .header h1 { font-size: 2.5em; }')
[void]$sbGlobal.AppendLine('        .content { padding: 40px; }')
[void]$sbGlobal.AppendLine('        .section-title { font-size: 2em; color: #10b981; margin-bottom: 20px; border-bottom: 3px solid #10b981; padding-bottom: 10px; }')
[void]$sbGlobal.AppendLine('        .accordion { background: white; border: 2px solid #e0e0e0; border-radius: 10px; margin-bottom: 15px; overflow: hidden; }')
[void]$sbGlobal.AppendLine('        .accordion-header { display: flex; justify-content: space-between; align-items: center; padding: 20px; background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); cursor: pointer; }')
[void]$sbGlobal.AppendLine('        .accordion-header:hover { background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%); }')
[void]$sbGlobal.AppendLine('        .accordion-title { font-size: 1.5em; color: #059669; font-weight: 700; }')
[void]$sbGlobal.AppendLine('        .stat-badge { padding: 8px 15px; border-radius: 20px; font-weight: 600; }')
[void]$sbGlobal.AppendLine('        .badge-vente { background: #d1fae5; color: #065f46; }')
[void]$sbGlobal.AppendLine('        .badge-achat { background: #fee2e2; color: #991b1b; }')
[void]$sbGlobal.AppendLine('        .toggle-icon { font-size: 1.5em; color: #10b981; transition: transform 0.3s; }')
[void]$sbGlobal.AppendLine('        .toggle-icon.active { transform: rotate(180deg); }')
[void]$sbGlobal.AppendLine('        .accordion-content { max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; }')
[void]$sbGlobal.AppendLine('        .accordion-content.active { max-height: 10000px; }')
[void]$sbGlobal.AppendLine('        .accordion-content-inner { padding: 20px; }')
[void]$sbGlobal.AppendLine('        table { width: 100%; border-collapse: collapse; box-shadow: 0 2px 15px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }')
[void]$sbGlobal.AppendLine('        thead { background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: white; }')
[void]$sbGlobal.AppendLine('        th { padding: 15px; text-align: left; }')
[void]$sbGlobal.AppendLine('        td { padding: 12px 15px; border-bottom: 1px solid #e0e0e0; }')
[void]$sbGlobal.AppendLine('        tbody tr:hover { background-color: #f0fdf4; }')
[void]$sbGlobal.AppendLine('        .montant-vente { color: #10b981; font-weight: 600; }')
[void]$sbGlobal.AppendLine('        .montant-achat { color: #ef4444; font-weight: 600; }')
[void]$sbGlobal.AppendLine('        .total-row { background-color: #f0fdf4 !important; font-weight: 700; border-top: 2px solid #10b981; }')
[void]$sbGlobal.AppendLine('        .periode { display: inline-block; background: #10b981; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; margin-left: 10px; }')
[void]$sbGlobal.AppendLine('        .expand-all-btn { background: #10b981; color: white; border: none; padding: 12px 25px; border-radius: 8px; cursor: pointer; font-size: 1em; margin-bottom: 20px; }')
[void]$sbGlobal.AppendLine('        .expand-all-btn:hover { background: #059669; }')
[void]$sbGlobal.AppendLine('    </style>')
[void]$sbGlobal.AppendLine('</head>')
[void]$sbGlobal.AppendLine('<body>')
[void]$sbGlobal.AppendLine('    <div class="container">')
[void]$sbGlobal.AppendLine('        <div class="header">')
[void]$sbGlobal.AppendLine('            <h1>Rapport GLOBAL</h1>')
[void]$sbGlobal.AppendLine("            <p>Genere le $dateGeneration</p>")
[void]$sbGlobal.AppendLine('        </div>')
[void]$sbGlobal.AppendLine('        <div class="content">')
[void]$sbGlobal.AppendLine('            <h2 class="section-title">Top Joueurs par Mois</h2>')
[void]$sbGlobal.AppendLine('            <button class="expand-all-btn" onclick="toggleAllAccordions()">Tout Deplier</button>')

$moisTries = $rapportGlobalMensuel.Keys | Sort-Object
$moisIndex = 0

foreach ($mois in $moisTries) {
    $debut = $rapportGlobalMensuel[$mois].DebutPeriode.ToString("dd/MM/yyyy")
    $fin = $rapportGlobalMensuel[$mois].FinPeriode.ToString("dd/MM/yyyy")
    $moisNum = [int]$mois.Substring(5)
    $moisNom = Get-NomMoisFR -mois $moisNum
    $annee = $mois.Substring(0, 4)
    
    $totalVenteMois = $rapportGlobalMensuel[$mois].TotalVente
    $totalAchatMois = $rapportGlobalMensuel[$mois].TotalAchat
    
    $uniqueId = "global_$moisIndex"
    $moisIndex++
    
    [void]$sbGlobal.AppendLine("            <div class='accordion'>")
    [void]$sbGlobal.AppendLine("                <div class='accordion-header' onclick='toggleAccordion(`"$uniqueId`")'>")
    [void]$sbGlobal.AppendLine("                    <div class='accordion-title'>$moisNom $annee <span class='periode'>$debut au $fin</span></div>")
    [void]$sbGlobal.AppendLine("                    <div style='display: flex; align-items: center; gap: 15px;'>")
    [void]$sbGlobal.AppendLine("                        <div style='display: flex; gap: 20px;'>")
    [void]$sbGlobal.AppendLine("                            <span class='stat-badge badge-vente'>Ventes: $" + $totalVenteMois.ToString('N2') + "</span>")
    [void]$sbGlobal.AppendLine("                            <span class='stat-badge badge-achat'>Achats: $" + $totalAchatMois.ToString('N2') + "</span>")
    [void]$sbGlobal.AppendLine("                        </div>")
    [void]$sbGlobal.AppendLine("                        <span class='toggle-icon' id='icon_$uniqueId'>&#9660;</span>")
    [void]$sbGlobal.AppendLine("                    </div>")
    [void]$sbGlobal.AppendLine("                </div>")
    [void]$sbGlobal.AppendLine("                <div class='accordion-content' id='$uniqueId'>")
    [void]$sbGlobal.AppendLine("                    <div class='accordion-content-inner'>")
    
    $top20 = $rapportGlobalMensuel[$mois].Joueurs.GetEnumerator() | 
             Sort-Object { $_.Value.Vente } -Descending | 
             Select-Object -First 20
    
    [void]$sbGlobal.AppendLine("                        <table>")
    [void]$sbGlobal.AppendLine("                            <thead><tr><th>Rang</th><th>Joueur</th><th>Ventes</th><th>Achats</th><th>Balance</th></tr></thead>")
    [void]$sbGlobal.AppendLine("                            <tbody>")
    
    $rang = 1
    foreach ($j in $top20) {
        $vente = $j.Value.Vente
        $achat = $j.Value.Achat
        $balance = $vente - $achat
        $balanceColor = if ($balance -ge 0) { "montant-vente" } else { "montant-achat" }
        
        [void]$sbGlobal.AppendLine("                                <tr>")
        [void]$sbGlobal.AppendLine("                                    <td><strong>#$rang</strong></td>")
        [void]$sbGlobal.AppendLine("                                    <td>$($j.Key)</td>")
        [void]$sbGlobal.AppendLine("                                    <td class='montant-vente'>$" + $vente.ToString('N2') + "</td>")
        [void]$sbGlobal.AppendLine("                                    <td class='montant-achat'>$" + $achat.ToString('N2') + "</td>")
        [void]$sbGlobal.AppendLine("                                    <td class='$balanceColor'>$" + $balance.ToString('N2') + "</td>")
        [void]$sbGlobal.AppendLine("                                </tr>")
        $rang++
    }
    
    $totalBalance = $totalVenteMois - $totalAchatMois
    
    [void]$sbGlobal.AppendLine("                                <tr class='total-row'>")
    [void]$sbGlobal.AppendLine("                                    <td colspan='2'><strong>TOTAL</strong></td>")
    [void]$sbGlobal.AppendLine("                                    <td class='montant-vente'>$" + $totalVenteMois.ToString('N2') + "</td>")
    [void]$sbGlobal.AppendLine("                                    <td class='montant-achat'>$" + $totalAchatMois.ToString('N2') + "</td>")
    [void]$sbGlobal.AppendLine("                                    <td class='montant-vente'>$" + $totalBalance.ToString('N2') + "</td>")
    [void]$sbGlobal.AppendLine("                                </tr>")
    [void]$sbGlobal.AppendLine("                            </tbody>")
    [void]$sbGlobal.AppendLine("                        </table>")
    [void]$sbGlobal.AppendLine("                    </div>")
    [void]$sbGlobal.AppendLine("                </div>")
    [void]$sbGlobal.AppendLine("            </div>")
}

[void]$sbGlobal.AppendLine('        </div>')
[void]$sbGlobal.AppendLine('    </div>')
[void]$sbGlobal.AppendLine('    <script>')
[void]$sbGlobal.AppendLine('        function toggleAccordion(id) {')
[void]$sbGlobal.AppendLine('            document.getElementById(id).classList.toggle("active");')
[void]$sbGlobal.AppendLine('            document.getElementById("icon_" + id).classList.toggle("active");')
[void]$sbGlobal.AppendLine('        }')
[void]$sbGlobal.AppendLine('        function toggleAllAccordions() {')
[void]$sbGlobal.AppendLine('            var btn = document.querySelector(".expand-all-btn");')
[void]$sbGlobal.AppendLine('            var isExpanding = btn.textContent.includes("Deplier");')
[void]$sbGlobal.AppendLine('            document.querySelectorAll(".accordion-content").forEach(el => isExpanding ? el.classList.add("active") : el.classList.remove("active"));')
[void]$sbGlobal.AppendLine('            document.querySelectorAll(".toggle-icon").forEach(el => isExpanding ? el.classList.add("active") : el.classList.remove("active"));')
[void]$sbGlobal.AppendLine('            btn.textContent = isExpanding ? "Tout Replier" : "Tout Deplier";')
[void]$sbGlobal.AppendLine('        }')
[void]$sbGlobal.AppendLine('    </script>')
[void]$sbGlobal.AppendLine('</body>')
[void]$sbGlobal.AppendLine('</html>')

$sbGlobal.ToString() | Out-File -FilePath $RapportGlobal -Encoding UTF8
Write-Host "  => $RapportGlobal" -ForegroundColor Green

# ===================================================
# RAPPORT 3 : JOUEURS
# ===================================================

Write-Host "Generation Rapport JOUEURS..." -ForegroundColor Cyan

$rapportJoueursMensuel = @{}

# Inclure les transactions normales ET les ventes sellwand nouveau format (avec items)
$transactionsJoueurs = $transactionsNormales + ($ventesSellwand | Where-Object { $_.Items.Count -gt 0 })

foreach ($trans in $transactionsJoueurs) {
    if ($trans.Type -ne "Vente") { continue }
    
    $mois = $trans.Date.ToString("yyyy-MM")
    
    if (-not $rapportJoueursMensuel.ContainsKey($mois)) {
        $rapportJoueursMensuel[$mois] = @{
            DebutPeriode = $trans.Date
            FinPeriode = $trans.Date
            Joueurs = @{}
        }
    }
    
    if ($trans.Date -lt $rapportJoueursMensuel[$mois].DebutPeriode) {
        $rapportJoueursMensuel[$mois].DebutPeriode = $trans.Date
    }
    if ($trans.Date -gt $rapportJoueursMensuel[$mois].FinPeriode) {
        $rapportJoueursMensuel[$mois].FinPeriode = $trans.Date
    }
    
    if (-not $rapportJoueursMensuel[$mois].Joueurs.ContainsKey($trans.Joueur)) {
        $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur] = @{
            Items = @{}
            TotalVente = 0
        }
    }
    
    $montantParItem = $trans.Montant / $trans.Items.Count
    
    foreach ($item in $trans.Items) {
        $itemNom = Normalize-ItemName -nom $item.Item
        
        if (-not $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur].Items.ContainsKey($itemNom)) {
            $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur].Items[$itemNom] = @{
                Quantite = 0
                Montant = 0
            }
        }
        
        $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur].Items[$itemNom].Quantite += $item.Quantite
        $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur].Items[$itemNom].Montant += $montantParItem
    }
    
    $rapportJoueursMensuel[$mois].Joueurs[$trans.Joueur].TotalVente += $trans.Montant
}

$sbJoueurs = New-Object System.Text.StringBuilder

[void]$sbJoueurs.AppendLine('<!DOCTYPE html>')
[void]$sbJoueurs.AppendLine('<html lang="fr">')
[void]$sbJoueurs.AppendLine('<head>')
[void]$sbJoueurs.AppendLine('    <meta charset="UTF-8">')
[void]$sbJoueurs.AppendLine('    <title>Rapport VENTES PAR JOUEUR</title>')
[void]$sbJoueurs.AppendLine('    <style>')
[void]$sbJoueurs.AppendLine('        * { margin: 0; padding: 0; box-sizing: border-box; }')
[void]$sbJoueurs.AppendLine('        body { font-family: "Segoe UI", sans-serif; background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); padding: 20px; }')
[void]$sbJoueurs.AppendLine('        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 15px; }')
[void]$sbJoueurs.AppendLine('        .header { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 40px; text-align: center; }')
[void]$sbJoueurs.AppendLine('        .header h1 { font-size: 2.5em; }')
[void]$sbJoueurs.AppendLine('        .content { padding: 40px; }')
[void]$sbJoueurs.AppendLine('        .section-title { font-size: 2em; color: #f59e0b; margin-bottom: 20px; border-bottom: 3px solid #f59e0b; padding-bottom: 10px; }')
[void]$sbJoueurs.AppendLine('        .accordion { background: white; border: 2px solid #e0e0e0; border-radius: 10px; margin-bottom: 15px; overflow: hidden; }')
[void]$sbJoueurs.AppendLine('        .accordion-header { display: flex; justify-content: space-between; padding: 20px; background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); cursor: pointer; }')
[void]$sbJoueurs.AppendLine('        .accordion-header:hover { background: linear-gradient(135deg, #fde68a 0%, #fcd34d 100%); }')
[void]$sbJoueurs.AppendLine('        .accordion-title { font-size: 1.5em; color: #d97706; font-weight: 700; }')
[void]$sbJoueurs.AppendLine('        .toggle-icon { font-size: 1.5em; color: #f59e0b; transition: transform 0.3s; }')
[void]$sbJoueurs.AppendLine('        .toggle-icon.active { transform: rotate(180deg); }')
[void]$sbJoueurs.AppendLine('        .accordion-content { max-height: 0; overflow: hidden; transition: max-height 0.3s; }')
[void]$sbJoueurs.AppendLine('        .accordion-content.active { max-height: 10000px; }')
[void]$sbJoueurs.AppendLine('        .accordion-content-inner { padding: 20px; }')
[void]$sbJoueurs.AppendLine('        .joueur-accordion { background: white; border: 1px solid #d1d5db; border-radius: 8px; margin-bottom: 10px; }')
[void]$sbJoueurs.AppendLine('        .joueur-header { display: flex; justify-content: space-between; padding: 15px; background: #f9fafb; cursor: pointer; }')
[void]$sbJoueurs.AppendLine('        .joueur-header:hover { background: #f3f4f6; }')
[void]$sbJoueurs.AppendLine('        .joueur-name { font-size: 1.2em; color: #1f2937; font-weight: 600; }')
[void]$sbJoueurs.AppendLine('        .joueur-total { color: #10b981; font-weight: 700; }')
[void]$sbJoueurs.AppendLine('        .joueur-content { max-height: 0; overflow: hidden; transition: max-height 0.3s; }')
[void]$sbJoueurs.AppendLine('        .joueur-content.active { max-height: 5000px; }')
[void]$sbJoueurs.AppendLine('        .joueur-content-inner { padding: 15px; background: #fafafa; }')
[void]$sbJoueurs.AppendLine('        table { width: 100%; border-collapse: collapse; box-shadow: 0 2px 15px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }')
[void]$sbJoueurs.AppendLine('        thead { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; }')
[void]$sbJoueurs.AppendLine('        th { padding: 15px; text-align: left; }')
[void]$sbJoueurs.AppendLine('        td { padding: 12px 15px; border-bottom: 1px solid #e0e0e0; }')
[void]$sbJoueurs.AppendLine('        tbody tr:hover { background-color: #fef3c7; }')
[void]$sbJoueurs.AppendLine('        .montant-vente { color: #10b981; font-weight: 600; }')
[void]$sbJoueurs.AppendLine('        .total-row { background-color: #fef3c7 !important; font-weight: 700; border-top: 2px solid #f59e0b; }')
[void]$sbJoueurs.AppendLine('        .periode { display: inline-block; background: #f59e0b; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; margin-left: 10px; }')
[void]$sbJoueurs.AppendLine('        .expand-all-btn { background: #f59e0b; color: white; border: none; padding: 12px 25px; border-radius: 8px; cursor: pointer; font-size: 1em; margin-bottom: 20px; }')
[void]$sbJoueurs.AppendLine('        .expand-all-btn:hover { background: #d97706; }')
[void]$sbJoueurs.AppendLine('        .mini-toggle { font-size: 1.2em; color: #f59e0b; transition: transform 0.3s; }')
[void]$sbJoueurs.AppendLine('        .mini-toggle.active { transform: rotate(180deg); }')
[void]$sbJoueurs.AppendLine('    </style>')
[void]$sbJoueurs.AppendLine('</head>')
[void]$sbJoueurs.AppendLine('<body>')
[void]$sbJoueurs.AppendLine('    <div class="container">')
[void]$sbJoueurs.AppendLine('        <div class="header">')
[void]$sbJoueurs.AppendLine('            <h1>Rapport VENTES PAR JOUEUR</h1>')
[void]$sbJoueurs.AppendLine("            <p>Genere le $dateGeneration</p>")
[void]$sbJoueurs.AppendLine('        </div>')
[void]$sbJoueurs.AppendLine('        <div class="content">')
[void]$sbJoueurs.AppendLine('            <h2 class="section-title">Ventes par Joueur</h2>')
[void]$sbJoueurs.AppendLine('            <button class="expand-all-btn" onclick="toggleAllMain()">Tout Deplier</button>')

$moisTries = $rapportJoueursMensuel.Keys | Sort-Object
$moisIndex = 0

foreach ($mois in $moisTries) {
    $debut = $rapportJoueursMensuel[$mois].DebutPeriode.ToString("dd/MM/yyyy")
    $fin = $rapportJoueursMensuel[$mois].FinPeriode.ToString("dd/MM/yyyy")
    $moisNum = [int]$mois.Substring(5)
    $moisNom = Get-NomMoisFR -mois $moisNum
    $annee = $mois.Substring(0, 4)
    
    $totalMois = 0
    foreach ($joueurEntry in $rapportJoueursMensuel[$mois].Joueurs.GetEnumerator()) {
        $totalMois += $joueurEntry.Value.TotalVente
    }
    
    $uniqueId = "joueurs_$moisIndex"
    $moisIndex++
    
    [void]$sbJoueurs.AppendLine("            <div class='accordion'>")
    [void]$sbJoueurs.AppendLine("                <div class='accordion-header' onclick='toggleAccordion(`"$uniqueId`")'>")
    [void]$sbJoueurs.AppendLine("                    <div class='accordion-title'>$moisNom $annee <span class='periode'>$debut au $fin</span></div>")
    [void]$sbJoueurs.AppendLine("                    <div style='display: flex; align-items: center; gap: 15px;'>")
    [void]$sbJoueurs.AppendLine("                        <span style='background: #dcfce7; color: #065f46; padding: 8px 15px; border-radius: 20px; font-weight: 600;'>Total: $" + $totalMois.ToString('N2') + "</span>")
    [void]$sbJoueurs.AppendLine("                        <span class='toggle-icon' id='icon_$uniqueId'>&#9660;</span>")
    [void]$sbJoueurs.AppendLine("                    </div>")
    [void]$sbJoueurs.AppendLine("                </div>")
    [void]$sbJoueurs.AppendLine("                <div class='accordion-content' id='$uniqueId'>")
    [void]$sbJoueurs.AppendLine("                    <div class='accordion-content-inner'>")
    
    $joueursTries = $rapportJoueursMensuel[$mois].Joueurs.GetEnumerator() | 
                    Sort-Object { $_.Value.TotalVente } -Descending
    
    $joueurIndex = 0
    foreach ($joueurEntry in $joueursTries) {
        $joueur = $joueurEntry.Key
        $totalVente = $joueurEntry.Value.TotalVente
        $items = $joueurEntry.Value.Items
        
        $joueurUniqueId = "$uniqueId`_j$joueurIndex"
        $joueurIndex++
        
        [void]$sbJoueurs.AppendLine("                        <div class='joueur-accordion'>")
        [void]$sbJoueurs.AppendLine("                            <div class='joueur-header' onclick='toggleJoueur(`"$joueurUniqueId`")'>")
        [void]$sbJoueurs.AppendLine("                                <span class='joueur-name'>$joueur</span>")
        [void]$sbJoueurs.AppendLine("                                <div style='display: flex; align-items: center; gap: 10px;'>")
        [void]$sbJoueurs.AppendLine("                                    <span class='joueur-total'>$" + $totalVente.ToString('N2') + "</span>")
        [void]$sbJoueurs.AppendLine("                                    <span class='mini-toggle' id='icon_$joueurUniqueId'>&#9660;</span>")
        [void]$sbJoueurs.AppendLine("                                </div>")
        [void]$sbJoueurs.AppendLine("                            </div>")
        [void]$sbJoueurs.AppendLine("                            <div class='joueur-content' id='$joueurUniqueId'>")
        [void]$sbJoueurs.AppendLine("                                <div class='joueur-content-inner'>")
        [void]$sbJoueurs.AppendLine("                                    <table>")
        [void]$sbJoueurs.AppendLine("                                        <thead><tr><th>Item</th><th>Quantite</th><th>Montant</th></tr></thead>")
        [void]$sbJoueurs.AppendLine("                                        <tbody>")
        
        $itemsTries = $items.GetEnumerator() | Sort-Object { $_.Value.Montant } -Descending
        
        foreach ($itemEntry in $itemsTries) {
            $itemNom = $itemEntry.Key
            $quantite = $itemEntry.Value.Quantite
            $montant = $itemEntry.Value.Montant
            
            [void]$sbJoueurs.AppendLine("                                            <tr>")
            [void]$sbJoueurs.AppendLine("                                                <td><strong>$itemNom</strong></td>")
            [void]$sbJoueurs.AppendLine("                                                <td>x$quantite</td>")
            [void]$sbJoueurs.AppendLine("                                                <td class='montant-vente'>$" + $montant.ToString('N2') + "</td>")
            [void]$sbJoueurs.AppendLine("                                            </tr>")
        }
        
        [void]$sbJoueurs.AppendLine("                                            <tr class='total-row'>")
        [void]$sbJoueurs.AppendLine("                                                <td colspan='2'><strong>TOTAL $joueur</strong></td>")
        [void]$sbJoueurs.AppendLine("                                                <td class='montant-vente'>$" + $totalVente.ToString('N2') + "</td>")
        [void]$sbJoueurs.AppendLine("                                            </tr>")
        [void]$sbJoueurs.AppendLine("                                        </tbody>")
        [void]$sbJoueurs.AppendLine("                                    </table>")
        [void]$sbJoueurs.AppendLine("                                </div>")
        [void]$sbJoueurs.AppendLine("                            </div>")
        [void]$sbJoueurs.AppendLine("                        </div>")
    }
    
    [void]$sbJoueurs.AppendLine("                    </div>")
    [void]$sbJoueurs.AppendLine("                </div>")
    [void]$sbJoueurs.AppendLine("            </div>")
}

[void]$sbJoueurs.AppendLine('        </div>')
[void]$sbJoueurs.AppendLine('    </div>')
[void]$sbJoueurs.AppendLine('    <script>')
[void]$sbJoueurs.AppendLine('        function toggleAccordion(id) {')
[void]$sbJoueurs.AppendLine('            document.getElementById(id).classList.toggle("active");')
[void]$sbJoueurs.AppendLine('            document.getElementById("icon_" + id).classList.toggle("active");')
[void]$sbJoueurs.AppendLine('        }')
[void]$sbJoueurs.AppendLine('        function toggleJoueur(id) {')
[void]$sbJoueurs.AppendLine('            document.getElementById(id).classList.toggle("active");')
[void]$sbJoueurs.AppendLine('            document.getElementById("icon_" + id).classList.toggle("active");')
[void]$sbJoueurs.AppendLine('        }')
[void]$sbJoueurs.AppendLine('        function toggleAllMain() {')
[void]$sbJoueurs.AppendLine('            var btn = document.querySelector(".expand-all-btn");')
[void]$sbJoueurs.AppendLine('            var isExpanding = btn.textContent.includes("Deplier");')
[void]$sbJoueurs.AppendLine('            document.querySelectorAll(".accordion-content").forEach(el => isExpanding ? el.classList.add("active") : el.classList.remove("active"));')
[void]$sbJoueurs.AppendLine('            document.querySelectorAll(".toggle-icon").forEach(el => isExpanding ? el.classList.add("active") : el.classList.remove("active"));')
[void]$sbJoueurs.AppendLine('            btn.textContent = isExpanding ? "Tout Replier" : "Tout Deplier";')
[void]$sbJoueurs.AppendLine('        }')
[void]$sbJoueurs.AppendLine('    </script>')
[void]$sbJoueurs.AppendLine('</body>')
[void]$sbJoueurs.AppendLine('</html>')

$sbJoueurs.ToString() | Out-File -FilePath $RapportJoueurs -Encoding UTF8
Write-Host "  => $RapportJoueurs" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RAPPORTS GENERES !" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

Start-Process $RapportItems
Start-Sleep -Seconds 1
Start-Process $RapportGlobal
Start-Sleep -Seconds 1
Start-Process $RapportJoueurs