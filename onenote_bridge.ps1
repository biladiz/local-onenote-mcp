param(
    [Parameter(Mandatory=$true)][string]$Cmd,
    [string]$P1 = "",
    [string]$P2 = "",
    [string]$P3 = ""
)

$ErrorActionPreference = "Stop"
$on = New-Object -ComObject OneNote.Application
$NS = "http://schemas.microsoft.com/office/onenote/2013/main"

function Get-OneNoteHierarchy {
    param([string]$StartID, [int]$Scope)
    $xmlOut = ""
    $on.GetHierarchy($StartID, $Scope, [ref]$xmlOut, 0)
    return $xmlOut
}

function Get-NotebookIDs {
    $xmlOut = Get-OneNoteHierarchy -StartID "" -Scope 2
    $xml = [xml]$xmlOut
    $ids = @{}

    # Handle both prefixed and non-prefixed attributes/elements
    $notebooks = $xml.Notebooks.Notebook
    if ($null -eq $notebooks -and $xml.ChildNodes.Count -gt 0) {
        $notebooks = $xml.GetElementsByTagName("one:Notebook")
        if ($notebooks.Count -eq 0) {
            $notebooks = $xml.GetElementsByTagName("Notebook")
        }
    }

    foreach ($nb in $notebooks) {
        $name = $nb.name
        $id = $nb.ID
        $path = $nb.path

        # FILTER: Only include local notebooks (Exclude OneDrive/https paths)
        if ($path -notlike "http*") {
            if (-not $ids.ContainsKey($name)) {
                $ids[$name] = $id
            }
        }
    }
    return $ids
}

function Get-SectionIDs {
    param([string]$NotebookID)
    $xml = Get-OneNoteHierarchy -StartID $NotebookID -Scope 1
    $ids = @{}
    $matchPattern = 'name="([^"]+)"[^>]*ID="([^"]+)"'
    $matches = [regex]::Matches($xml, $matchPattern)
    foreach ($m in $matches) {
        $name = $m.Groups[1].Value
        $id = $m.Groups[2].Value
        $ids[$name] = $id
    }
    return $ids
}

function Get-PageIDs {
    param([string]$NotebookID)
    $xml = Get-OneNoteHierarchy -StartID $NotebookID -Scope 4
    $ids = @()
    $matches = [regex]::Matches($xml, 'name="([^"]+)"[^>]*ID="([^"]+)"')
    foreach ($m in $matches) {
        $ids += @{name=$m.Groups[1].Value; id=$m.Groups[2].Value}
    }
    return $ids
}

function Get-PageText {
    param([string]$PageID)
    $xmlOut = ""
    $on.GetPageContent($PageID, [ref]$xmlOut, 1, 0)
    $text = ""
    $matches = [regex]::Matches($xmlOut, '<one:T[^>]*><!\[CDATA\[([^\]]+)\]\]></one:T>')
    foreach ($m in $matches) {
        $text += $m.Groups[1].Value + "`n"
    }
    return $text.Trim()
}

function Search-Pages {
    param([string]$Query)
    $xmlOut = ""
    $on.FindPages("", $Query, [ref]$xmlOut, $false, $false)
    $results = @()
    $matches = [regex]::Matches($xmlOut, 'name="([^"]+)"[^>]*ID="([^"]+)"')
    foreach ($m in $matches) {
        $results += "$($m.Groups[1].Value) (ID: $($m.Groups[2].Value))"
    }
    if ($results.Count -eq 0) { return "No results found" }
    return "Search Results:`n" + ($results -join "`n")
}

switch ($Cmd.ToLower()) {
    "list" {
        $ids = Get-NotebookIDs
        $ids.Keys | ForEach-Object { Write-Output $_ }
    }
    "getnotebook" {
        $ids = Get-NotebookIDs
        if ($ids.ContainsKey($P1)) {
            "Notebook: $P1`nID: $($ids[$P1])"
        } else {
            "Notebook '$P1' not found. Available: $($ids.Keys -join ', ')"
        }
    }
    "readnotebook" {
        $ids = Get-NotebookIDs
        $nbName = $P1
        $nbId = $null
        foreach ($key in $ids.Keys) {
            if ($key -like $nbName) { $nbId = $ids[$key]; break }
        }
        if (-not $nbId) {
            "Notebook not found"
            exit
        }
        $pages = Get-PageIDs -NotebookID $nbId
        $output = @("--- NOTEBOOK: $P1 ---")
        foreach ($p in $pages) {
            $text = Get-PageText -PageID $p.id
            $output += "`n[Page: $($p.name)]`n$text"
        }
        $output -join "`n"
    }
    "sections" {
        $ids = Get-NotebookIDs
        $nbName = $P1
        $nbId = $null
        foreach ($key in $ids.Keys) {
            if ($key -like $nbName) { $nbId = $ids[$key]; break }
        }
        if (-not $nbId) {
            "Notebook not found"
            exit
        }
        $sectionIds = Get-SectionIDs -NotebookID $nbId
        if ($sectionIds.Count -eq 0) {
            "No sections found"
        } else {
            $sectionIds.Keys | ForEach-Object { Write-Output $_ }
        }
    }
    "search" {
        Search-Pages -Query $P1
    }
    "readpage" {
        Get-PageText -PageID $P1
    }
    "lastupdated" {
        $localNbs = Get-NotebookIDs
        $localIds = $localNbs.Values

        $xmlOut = Get-OneNoteHierarchy -StartID "" -Scope 4
        $xml = [xml]$xmlOut

        $pages = $xml.GetElementsByTagName("one:Page")
        if ($pages.Count -eq 0) { $pages = $xml.GetElementsByTagName("Page") }

        $lastPage = ""
        $lastTime = ""

        foreach ($p in $pages) {
            $parentNB = $p
            while ($parentNB -and $parentNB.LocalName -ne "Notebook") {
                $parentNB = $parentNB.ParentNode
            }
            if ($parentNB -and $localIds -contains $parentNB.ID) {
                if ($p.lastModifiedTime -gt $lastTime) {
                    $lastTime = $p.lastModifiedTime
                    $lastPage = $p.name
                }
            }
        }

        if ($lastPage) {
            "Last updated (Local): $lastPage (Modified: $lastTime)"
        } else {
            "No local pages found"
        }
    }
    "lastpages" {
        $nbName = $P1
        $limit = if ($P2) { [int]$P2 } else { 5 }
        $startID = ""
        $localNbs = Get-NotebookIDs

        if ($nbName) {
            if ($localNbs.ContainsKey($nbName)) {
                $startID = $localNbs[$nbName]
            } else {
                foreach ($key in $localNbs.Keys) {
                    if ($key -ilike "*$nbName*") { $startID = $localNbs[$key]; break }
                }
            }
            if (-not $startID) {
                "Local notebook '$nbName' not found."
                exit
            }
        }

        $xmlOut = Get-OneNoteHierarchy -StartID $startID -Scope 4
        $xml = [xml]$xmlOut
        $results = @()

        $items = $xml.GetElementsByTagName("*") | Where-Object { $_.LocalName -in @("Page", "Section", "SectionGroup", "Notebook") }

        foreach ($item in $items) {
            $isLocal = $true
            if (-not $startID) {
                $parentNB = $item
                while ($parentNB -and $parentNB.LocalName -ne "Notebook") { $parentNB = $parentNB.ParentNode }
                if ($parentNB -and $localNbs.Values -notcontains $parentNB.ID) { $isLocal = $false }
            }

            if ($isLocal -and $item.lastModifiedTime) {
                $type = $item.LocalName
                if ($type -eq "SectionGroup") { $type = "Section Group" }

                $results += [PSCustomObject]@{
                    Type = $type
                    Name = $item.name
                    ID = $item.ID
                    Modified = $item.lastModifiedTime
                }
            }
        }

        if ($results.Count -eq 0) {
            "No local items found"
        } else {
            $sorted = $results | Sort-Object Modified -Descending | Select-Object -First $limit
            $sorted | ForEach-Object { "$($_.Type): $($_.Name) (Modified: $($_.Modified))" } | Out-String
        }
    }
    "createpage" {
        $notebookName = $P1
        $sectionName = $P2
        $title = $P3

        $ids = Get-NotebookIDs
        if (-not $ids.ContainsKey($notebookName)) {
            "Notebook not found"
            exit
        }

        $sectionIds = Get-SectionIDs -NotebookID $ids[$notebookName]
        if (-not $sectionIds.ContainsKey($sectionName)) {
            "Section not found. Available: $($sectionIds.Keys -join ', ')"
            exit
        }

        $pageID = ""
        $on.CreateNewPage($sectionIds[$sectionName], [ref]$pageID, 0)

        $xmlContent = @"
<one:Page xmlns:one="$NS" ID="$pageID">
    <one:Title><one:OE><one:T><![CDATA[$title]]></one:T></one:OE></one:Title>
    <one:Outline><one:OEChildren><one:OE><one:T><![CDATA[$title]]></one:T></one:OE></one:OEChildren></one:Outline>
</one:Page>
"@
        $on.UpdatePageContent($xmlContent)
        "Created page: $title in section $sectionName"
    }
    default {
        "Unknown command: $Cmd"
    }
}
