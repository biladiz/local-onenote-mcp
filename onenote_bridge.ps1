param(
    [string]$Cmd   = "",
    [string]$P1    = "",
    [string]$P2    = "",
    [string]$P3    = "",
    [switch]$Loop
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"
$NS = "http://schemas.microsoft.com/office/onenote/2013/main"

# ── Output helpers ────────────────────────────────────────────────────

function Send-Ok {
    param($Data)
    $json = (@{ status = "ok"; data = $Data } | ConvertTo-Json -Compress -Depth 20)
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function Send-Err {
    param([string]$Message)
    $json = (@{ status = "error"; error = $Message } | ConvertTo-Json -Compress)
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

# ── OneNote COM ───────────────────────────────────────────────────

$script:on = $null

function Get-ON {
    if ($null -eq $script:on) { $script:on = New-Object -ComObject OneNote.Application }
    return $script:on
}

function Get-Hierarchy {
    param([string]$StartID = "", [int]$Scope = 2)
    $xmlOut = ""
    (Get-ON).GetHierarchy($StartID, $Scope, [ref]$xmlOut, 0)
    return [xml]$xmlOut
}

# ── Local notebook filter ────────────────────────────────────────────

function Get-LocalNotebooks {
    <# Returns hashtable: name -> ID for local (non-http) notebooks only. #>
    $xml   = Get-Hierarchy -Scope 2
    $ids   = @{}
    $nodes = $xml.GetElementsByTagName("one:Notebook")
    if ($nodes.Count -eq 0) { $nodes = $xml.GetElementsByTagName("Notebook") }
    foreach ($nb in $nodes) {
        if ([string]$nb.path -notlike "http*" -and -not $ids.ContainsKey($nb.name)) {
            $ids[$nb.name] = $nb.ID
        }
    }
    return $ids
}

function Find-NotebookID {
    param([string]$Name)
    $nbs = Get-LocalNotebooks
    if ($nbs.ContainsKey($Name)) { return $nbs[$Name] }
    foreach ($k in $nbs.Keys) { if ($k -ilike "*$Name*") { return $nbs[$k] } }
    return $null
}

# ── Sections ──────────────────────────────────────────────────────

function Get-SectionMap {
    param([string]$NotebookID)
    $xml   = Get-Hierarchy -StartID $NotebookID -Scope 1
    $ids   = @{}
    $nodes = $xml.GetElementsByTagName("one:Section")
    if ($nodes.Count -eq 0) { $nodes = $xml.GetElementsByTagName("Section") }
    foreach ($s in $nodes) {
        if ($s.name -and $s.ID -and -not $ids.ContainsKey($s.name)) {
            $ids[$s.name] = $s.ID
        }
    }
    return $ids
}

function Find-SectionID {
    param([string]$NotebookID, [string]$SectionName)
    $secs = Get-SectionMap -NotebookID $NotebookID
    if ($secs.ContainsKey($SectionName)) { return $secs[$SectionName] }
    foreach ($k in $secs.Keys) { if ($k -ilike "*$SectionName*") { return $secs[$k] } }
    return $null
}

# ── Pages ──────────────────────────────────────────────────────────

function Get-PageList {
    param([string]$ScopeID)
    $xml   = Get-Hierarchy -StartID $ScopeID -Scope 4
    $nodes = $xml.GetElementsByTagName("one:Page")
    if ($nodes.Count -eq 0) { $nodes = $xml.GetElementsByTagName("Page") }
    $pages = @()
    foreach ($p in $nodes) {
        $pages += [ordered]@{ name = $p.name; id = $p.ID; modified = $p.lastModifiedTime }
    }
    return $pages
}

function Get-PageText {
    param([string]$PageID, [switch]$StripHtml)
    $xmlOut = ""
    (Get-ON).GetPageContent($PageID, [ref]$xmlOut, 1, 0)
    $text = ""
    [regex]::Matches($xmlOut, '<one:T[^>]*><!\[CDATA\[([^\]]*(?:\][^\]])*?)\]\]></one:T>') | ForEach-Object {
        $text += $_.Groups[1].Value + "`n"
    }
    $text = $text.Trim()
    if ($StripHtml) {
        $text = [regex]::Replace($text, '<[^>]+>', '')
        $text = $text -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&nbsp;', ' ' -replace '&#[0-9]+;', ''
        $text = [regex]::Replace($text, '\s{2,}', ' ')
    }
    return $text
}

# ── Command dispatcher ───────────────────────────────────────────────────

function Invoke-Cmd {
    param([string]$Cmd, [string]$P1 = "", [string]$P2 = "", [string]$P3 = "")

    switch ($Cmd.ToLower()) {

        "list" {
            $nbs = Get-LocalNotebooks
            Send-Ok ($nbs.Keys | Sort-Object)
        }

        "getnotebook" {
            $nbs = Get-LocalNotebooks
            if ($nbs.ContainsKey($P1)) {
                Send-Ok ([ordered]@{ name = $P1; id = $nbs[$P1] })
            } else {
                Send-Err "Notebook '$P1' not found. Available: $($nbs.Keys -join ', ')"
            }
        }

        "sections" {
            $nbId = Find-NotebookID -Name $P1
            if (-not $nbId) { Send-Err "Notebook '$P1' not found."; return }
            Send-Ok ((Get-SectionMap -NotebookID $nbId).Keys | Sort-Object)
        }

        "listpages" {
            $nbId = Find-NotebookID -Name $P1
            if (-not $nbId) { Send-Err "Notebook '$P1' not found."; return }
            $secId = Find-SectionID -NotebookID $nbId -SectionName $P2
            if (-not $secId) { Send-Err "Section '$P2' not found."; return }
            Send-Ok (Get-PageList -ScopeID $secId)
        }

        "readnotebook" {
            $nbId = Find-NotebookID -Name $P1
            if (-not $nbId) { Send-Err "Notebook '$P1' not found."; return }
            $output = @()
            foreach ($p in (Get-PageList -ScopeID $nbId)) {
                $output += [ordered]@{ page = $p.name; text = (Get-PageText -PageID $p.id) }
            }
            Send-Ok $output
        }

        "readpage"   { Send-Ok (Get-PageText -PageID $P1) }
        "exporttext" { Send-Ok (Get-PageText -PageID $P1 -StripHtml) }

        "pagemetadata" {
            $fullXml  = Get-Hierarchy -Scope 4
            $allPages = $fullXml.GetElementsByTagName("one:Page")
            if ($allPages.Count -eq 0) { $allPages = $fullXml.GetElementsByTagName("Page") }
            $found = $null
            foreach ($p in $allPages) { if ($p.ID -ieq $P1) { $found = $p; break } }
            if (-not $found) { Send-Err "Page ID '$P1' not found in local notebooks."; return }
            $sec = $found.ParentNode
            $nb  = $sec
            while ($nb -and $nb.LocalName -ne "Notebook") { $nb = $nb.ParentNode }
            Send-Ok ([ordered]@{
                id           = $P1
                name         = $found.name
                section      = $sec.name
                notebook     = if ($nb) { $nb.name } else { "" }
                lastModified = $found.lastModifiedTime
            })
        }

        "search" {
            $xmlOut = ""
            (Get-ON).FindPages("", $P1, [ref]$xmlOut, $false, $false)
            $searchXml = [xml]$xmlOut
            $localNbs  = Get-LocalNotebooks
            $results   = @()
            $allPages  = $searchXml.GetElementsByTagName("one:Page")
            if ($allPages.Count -eq 0) { $allPages = $searchXml.GetElementsByTagName("Page") }
            foreach ($p in $allPages) {
                $parentNb = $p.ParentNode
                while ($parentNb -and $parentNb.LocalName -ne "Notebook") { $parentNb = $parentNb.ParentNode }
                if ($parentNb -and $localNbs.Values -contains $parentNb.ID) {
                    $results += [ordered]@{ name = $p.name; id = $p.ID }
                }
            }
            Send-Ok $results
        }

        "lastupdated" {
            $localNbs = Get-LocalNotebooks
            $fullXml  = Get-Hierarchy -Scope 4
            $allPages = $fullXml.GetElementsByTagName("one:Page")
            if ($allPages.Count -eq 0) { $allPages = $fullXml.GetElementsByTagName("Page") }
            $best = $null
            foreach ($p in $allPages) {
                $parentNb = $p
                while ($parentNb -and $parentNb.LocalName -ne "Notebook") { $parentNb = $parentNb.ParentNode }
                if ($parentNb -and $localNbs.Values -contains $parentNb.ID) {
                    if ($null -eq $best -or $p.lastModifiedTime -gt $best.lastModifiedTime) { $best = $p }
                }
            }
            if ($best) {
                Send-Ok ([ordered]@{ name = $best.name; id = $best.ID; modified = $best.lastModifiedTime })
            } else { Send-Err "No local pages found." }
        }

        "lastpages" {
            $limit    = if ($P2 -and $P2 -match '^\d+$') { [int]$P2 } else { 5 }
            $startID  = ""
            $localNbs = Get-LocalNotebooks
            if ($P1) {
                $startID = Find-NotebookID -Name $P1
                if (-not $startID) { Send-Err "Local notebook '$P1' not found."; return }
            }
            $scopeXml = Get-Hierarchy -StartID $startID -Scope 4
            $results  = @()
            $items    = $scopeXml.GetElementsByTagName("*") | Where-Object { $_.LocalName -in @("Page","Section","SectionGroup","Notebook") }
            foreach ($item in $items) {
                $isLocal = $true
                if (-not $startID) {
                    $parentNb = $item
                    while ($parentNb -and $parentNb.LocalName -ne "Notebook") { $parentNb = $parentNb.ParentNode }
                    if ($parentNb -and $localNbs.Values -notcontains $parentNb.ID) { $isLocal = $false }
                }
                if ($isLocal -and $item.lastModifiedTime) {
                    $results += [ordered]@{ type = $item.LocalName; name = $item.name; id = $item.ID; modified = $item.lastModifiedTime }
                }
            }
            Send-Ok ($results | Sort-Object { $_.modified } -Descending | Select-Object -First $limit)
        }

        "createsection" {
            $nbId = Find-NotebookID -Name $P1
            if (-not $nbId) { Send-Err "Notebook '$P1' not found."; return }
            $createXml = "<one:Notebooks xmlns:one=""$NS""><one:Notebook ID=""$nbId""><one:Section name=""$P2""/></one:Notebook></one:Notebooks>"
            (Get-ON).UpdateHierarchy($createXml)
            Send-Ok ([ordered]@{ created = $true; section = $P2; notebook = $P1 })
        }

        "createpage" {
            $nbId  = Find-NotebookID -Name $P1
            if (-not $nbId) { Send-Err "Notebook '$P1' not found."; return }
            $secId = Find-SectionID -NotebookID $nbId -SectionName $P2
            if (-not $secId) { Send-Err "Section '$P2' not found."; return }
            $pageID = ""
            (Get-ON).CreateNewPage($secId, [ref]$pageID, 0)
            $xmlContent = @"
<one:Page xmlns:one="$NS" ID="$pageID">
  <one:Title><one:OE><one:T><![CDATA[$P3]]></one:T></one:OE></one:Title>
</one:Page>
"@
            (Get-ON).UpdatePageContent($xmlContent)
            Send-Ok ([ordered]@{ created = $true; title = $P3; pageId = $pageID; section = $P2; notebook = $P1 })
        }

        "updatepage" {
            $mode        = if ($P3 -ieq "replace") { "replace" } else { "append" }
            $safeContent = $P2 -replace '\]\]>', ']]]]><![CDATA[>'
            $xmlOut      = ""
            (Get-ON).GetPageContent($P1, [ref]$xmlOut, 1, 0)
            $pageXml = [xml]$xmlOut
            if ($mode -eq "replace") {
                foreach ($o in @($pageXml.GetElementsByTagName("one:Outline"))) {
                    $o.ParentNode.RemoveChild($o) | Out-Null
                }
            }
            $outline = $pageXml.GetElementsByTagName("one:Outline") | Select-Object -First 1
            if (-not $outline) {
                $outline    = $pageXml.CreateElement("one:Outline", $NS)
                $oeChildren = $pageXml.CreateElement("one:OEChildren", $NS)
                $outline.AppendChild($oeChildren) | Out-Null
                $pageXml.Page.AppendChild($outline) | Out-Null
            }
            $oeChildrenNode = $outline.GetElementsByTagName("one:OEChildren") | Select-Object -First 1
            if (-not $oeChildrenNode) {
                $oeChildrenNode = $pageXml.CreateElement("one:OEChildren", $NS)
                $outline.AppendChild($oeChildrenNode) | Out-Null
            }
            $newOE = $pageXml.CreateElement("one:OE", $NS)
            $newT  = $pageXml.CreateElement("one:T", $NS)
            $newT.AppendChild($pageXml.CreateCDataSection($safeContent)) | Out-Null
            $newOE.AppendChild($newT) | Out-Null
            $oeChildrenNode.AppendChild($newOE) | Out-Null
            (Get-ON).UpdatePageContent($pageXml.OuterXml)
            Send-Ok ([ordered]@{ updated = $true; pageId = $P1; mode = $mode })
        }

        default { Send-Err "Unknown command: $Cmd" }
    }
}

# ── Entry point ────────────────────────────────────────────────────────

$script:on = New-Object -ComObject OneNote.Application

if ($Loop) {
    [Console]::Out.WriteLine('{"status":"ready"}')
    [Console]::Out.Flush()
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ($line -eq "") { continue }
        try {
            $req = $line | ConvertFrom-Json
            $c1  = [string]$(if ($null -ne $req.p1) { $req.p1 } else { "" })
            $c2  = [string]$(if ($null -ne $req.p2) { $req.p2 } else { "" })
            $c3  = [string]$(if ($null -ne $req.p3) { $req.p3 } else { "" })
            Invoke-Cmd -Cmd ([string]$req.cmd) -P1 $c1 -P2 $c2 -P3 $c3
        } catch {
            Send-Err $_.Exception.Message
        }
    }
} else {
    try {
        Invoke-Cmd -Cmd $Cmd -P1 $P1 -P2 $P2 -P3 $P3
    } catch {
        Send-Err $_.Exception.Message
    }
}
