<#
.SYNOPSIS
    ACPackageSplitter - splits a large Adobe Campaign (v7 / v8) package XML into smaller packages
    that import cleanly, one after another, without breaking XML structure or entity dependencies.

.DESCRIPTION
    Runs on any Windows PC with the built-in Windows PowerShell 5.1 / .NET Framework. Nothing to install.

    What it guarantees
      * Every part is a complete, well-formed package: same <package> header (author, build, version),
        the original <entities schema="..."> wrappers, and whole top-level entities only (never cut mid-object).
      * Every top-level entity is written exactly once. Objects that are linked to each other
        (e.g. a delivery activity that points to a delivery owned by another campaign, or the same
        workflow / delivery defined under two campaigns) are kept in the SAME part, so no part
        overwrites or duplicates what another part imports.
      * If a group of linked objects is itself bigger than the size limit, it is spread over consecutive
        parts in dependency order (what is referenced always lands in an earlier part) and the manifest
        says which parts must be imported first.
      * After writing, every part is re-read and checked against the source (verification).

    Outputs (in <package name>_split\)
      00_IMPORT_ORDER.txt         import order, per-part dependencies, warnings, external dependencies
      <name>_Part-NN_of_MM.xml    the packages
      split_map.csv               which entity went to which part
      external_dependencies.csv   objects referenced but not in the package (must exist on the target)
      check_target_clashes.js     optional: run on the TARGET instance before importing, to find
                                  workflows / deliveries / campaigns that already exist there under the
                                  same internal name but belong to something else (the cause of
                                  "duplicate key value violates unique constraint xtkworkflow_id").

.EXAMPLE
    Split-ACPackage.bat                                   (double-click: opens the window)
    Split-ACPackage.bat -InputFile "C:\pkg\big.xml" -MaxSizeMB 10
    Split-ACPackage.bat -Analyze -InputFile "C:\pkg\part1.xml" "C:\pkg\part2.xml"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)][string[]]$InputFile,
    [string]$OutputFolder,
    [double]$MaxSizeMB = 10,
    [int]$MaxEntities = 0,
    [switch]$Analyze,
    [switch]$NoTargetCheck,
    [switch]$Gui,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$MoreFiles
)

$ErrorActionPreference = 'Stop'
$script:Version = '1.0.0'
$script:LogBox  = $null
$script:LogFile = $null

# ----------------------------------------------------------------------------------------------
#  Small helpers
# ----------------------------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0}  {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    if ($script:LogBox) {
        $script:LogBox.AppendText($line + [Environment]::NewLine)
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'OK' { 'Green' } default { 'Gray' } }
        Write-Host $line -ForegroundColor $color
    }
}

function New-List    { , (New-Object 'System.Collections.Generic.List[object]') }
function New-IntList { , (New-Object 'System.Collections.Generic.List[int]') }
function New-StrDict { , (New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::Ordinal)) }
function New-StrSet  { , (New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)) }

function Format-Size([double]$Bytes) {
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0} B' -f $Bytes)
}

function Open-XmlReader([string]$Path) {
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read, 1048576)
    return [System.Xml.XmlReader]::Create($fs, (New-ReaderSettings))
}

function New-ReaderSettings {
    $s = New-Object System.Xml.XmlReaderSettings
    $s.CloseInput       = $true
    $s.DtdProcessing    = [System.Xml.DtdProcessing]::Ignore
    $s.IgnoreWhitespace = $false
    $s.IgnoreComments   = $false
    $s.CheckCharacters  = $false
    $s.XmlResolver      = $null
    return $s
}

function Get-ReaderAttrs($r) {
    $list = New-Object 'System.Collections.Generic.List[object]'
    if ($r.MoveToFirstAttribute()) {
        do {
            $list.Add([pscustomobject]@{ Name = $r.Name; Prefix = $r.Prefix; LocalName = $r.LocalName; Ns = $r.NamespaceURI; Value = $r.Value })
        } while ($r.MoveToNextAttribute())
        [void]$r.MoveToElement()
    }
    , $list
}

function Get-AttrSignature($attrs) {
    ($attrs | ForEach-Object { '{0}={1}' -f $_.Name, $_.Value }) -join ' | '
}

# Key of an element as Adobe Campaign identifies it: internalName, else namespace:name, else name
function Get-ElKey([System.Xml.XmlElement]$el) {
    $i = $el.GetAttribute('internalName'); if ($i) { return $i }
    $n = $el.GetAttribute('name'); $ns = $el.GetAttribute('namespace')
    if ($n -and $ns) { return ('{0}:{1}' -f $ns, $n) }
    if ($n) { return $n }
    return $null
}
function Get-ReaderKey($r) {
    $i = $r.GetAttribute('internalName'); if ($i) { return $i }
    $n = $r.GetAttribute('name'); $ns = $r.GetAttribute('namespace')
    if ($n -and $ns) { return ('{0}:{1}' -f $ns, $n) }
    if ($n) { return $n }
    return $null
}

function New-Context {
    [pscustomobject]@{
        Entities     = (New-List)
        DefOwner     = (New-StrDict)   # "tag|key"     -> List[int] entity indexes that DEFINE it
        DefByIName   = (New-StrDict)   # internalName  -> List[int]
        NestedCounts = (New-StrDict)   # tag -> count of nested definitions
        Warnings     = (New-List)
        Packages     = (New-List)
        External     = $null
        Duplicates   = $null
    }
}

function Add-Def($Ctx, [string]$DefKey, [int]$Idx, [string]$IName) {
    $l = $null
    if (-not $Ctx.DefOwner.TryGetValue($DefKey, [ref]$l)) { $l = New-Object 'System.Collections.Generic.List[int]'; $Ctx.DefOwner[$DefKey] = $l }
    if (-not $l.Contains($Idx)) { $l.Add($Idx) }
    if ($IName) {
        $m = $null
        if (-not $Ctx.DefByIName.TryGetValue($IName, [ref]$m)) { $m = New-Object 'System.Collections.Generic.List[int]'; $Ctx.DefByIName[$IName] = $m }
        if (-not $m.Contains($Idx)) { $m.Add($Idx) }
    }
}

# ----------------------------------------------------------------------------------------------
#  Pass 1 - scan a package: header, <entities> blocks, and for every top-level entity its size,
#  the objects it defines and the objects it references (_operation="none" links)
# ----------------------------------------------------------------------------------------------
function Read-Entity($Reader, $Pkg, [int]$Block, $Ctx) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $sub = $Reader.ReadSubtree()
    try { $doc.Load($sub) } finally { $sub.Close() }
    [void]$Reader.Read()          # step past the entity's end tag

    $root   = $doc.DocumentElement
    $tag    = $root.LocalName
    $key    = Get-ElKey $root
    $schema = if ($Block -ge 0) { $Pkg.Blocks[$Block].Schema } else { '' }
    $idx    = $Ctx.Entities.Count
    $cmpKey = if ($key) { $key } else { '<no key>' }

    $e = [pscustomobject]@{
        Index = $idx; FileIndex = $Pkg.FileIndex; File = $Pkg.Name; Block = $Block; Schema = $schema
        Tag = $tag; Key = $key; Label = $root.GetAttribute('label')
        CmpId = ('{0}|{1}|{2}' -f $schema, $tag, $cmpKey)
        Bytes = ([long][System.Text.Encoding]::UTF8.GetByteCount($root.OuterXml) + 6)
        Defs = (New-List); Refs = (New-List); IdRefs = (New-List); Checks = (New-List)
        Deps = (New-Object 'System.Collections.Generic.HashSet[int]')
        Bundle = -1; Part = 0
    }

    # what this entity defines
    if ($key) {
        $dk = '{0}|{1}' -f $tag, $key
        $e.Defs.Add($dk)
        Add-Def $Ctx $dk $idx ($root.GetAttribute('internalName'))
    }
    $seenDefs = New-StrSet
    foreach ($n in $root.SelectNodes(".//*[@internalName and not(@_operation='none')]")) {
        $iname = $n.GetAttribute('internalName')
        $dk = '{0}|{1}' -f $n.LocalName, $iname
        if (-not $seenDefs.Add($dk)) {
            $Ctx.Warnings.Add(("'{0}' is defined more than once inside {1} '{2}' ({3}). The last definition wins on import." -f $dk, $tag, $cmpKey, $Pkg.Name))
            continue
        }
        $e.Defs.Add($dk)
        Add-Def $Ctx $dk $idx $iname
        $c = 0; [void]$Ctx.NestedCounts.TryGetValue($n.LocalName, [ref]$c); $Ctx.NestedCounts[$n.LocalName] = [int]$c + 1
    }

    # what it references (links resolved by key on import)
    $seenRefs = New-StrSet
    foreach ($n in $root.SelectNodes(".//*[@_operation='none']")) {
        $k = Get-ElKey $n
        if (-not $k) { continue }
        $rk = '{0}|{1}' -f $n.LocalName, $k
        if ($seenRefs.Add($rk)) {
            $e.Refs.Add([pscustomobject]@{ Key = $rk; Tag = $n.LocalName; Value = $k; IsIName = [bool]$n.GetAttribute('internalName') })
        }
    }

    # links written as numeric record IDs: they point to on-prem IDs and will not resolve on another instance
    foreach ($n in $root.SelectNodes('.//*[@id and @_cs]')) {
        $v = $n.GetAttribute('id')
        if ($v -match '^\d+$') { $t = ("<{0} id=""{1}""> '{2}'" -f $n.LocalName, $v, $n.GetAttribute('_cs')); if (-not $e.IdRefs.Contains($t)) { $e.IdRefs.Add($t) } }
    }

    # names to check on the target (workflow / delivery / campaign), with the campaign they must belong to
    $checkTags = @('workflow', 'delivery', 'operation')
    $nodes = New-Object 'System.Collections.Generic.List[System.Xml.XmlElement]'
    if ($checkTags -contains $tag -and $root.GetAttribute('internalName')) { $nodes.Add($root) }
    foreach ($n in $root.SelectNodes(".//workflow[@internalName and not(@_operation='none')] | .//delivery[@internalName and not(@_operation='none')] | .//operation[@internalName and not(@_operation='none')]")) { $nodes.Add($n) }
    foreach ($n in $nodes) {
        $exp = ''
        if ($n.LocalName -ne 'operation') {
            $op = $n.SelectSingleNode("operation[@_operation='none']")
            if ($op) { $exp = $op.GetAttribute('internalName') }
            elseif ($tag -eq 'operation') { $exp = $root.GetAttribute('internalName') }
        }
        $e.Checks.Add([pscustomobject]@{ Tag = $n.LocalName; Name = $n.GetAttribute('internalName'); Op = $exp })
    }

    $Ctx.Entities.Add($e)
}

function Read-PackageFile([string]$Path, [int]$FileIndex, $Ctx) {
    $fi = New-Object System.IO.FileInfo($Path)
    $pkg = [pscustomobject]@{
        Path = $fi.FullName; Name = $fi.Name; FileIndex = $FileIndex; SizeBytes = $fi.Length
        RootAttrs = (New-List); Blocks = (New-List); Extra = (New-List); First = $Ctx.Entities.Count; Count = 0
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $lastTick = 0
    $reader = Open-XmlReader $fi.FullName
    try {
        $block = -1
        [void]$reader.Read()
        while (-not $reader.EOF) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $d = $reader.Depth
                if ($d -eq 0) {
                    if ($reader.LocalName -ne 'package') { throw ("{0}: root element is <{1}>, expected <package>. Is this an Adobe Campaign package export?" -f $fi.Name, $reader.Name) }
                    $pkg.RootAttrs = Get-ReaderAttrs $reader
                    [void]$reader.Read(); continue
                }
                if ($d -eq 1) {
                    if ($reader.LocalName -eq 'entities') {
                        $attrs = Get-ReaderAttrs $reader
                        $pkg.Blocks.Add([pscustomobject]@{ Index = $pkg.Blocks.Count; Schema = $reader.GetAttribute('schema'); Attrs = $attrs; Count = 0 })
                        $block = $pkg.Blocks.Count - 1
                        [void]$reader.Read(); continue
                    }
                    $Ctx.Warnings.Add(("{0}: unexpected <{1}> directly under <package>; it is copied into every part." -f $fi.Name, $reader.Name))
                    $pkg.Extra.Add($reader.ReadOuterXml()); continue
                }
                if ($d -eq 2) {
                    Read-Entity $reader $pkg $block $Ctx
                    $pkg.Blocks[$block].Count++
                    $pkg.Count++
                    if ($sw.ElapsedMilliseconds - $lastTick -gt 3000) {
                        $lastTick = $sw.ElapsedMilliseconds
                        Write-Log ('   ... {0} entities read' -f $pkg.Count)
                    }
                    continue
                }
            }
            [void]$reader.Read()
        }
    } finally { $reader.Close() }
    if ($pkg.Count -eq 0) { throw ("{0}: no entities found under <package><entities>." -f $fi.Name) }
    $Ctx.Packages.Add($pkg)
    return $pkg
}

# ----------------------------------------------------------------------------------------------
#  Dependency graph
# ----------------------------------------------------------------------------------------------
function Resolve-Dependencies($Ctx) {
    $ext = New-StrDict
    foreach ($e in $Ctx.Entities) {
        foreach ($r in $e.Refs) {
            $owners = $null
            if (-not $Ctx.DefOwner.TryGetValue($r.Key, [ref]$owners)) {
                $owners = $null
                if ($r.IsIName) { [void]$Ctx.DefByIName.TryGetValue($r.Value, [ref]$owners) }
            }
            if ($null -ne $owners) {
                foreach ($o in $owners) { if ($o -ne $e.Index) { [void]$e.Deps.Add($o) } }
            } else {
                $x = $null
                if (-not $ext.TryGetValue($r.Key, [ref]$x)) {
                    $x = [pscustomobject]@{ Tag = $r.Tag; Value = $r.Value; Count = 0; Sample = ('{0} {1}' -f $e.Tag, $e.Key) }
                    $ext[$r.Key] = $x
                }
                $x.Count++
            }
        }
    }
    $Ctx.External = $ext

    # the same object defined by more than one top-level entity: importing both would overwrite one with the other
    $dups = New-List
    foreach ($kv in $Ctx.DefOwner.GetEnumerator()) {
        if ($kv.Value.Count -gt 1) {
            $own = @($kv.Value)
            $dups.Add([pscustomobject]@{ Key = $kv.Key; Owners = $own })
            foreach ($a in $own) { foreach ($b in $own) { if ($a -ne $b) { [void]$Ctx.Entities[$a].Deps.Add($b) } } }
        }
    }
    $Ctx.Duplicates = $dups
}

# groups of entities that are linked to each other (union-find over the links, both directions)
function Get-Bundles($Ctx) {
    $n = $Ctx.Entities.Count
    $parent = New-Object 'int[]' $n
    for ($i = 0; $i -lt $n; $i++) { $parent[$i] = $i }
    $find = { param($x) while ($parent[$x] -ne $x) { $parent[$x] = $parent[$parent[$x]]; $x = $parent[$x] }; $x }
    foreach ($e in $Ctx.Entities) {
        foreach ($d in $e.Deps) {
            $a = & $find $e.Index; $b = & $find $d
            if ($a -ne $b) { if ($a -lt $b) { $parent[$b] = $a } else { $parent[$a] = $b } }
        }
    }
    $map = @{}
    $bundles = New-Object 'System.Collections.Generic.List[object]'
    foreach ($e in $Ctx.Entities) {
        $r = [int](& $find $e.Index)
        if (-not $map.ContainsKey($r)) {
            $nb = [pscustomobject]@{ Id = $bundles.Count; Members = (New-IntList); Bytes = [long]0; Spread = $false }
            $map[$r] = $nb; $bundles.Add($nb)
        }
        $b = $map[$r]; $b.Members.Add($e.Index); $b.Bytes += $e.Bytes; $e.Bundle = $b.Id
    }
    , $bundles
}

# strongly connected components of a linked group, dependencies first (iterative Tarjan)
function Get-SccOrder($Ctx, $Members) {
    $inB = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($m in $Members) { [void]$inB.Add($m) }
    $index = @{}; $low = @{}
    $onStack = New-Object 'System.Collections.Generic.HashSet[int]'
    $stack = New-Object 'System.Collections.Generic.Stack[int]'
    $result = New-Object 'System.Collections.Generic.List[object]'
    $counter = 0
    foreach ($s in $Members) {
        if ($index.ContainsKey($s)) { continue }
        $work = New-Object 'System.Collections.Generic.Stack[object]'
        $work.Push([pscustomobject]@{ V = $s; E = $null })
        while ($work.Count -gt 0) {
            $f = $work.Peek(); $v = $f.V
            if ($null -eq $f.E) {
                $index[$v] = $counter; $low[$v] = $counter; $counter++
                $stack.Push($v); [void]$onStack.Add($v)
                $nb = @($Ctx.Entities[$v].Deps | Where-Object { $inB.Contains($_) } | Sort-Object)
                $f.E = $nb.GetEnumerator()
            }
            $pushed = $false
            while ($f.E.MoveNext()) {
                $w = [int]$f.E.Current
                if (-not $index.ContainsKey($w)) { $work.Push([pscustomobject]@{ V = $w; E = $null }); $pushed = $true; break }
                elseif ($onStack.Contains($w)) { $low[$v] = [Math]::Min($low[$v], $index[$w]) }
            }
            if ($pushed) { continue }
            if ($low[$v] -eq $index[$v]) {
                $scc = New-Object 'System.Collections.Generic.List[int]'
                do { $x = $stack.Pop(); [void]$onStack.Remove($x); $scc.Add($x) } while ($x -ne $v)
                $scc.Sort()
                $result.Add($scc)
            }
            [void]$work.Pop()
            if ($work.Count -gt 0) { $p = $work.Peek().V; $low[$p] = [Math]::Min($low[$p], $low[$v]) }
        }
    }
    , $result
}

function Build-Parts($Ctx, $Bundles, [long]$MaxBytes, [int]$MaxEnt) {
    $units = New-Object 'System.Collections.Generic.List[object]'
    foreach ($b in $Bundles) {
        $fits = (($MaxBytes -le 0) -or ($b.Bytes -le $MaxBytes)) -and (($MaxEnt -le 0) -or ($b.Members.Count -le $MaxEnt))
        if ($fits -or $b.Members.Count -eq 1) {
            $units.Add([pscustomobject]@{ Members = $b.Members; Bytes = $b.Bytes; Bundle = $b.Id })
        } else {
            $b.Spread = $true
            foreach ($scc in (Get-SccOrder $Ctx $b.Members)) {
                $sz = [long]0; foreach ($i in $scc) { $sz += $Ctx.Entities[$i].Bytes }
                $units.Add([pscustomobject]@{ Members = $scc; Bytes = $sz; Bundle = $b.Id })
            }
        }
    }

    $parts = New-Object 'System.Collections.Generic.List[object]'
    $cur = $null
    foreach ($u in $units) {
        $tooBig = (($MaxBytes -gt 0) -and ($u.Bytes -gt $MaxBytes)) -or (($MaxEnt -gt 0) -and ($u.Members.Count -gt $MaxEnt))
        if ($tooBig -and -not $Bundles[$u.Bundle].Spread) {
            # a self-contained object bigger than the limit: its own part; the part being filled stays open
            $solo = [pscustomobject]@{ Number = $parts.Count + 1; Members = (New-IntList); Bytes = [long]0; DependsOn = @(); FileName = ''; Path = '' }
            $parts.Add($solo)
            foreach ($i in $u.Members) { $solo.Members.Add($i); $Ctx.Entities[$i].Part = $solo.Number }
            $solo.Bytes = $u.Bytes
            $first = $Ctx.Entities[$u.Members[0]]
            if ($u.Members.Count -eq 1) {
                $Ctx.Warnings.Add(("{0} '{1}' alone is {2}, above the limit. It is placed on its own in part {3} (an entity cannot be cut)." -f $first.Tag, $first.Key, (Format-Size $u.Bytes), $solo.Number))
            } else {
                $Ctx.Warnings.Add(("{0} linked entities (starting with {1} '{2}') total {3}, above the limit. They must be imported together, so they share part {4}." -f $u.Members.Count, $first.Tag, $first.Key, (Format-Size $u.Bytes), $solo.Number))
            }
            continue
        }
        if ($null -ne $cur -and $cur.Members.Count -gt 0) {
            $over = (($MaxBytes -gt 0) -and ($cur.Bytes + $u.Bytes -gt $MaxBytes)) -or (($MaxEnt -gt 0) -and ($cur.Members.Count + $u.Members.Count -gt $MaxEnt))
            if ($over) { $cur = $null }
        }
        if ($null -eq $cur) {
            $cur = [pscustomobject]@{ Number = $parts.Count + 1; Members = (New-IntList); Bytes = [long]0; DependsOn = @(); FileName = ''; Path = '' }
            $parts.Add($cur)
        }
        foreach ($i in $u.Members) { $cur.Members.Add($i); $Ctx.Entities[$i].Part = $cur.Number }
        $cur.Bytes += $u.Bytes
        if ($MaxBytes -gt 0 -and $u.Bytes -gt $MaxBytes) {
            $first = $Ctx.Entities[$u.Members[0]]
            if ($u.Members.Count -eq 1) {
                $Ctx.Warnings.Add(("{0} '{1}' alone is {2}, above the {3} limit. It is placed on its own in part {4} (an entity cannot be cut)." -f $first.Tag, $first.Key, (Format-Size $u.Bytes), (Format-Size $MaxBytes), $cur.Number))
            } else {
                $Ctx.Warnings.Add(("{0} entities that reference each other in a loop (starting with {1} '{2}') total {3}, above the limit. They must be imported together, so they share part {4}." -f $u.Members.Count, $first.Tag, $first.Key, (Format-Size $u.Bytes), $cur.Number))
            }
        }
    }
    foreach ($p in $parts) { $p.Members.Sort() }

    foreach ($b in $Bundles) {
        if ($b.Spread) {
            $ps = @($b.Members | ForEach-Object { $Ctx.Entities[$_].Part } | Sort-Object -Unique)
            $Ctx.Warnings.Add(("A group of {0} linked entities ({1}) is larger than the limit, so it is spread over parts {2} in dependency order. Import those parts strictly in order." -f $b.Members.Count, (Format-Size $b.Bytes), ($ps -join ', ')))
        }
    }

    # which earlier parts each part needs
    foreach ($p in $parts) {
        $need = New-Object 'System.Collections.Generic.SortedSet[int]'
        foreach ($i in $p.Members) { foreach ($d in $Ctx.Entities[$i].Deps) { $q = $Ctx.Entities[$d].Part; if ($q -ne $p.Number) { [void]$need.Add($q) } } }
        $p.DependsOn = @($need)
        foreach ($q in $need) {
            if ($q -gt $p.Number) { $Ctx.Warnings.Add(("Part {0} references objects in LATER part {1}. Import part {1} before part {0}." -f $p.Number, $q)) }
        }
    }
    , $parts
}

# ----------------------------------------------------------------------------------------------
#  Pass 2 - write the parts (streamed; entities are copied node-for-node, never re-typed)
# ----------------------------------------------------------------------------------------------
function Write-StartTag($W, [string]$Name, $Attrs) {
    $W.WriteStartElement($Name)
    foreach ($a in $Attrs) {
        if ($a.Name -eq 'xmlns' -or $a.Prefix -eq 'xmlns') { continue }   # namespace declarations are emitted by the writer
        if ($a.Prefix) { $W.WriteAttributeString($a.Prefix, $a.LocalName, $a.Ns, $a.Value) }
        else { $W.WriteAttributeString($a.LocalName, $a.Value) }
    }
}

function Write-Parts($Ctx, $Pkg, $Parts, [string]$OutDir, [string]$BaseName) {
    $total = $Parts.Count
    $digits = [Math]::Max(2, ([string]$total).Length)
    foreach ($p in $Parts) {
        $p.FileName = '{0}_Part-{1}_of_{2}.xml' -f $BaseName, $p.Number.ToString("D$digits"), $total.ToString("D$digits")
        $p.Path = Join-Path $OutDir $p.FileName
    }
    $ws = New-Object System.Xml.XmlWriterSettings
    $ws.Encoding = New-Object System.Text.UTF8Encoding($true)
    $ws.Indent = $false
    $ws.NewLineHandling = [System.Xml.NewLineHandling]::Entitize
    $ws.CheckCharacters = $false

    $writers = @{}; $curBlock = @{}
    $reader = Open-XmlReader $Pkg.Path
    try {
        $block = -1; $i = $Pkg.First
        [void]$reader.Read()
        while (-not $reader.EOF) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $d = $reader.Depth
                if ($d -eq 1) {
                    if ($reader.LocalName -eq 'entities') { $block++; [void]$reader.Read(); continue }
                    [void]$reader.ReadOuterXml(); continue
                }
                if ($d -eq 2) {
                    $e = $Ctx.Entities[$i]; $i++
                    $pn = $e.Part
                    if (-not $writers.ContainsKey($pn)) {
                        $w = [System.Xml.XmlWriter]::Create($Parts[$pn - 1].Path, $ws)
                        $w.WriteStartDocument()
                        $w.WriteWhitespace("`n")
                        Write-StartTag $w 'package' $Pkg.RootAttrs
                        foreach ($x in $Pkg.Extra) { $w.WriteWhitespace("`n  "); $w.WriteRaw($x) }
                        $writers[$pn] = $w; $curBlock[$pn] = -1
                    }
                    $w = $writers[$pn]
                    if ($curBlock[$pn] -ne $block) {
                        if ($curBlock[$pn] -ge 0) { $w.WriteWhitespace("`n  "); $w.WriteEndElement() }
                        $w.WriteWhitespace("`n  ")
                        Write-StartTag $w 'entities' $Pkg.Blocks[$block].Attrs
                        $curBlock[$pn] = $block
                    }
                    $w.WriteWhitespace("`n    ")
                    $w.WriteNode($reader, $true)      # copies the whole entity and moves the reader past it
                    continue
                }
            }
            [void]$reader.Read()
        }
    } finally {
        $reader.Close()
        foreach ($pn in @($writers.Keys)) {
            $w = $writers[$pn]
            try {
                if ($curBlock[$pn] -ge 0) { $w.WriteWhitespace("`n  "); $w.WriteEndElement() }
                $w.WriteWhitespace("`n"); $w.WriteEndElement(); $w.WriteEndDocument()
            } finally { $w.Close() }
        }
    }
}

# ----------------------------------------------------------------------------------------------
#  Verification - re-read every part and compare with the source scan
# ----------------------------------------------------------------------------------------------
function Test-Parts($Ctx, $Pkg, $Parts) {
    $errors = New-List
    $rootSig = Get-AttrSignature $Pkg.RootAttrs
    $assigned = 0
    foreach ($p in $Parts) {
        $assigned += $p.Members.Count
        $expected = New-Object 'System.Collections.Generic.List[string]'
        foreach ($i in $p.Members) { $expected.Add($Ctx.Entities[$i].CmpId) }
        $actual = New-Object 'System.Collections.Generic.List[string]'
        $r = $null
        try {
            $r = Open-XmlReader $p.Path
            $schema = ''
            [void]$r.Read()
            while (-not $r.EOF) {
                if ($r.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                    if ($r.Depth -eq 0) {
                        $sig = Get-AttrSignature (Get-ReaderAttrs $r)
                        if ($sig -ne $rootSig) { $errors.Add(("Part {0}: <package> header differs from the source." -f $p.Number)) }
                    } elseif ($r.Depth -eq 1 -and $r.LocalName -eq 'entities') {
                        $schema = $r.GetAttribute('schema')
                    } elseif ($r.Depth -eq 2) {
                        $k = Get-ReaderKey $r; if (-not $k) { $k = '<no key>' }
                        $actual.Add(('{0}|{1}|{2}' -f $schema, $r.LocalName, $k))
                        $r.Skip(); continue
                    }
                }
                [void]$r.Read()
            }
        } catch {
            $errors.Add(("Part {0} ({1}) is not well-formed XML: {2}" -f $p.Number, $p.FileName, $_.Exception.Message))
        } finally { if ($r) { $r.Close() } }
        if (($expected -join "`n") -ne ($actual -join "`n")) {
            $errors.Add(("Part {0}: expected {1} entities, found {2} (or in a different order)." -f $p.Number, $expected.Count, $actual.Count))
        }
    }
    if ($assigned -ne $Ctx.Entities.Count) { $errors.Add(("{0} entities in the source but {1} assigned to parts." -f $Ctx.Entities.Count, $assigned)) }
    foreach ($e in $Ctx.Entities) { if ($e.Part -le 0) { $errors.Add(("{0} '{1}' was not assigned to any part." -f $e.Tag, $e.Key)) } }
    , $errors
}

# ----------------------------------------------------------------------------------------------
#  Reports
# ----------------------------------------------------------------------------------------------
function Get-SchemaSummary($Pkg) {
    $order = New-Object 'System.Collections.Generic.List[string]'; $sum = @{}
    foreach ($b in $Pkg.Blocks) { if (-not $sum.ContainsKey($b.Schema)) { $order.Add($b.Schema); $sum[$b.Schema] = 0 }; $sum[$b.Schema] += $b.Count }
    ($order | ForEach-Object { '{0} ({1})' -f $_, $sum[$_] }) -join ', '
}

function Get-EntityLabel($e) {
    $s = '{0} {1}' -f $e.Tag, $(if ($e.Key) { $e.Key } else { '<no key>' })
    if ($e.Label) { $s += (" '{0}'" -f $e.Label) }
    return $s
}

function Add-CommonSections($Ctx, $sb) {
    [void]$sb.AppendLine()
    [void]$sb.AppendLine(('WARNINGS ({0})' -f $Ctx.Warnings.Count))
    if ($Ctx.Warnings.Count -eq 0) { [void]$sb.AppendLine('  none') }
    foreach ($w in $Ctx.Warnings) { [void]$sb.AppendLine('  - ' + $w) }

    [void]$sb.AppendLine()
    [void]$sb.AppendLine(('DUPLICATE DEFINITIONS ({0}) - the same object is defined by more than one top-level entity. Importing both makes the later one overwrite the earlier one; the splitter keeps them in the same part so the result is the same as importing the original file.' -f $Ctx.Duplicates.Count))
    if ($Ctx.Duplicates.Count -eq 0) { [void]$sb.AppendLine('  none') }
    foreach ($d in $Ctx.Duplicates) {
        $who = ($d.Owners | ForEach-Object { $e = $Ctx.Entities[$_]; '{0} [{1}]' -f (Get-EntityLabel $e), $e.File }) -join '; '
        [void]$sb.AppendLine(('  - {0}  defined in: {1}' -f $d.Key, $who))
    }

    $idr = @($Ctx.Entities | Where-Object { $_.IdRefs.Count -gt 0 })
    $idCount = 0; foreach ($e in $idr) { $idCount += $e.IdRefs.Count }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine(('NUMERIC ID LINKS ({0}) - these point to record IDs of the SOURCE instance and will not resolve (or will resolve to the wrong record) on the target. Re-link them after import.' -f $idCount))
    if ($idCount -eq 0) { [void]$sb.AppendLine('  none') }
    foreach ($e in $idr) { foreach ($x in $e.IdRefs) { [void]$sb.AppendLine(('  - in {0}: {1}' -f (Get-EntityLabel $e), $x)) } }

    [void]$sb.AppendLine()
    [void]$sb.AppendLine(('EXTERNAL DEPENDENCIES ({0}) - referenced but not contained in the package. They must already exist on the target (same name / internal name) before you import. Full list: external_dependencies.csv' -f $Ctx.External.Count))
    $byTag = $Ctx.External.Values | Group-Object Tag | Sort-Object Count -Descending
    foreach ($g in $byTag) {
        $names = @($g.Group | Sort-Object Value | ForEach-Object { $_.Value })
        $shown = if ($names.Count -gt 40) { ($names[0..39] -join ', ') + (' ... (+{0} more)' -f ($names.Count - 40)) } else { $names -join ', ' }
        [void]$sb.AppendLine(('  <{0}> x{1}: {2}' -f $g.Name, $names.Count, $shown))
    }
}

function Write-ExternalCsv($Ctx, [string]$Path) {
    $Ctx.External.Values | Sort-Object Tag, Value |
        Select-Object @{n = 'Link element'; e = { $_.Tag } }, @{n = 'Key (name / internalName)'; e = { $_.Value } }, @{n = 'Referenced by (count)'; e = { $_.Count } }, @{n = 'Example referencing entity'; e = { $_.Sample } } |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function ConvertTo-JsString([string]$s) {
    '"' + ($s -replace '\\', '\\' -replace '"', '\"' -replace "`r", '\r' -replace "`n", '\n') + '"'
}

function Write-TargetCheckJs($Ctx, [string]$Path, [string]$SourceName) {
    $maps = @{ workflow = (New-StrDict); delivery = (New-StrDict); operation = (New-StrDict) }
    foreach ($e in $Ctx.Entities) {
        foreach ($c in $e.Checks) {
            if ($c.Name -and -not $maps[$c.Tag].ContainsKey($c.Name)) { $maps[$c.Tag][$c.Name] = @($c.Op, $e.Part) }
        }
    }
    $toJson = {
        param($d)
        $items = foreach ($kv in $d.GetEnumerator()) { '{0}:[{1},{2}]' -f (ConvertTo-JsString $kv.Key), (ConvertTo-JsString ([string]$kv.Value[0])), [int]$kv.Value[1] }
        '{' + (@($items) -join ',') + '}'
    }
    $js = @"
// ACPackageSplitter v$($script:Version) - pre-import clash check for: $SourceName
// Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')
//
// HOW TO USE: on the TARGET instance, create a test workflow with one JavaScript code activity,
// paste this whole file into it, run it, then open the workflow journal (log).
//
// It lists every workflow / delivery / campaign in this package whose internal name ALREADY exists on
// the target but belongs to a different campaign (or to none). On import such an object is inserted
// again with the existing record's ID -> "duplicate key value violates unique constraint xtkworkflow_id".
// Fix each CLASH (delete or rename the object on the target) before importing the part shown.
// Objects already attached to the right campaign are fine: the import simply updates them.
// If sqlSelect is blocked on your instance, filter the Explorer lists on the internal names instead.

var W = $(& $toJson $maps.workflow);
var D = $(& $toJson $maps.delivery);
var O = $(& $toJson $maps.operation);

function q(s){ return "'" + String(s).replace(/'/g,"''") + "'"; }
function check(label, map, sql, compareCampaign){
  var names = []; for (var k in map) names.push(k);
  var hits = 0;
  for (var i = 0; i < names.length; i += 200){
    var inList = names.slice(i, i + 200).map(q).join(",");
    var res = sqlSelect("r,@id:string,@name:string,@label:string,@op:string,@created:datetime", sql.replace("#IN#", inList));
    for each (var r in res.r){
      var exp = map[String(r.@name)];
      if (compareCampaign && exp && String(r.@op) == exp[0]) continue;
      hits++;
      var msg = label + (compareCampaign ? " CLASH" : " ALREADY ON TARGET (will be updated)") +
                "  name=" + r.@name + "  id=" + r.@id + "  label='" + r.@label + "'" +
                "  targetCampaign='" + r.@op + "'  packageCampaign='" + (exp ? exp[0] : "") + "'" +
                "  part=" + (exp ? exp[1] : "") + "  created=" + r.@created;
      if (compareCampaign) logWarning(msg); else logInfo(msg);
    }
  }
  logInfo(label + ": " + names.length + " names checked, " + hits + (compareCampaign ? " clash(es)" : " already on target"));
}
check("WORKFLOW", W, "select w.iWorkflowId, w.sInternalName, w.sLabel, coalesce(o.sInternalName,''), w.tsCreated from XtkWorkflow w left join NmsOperation o on o.iOperationId = w.iOperationId where w.sInternalName in (#IN#)", true);
check("DELIVERY", D, "select d.iDeliveryId, d.sInternalName, d.sLabel, coalesce(o.sInternalName,''), d.tsCreated from NmsDelivery d left join NmsOperation o on o.iOperationId = d.iOperationId where d.sInternalName in (#IN#)", true);
check("CAMPAIGN", O, "select o.iOperationId, o.sInternalName, o.sLabel, '', o.tsCreated from NmsOperation o where o.sInternalName in (#IN#)", false);
"@
    [System.IO.File]::WriteAllText($Path, $js, (New-Object System.Text.UTF8Encoding($false)))
    return ($maps.workflow.Count + $maps.delivery.Count + $maps.operation.Count)
}

# ----------------------------------------------------------------------------------------------
#  Commands
# ----------------------------------------------------------------------------------------------
function Get-UniqueFolder([string]$Path) {
    if (-not [System.IO.Directory]::Exists($Path)) { return $Path }
    if (@([System.IO.Directory]::GetFileSystemEntries($Path)).Count -eq 0) { return $Path }
    return ('{0}_{1}' -f $Path, (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

function Invoke-Split {
    param([string]$Path, [string]$OutDir, [double]$MaxMB, [int]$MaxEnt, [bool]$TargetCheck)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $fi = New-Object System.IO.FileInfo($Path)
    if (-not $fi.Exists) { throw "File not found: $Path" }
    Write-Log ('Reading {0} ({1}) ...' -f $fi.Name, (Format-Size $fi.Length))
    $ctx = New-Context
    $pkg = Read-PackageFile $fi.FullName 0 $ctx
    Write-Log ('{0} top-level entities in {1} block(s): {2}' -f $pkg.Count, $pkg.Blocks.Count, (Get-SchemaSummary $pkg))

    Resolve-Dependencies $ctx
    $bundles = Get-Bundles $ctx
    $maxBytes = [long]($MaxMB * 1MB)
    $parts = Build-Parts $ctx $bundles $maxBytes $MaxEnt
    $linked = @($bundles | Where-Object { $_.Members.Count -gt 1 }).Count
    Write-Log ('{0} independent group(s) ({1} with linked entities kept together) -> {2} part(s)' -f $bundles.Count, $linked, $parts.Count)

    $base = [System.IO.Path]::GetFileNameWithoutExtension($fi.Name)
    if (-not $OutDir) { $OutDir = Join-Path $fi.DirectoryName ($base + '_split') }
    $OutDir = Get-UniqueFolder $OutDir
    [void][System.IO.Directory]::CreateDirectory($OutDir)
    Write-Log ('Writing parts to {0}' -f $OutDir)
    Write-Parts $ctx $pkg $parts $OutDir $base

    Write-Log 'Verifying parts against the source ...'
    $errs = Test-Parts $ctx $pkg $parts
    foreach ($x in $errs) { Write-Log $x 'ERROR' }

    # manifest
    $sb = New-Object System.Text.StringBuilder
    $bar = '=' * 100
    [void]$sb.AppendLine($bar)
    [void]$sb.AppendLine((' Adobe Campaign package split - {0}' -f $fi.Name))
    [void]$sb.AppendLine($bar)
    $hdr = @{}; foreach ($a in $pkg.RootAttrs) { $hdr[$a.Name] = $a.Value }
    [void]$sb.AppendLine(('Generated      : {0} by ACPackageSplitter v{1} ({2})' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $script:Version, $env:USERNAME))
    [void]$sb.AppendLine(('Source file    : {0}' -f $fi.FullName))
    [void]$sb.AppendLine(('Source size    : {0}   build {1}, version {2}, author {3}' -f (Format-Size $fi.Length), $hdr['buildNumber'], $hdr['buildVersion'], $hdr['author']))
    [void]$sb.AppendLine(('Entities       : {0} top-level - {1}' -f $pkg.Count, (Get-SchemaSummary $pkg)))
    $nested = ($ctx.NestedCounts.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { '{0} {1}' -f $_.Value, $_.Key }) -join ', '
    if ($nested) { [void]$sb.AppendLine(('Nested objects : {0}' -f $nested)) }
    [void]$sb.AppendLine(('Limits         : {0} per part{1}' -f $(if ($maxBytes -gt 0) { Format-Size $maxBytes } else { 'no size limit' }), $(if ($MaxEnt -gt 0) { ", max $MaxEnt entities" } else { '' })))
    [void]$sb.AppendLine(('Result         : {0} part(s)' -f $parts.Count))
    if ($errs.Count -eq 0) {
        [void]$sb.AppendLine('Verification   : PASSED - all parts are well-formed, every entity is written exactly once, headers match the source.')
    } else {
        [void]$sb.AppendLine(('Verification   : FAILED ({0} problem(s)) - DO NOT IMPORT these parts:' -f $errs.Count))
        foreach ($x in $errs) { [void]$sb.AppendLine('                 ' + $x) }
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('HOW TO IMPORT')
    [void]$sb.AppendLine('  1. Before the first import, run check_target_clashes.js on the target (see the top of that file) and fix every CLASH it reports.')
    [void]$sb.AppendLine('  2. Import the parts one at a time, in the order below (Tools > Advanced > Import package).')
    [void]$sb.AppendLine('  3. Read each import log before starting the next part. If a part fails, fix the cause and re-import that same part;')
    [void]$sb.AppendLine('     parts that already imported do not need to be imported again (import matches objects by internal name and updates them).')
    [void]$sb.AppendLine('  4. "independent" parts have no links to other parts. A part that depends on others must be imported after them.')
    [void]$sb.AppendLine()
    $w1 = [Math]::Max(40, (@($parts | ForEach-Object { $_.FileName.Length }) | Measure-Object -Maximum).Maximum)
    [void]$sb.AppendLine(('  {0,-4}  {1}  {2,10}  {3,8}  {4}' -f 'Part', 'File'.PadRight($w1), 'Size', 'Entities', 'Depends on parts'))
    [void]$sb.AppendLine(('  {0,-4}  {1}  {2,10}  {3,8}  {4}' -f '----', ('-' * $w1), '----------', '--------', '----------------'))
    foreach ($p in $parts) {
        $size = (New-Object System.IO.FileInfo($p.Path)).Length
        $dep = if ($p.DependsOn.Count -eq 0) { 'independent' } else { ($p.DependsOn -join ', ') }
        [void]$sb.AppendLine(('  {0,-4}  {1}  {2,10}  {3,8}  {4}' -f $p.Number, $p.FileName.PadRight($w1), (Format-Size $size), $p.Members.Count, $dep))
    }
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('CONTENTS')
    foreach ($p in $parts) {
        [void]$sb.AppendLine(('  Part {0}:' -f $p.Number))
        foreach ($i in $p.Members) { $e = $ctx.Entities[$i]; [void]$sb.AppendLine(('     {0,-60} {1,10}' -f (Get-EntityLabel $e), (Format-Size $e.Bytes))) }
    }
    Add-CommonSections $ctx $sb
    $manifest = Join-Path $OutDir '00_IMPORT_ORDER.txt'
    [System.IO.File]::WriteAllText($manifest, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    # entity map
    $ctx.Entities | ForEach-Object {
        $e = $_
        $depKeys = @($e.Deps | ForEach-Object { $d = $ctx.Entities[$_]; '{0} {1} (part {2})' -f $d.Tag, $d.Key, $d.Part }) -join '; '
        [pscustomobject]@{
            'Part' = $e.Part; 'Part file' = $parts[$e.Part - 1].FileName; 'Source order' = $e.Index + 1; 'Schema' = $e.Schema
            'Element' = $e.Tag; 'Key' = $e.Key; 'Label' = $e.Label; 'Size KB' = [Math]::Round($e.Bytes / 1KB, 1)
            'Nested objects' = [Math]::Max(0, $e.Defs.Count - 1); 'Linked group' = $e.Bundle + 1
            'Links to other entities' = $depKeys; 'Numeric ID links' = ($e.IdRefs -join '; ')
        }
    } | Export-Csv -LiteralPath (Join-Path $OutDir 'split_map.csv') -NoTypeInformation -Encoding UTF8
    Write-ExternalCsv $ctx (Join-Path $OutDir 'external_dependencies.csv')

    if ($TargetCheck) {
        $n = Write-TargetCheckJs $ctx (Join-Path $OutDir 'check_target_clashes.js') $fi.Name
        Write-Log ('check_target_clashes.js written ({0} names to check on the target)' -f $n)
    }

    foreach ($w in $ctx.Warnings) { Write-Log $w 'WARN' }
    if ($ctx.Duplicates.Count -gt 0) { Write-Log ('{0} object(s) are defined more than once in the source; kept together (see 00_IMPORT_ORDER.txt)' -f $ctx.Duplicates.Count) 'WARN' }
    foreach ($p in $parts) {
        $dep = if ($p.DependsOn.Count -eq 0) { 'independent' } else { 'after part(s) ' + ($p.DependsOn -join ', ') }
        Write-Log ('  Part {0}: {1,4} entities  {2,9}  {3}' -f $p.Number, $p.Members.Count, (Format-Size (New-Object System.IO.FileInfo($p.Path)).Length), $dep)
    }
    if ($errs.Count -eq 0) { Write-Log ('Done in {0:N0}s. Verification PASSED. Import order: 00_IMPORT_ORDER.txt' -f $sw.Elapsed.TotalSeconds) 'OK' }
    else { Write-Log ('Done in {0:N0}s, but verification FAILED - do not import. See 00_IMPORT_ORDER.txt' -f $sw.Elapsed.TotalSeconds) 'ERROR' }
    return [pscustomobject]@{ OutDir = $OutDir; Parts = $parts.Count; Passed = ($errs.Count -eq 0) }
}

function Invoke-Analyze {
    param([string[]]$Paths, [double]$MaxMB, [int]$MaxEnt)
    $files = @($Paths | ForEach-Object { New-Object System.IO.FileInfo($_) } | Sort-Object Name)
    $ctx = New-Context
    $fileIdx = 0
    foreach ($f in $files) {
        if (-not $f.Exists) { throw "File not found: $($f.FullName)" }
        Write-Log ('Reading {0} ({1}) ...' -f $f.Name, (Format-Size $f.Length))
        [void](Read-PackageFile $f.FullName $fileIdx $ctx)
        $fileIdx++
    }
    Resolve-Dependencies $ctx
    $bundles = Get-Bundles $ctx

    $sb = New-Object System.Text.StringBuilder
    $bar = '=' * 100
    [void]$sb.AppendLine($bar)
    [void]$sb.AppendLine((' Adobe Campaign package analysis - {0} file(s)' -f $files.Count))
    [void]$sb.AppendLine($bar)
    [void]$sb.AppendLine(('Generated : {0} by ACPackageSplitter v{1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), $script:Version))
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('FILES (assumed import order = alphabetical)')
    $sigs = @{}
    foreach ($p in $ctx.Packages) {
        $hdr = @{}; foreach ($a in $p.RootAttrs) { $hdr[$a.Name] = $a.Value }
        $sigs[('{0}/{1}' -f $hdr['buildNumber'], $hdr['buildVersion'])] = 1
        [void]$sb.AppendLine(('  {0,2}. {1}  {2}, {3} entities, build {4}: {5}' -f ($p.FileIndex + 1), $p.Name, (Format-Size $p.SizeBytes), $p.Count, $hdr['buildNumber'], (Get-SchemaSummary $p)))
    }
    if ($sigs.Count -gt 1) { $ctx.Warnings.Add('The files come from different build numbers. Adobe does not support importing a package into a different build.') }

    if ($files.Count -gt 1) {
        # same top-level entity in several files
        $byId = $ctx.Entities | Group-Object CmpId | Where-Object { $_.Count -gt 1 -and ($_.Group | Select-Object -ExpandProperty FileIndex -Unique).Count -gt 1 }
        [void]$sb.AppendLine()
        [void]$sb.AppendLine(('SAME ENTITY IN MORE THAN ONE FILE ({0}) - the later import overwrites the earlier one' -f @($byId).Count))
        if (@($byId).Count -eq 0) { [void]$sb.AppendLine('  none') }
        foreach ($g in $byId) { [void]$sb.AppendLine(('  - {0}: {1}' -f $g.Name, (($g.Group | ForEach-Object { $_.File }) -join ', '))) }

        # links between files
        $cross = @{}
        foreach ($e in $ctx.Entities) {
            foreach ($d in $e.Deps) {
                $t = $ctx.Entities[$d]
                if ($t.FileIndex -ne $e.FileIndex) {
                    $k = '{0}|{1}' -f $e.FileIndex, $t.FileIndex
                    if (-not $cross.ContainsKey($k)) { $cross[$k] = New-Object 'System.Collections.Generic.List[string]' }
                    $cross[$k].Add(('{0} -> {1}' -f (Get-EntityLabel $e), (Get-EntityLabel $t)))
                }
            }
        }
        [void]$sb.AppendLine()
        [void]$sb.AppendLine(('LINKS BETWEEN FILES ({0} file pair(s)) - these files cannot be imported independently' -f $cross.Count))
        if ($cross.Count -eq 0) { [void]$sb.AppendLine('  none - every file is self-contained') }
        foreach ($k in ($cross.Keys | Sort-Object)) {
            $a, $b = $k -split '\|'
            $fa = $ctx.Packages[[int]$a].Name; $fb = $ctx.Packages[[int]$b].Name
            $order = if ([int]$b -gt [int]$a) { '  <-- ORDER PROBLEM: the referenced object is in a LATER file' } else { '' }
            [void]$sb.AppendLine(('  {0}  needs  {1}  ({2} link(s)){3}' -f $fa, $fb, $cross[$k].Count, $order))
            foreach ($x in ($cross[$k] | Select-Object -First 10)) { [void]$sb.AppendLine('      ' + $x) }
        }
    }

    [void]$sb.AppendLine()
    $big = $bundles | Sort-Object Bytes -Descending | Select-Object -First 1
    [void]$sb.AppendLine('SPLIT OUTLOOK')
    [void]$sb.AppendLine(('  {0} top-level entities form {1} independent group(s). Largest group: {2} entities, {3}.' -f $ctx.Entities.Count, $bundles.Count, $big.Members.Count, (Format-Size $big.Bytes)))
    if ($ctx.Packages.Count -eq 1) {
        $parts = Build-Parts $ctx $bundles ([long]($MaxMB * 1MB)) $MaxEnt
        [void]$sb.AppendLine(('  With a limit of {0} MB{1} the split gives {2} part(s):' -f $MaxMB, $(if ($MaxEnt -gt 0) { " / $MaxEnt entities" } else { '' }), $parts.Count))
        foreach ($p in $parts) {
            $dep = if ($p.DependsOn.Count -eq 0) { 'independent' } else { 'after part(s) ' + ($p.DependsOn -join ', ') }
            [void]$sb.AppendLine(('     Part {0,-3} {1,5} entities  ~{2,9}  {3}' -f $p.Number, $p.Members.Count, (Format-Size $p.Bytes), $dep))
        }
    }
    Add-CommonSections $ctx $sb

    $first = $files[0]
    $report = Join-Path $first.DirectoryName ('{0}_analysis_{1}.txt' -f [System.IO.Path]::GetFileNameWithoutExtension($first.Name), (Get-Date -Format 'yyyyMMdd-HHmmss'))
    [System.IO.File]::WriteAllText($report, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    foreach ($line in ($sb.ToString() -split "`r?`n")) { if ($line.Length -gt 0) { Write-Log $line } }
    Write-Log ('Report saved: {0}' -f $report) 'OK'
    return $report
}

# ----------------------------------------------------------------------------------------------
#  Window
# ----------------------------------------------------------------------------------------------
function Show-Gui([string[]]$Preset) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Adobe Campaign Package Splitter  v$($script:Version)"
    $form.Size = New-Object System.Drawing.Size(980, 680)
    $form.MinimumSize = New-Object System.Drawing.Size(760, 480)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.AllowDrop = $true
    $TL  = [System.Windows.Forms.AnchorStyles]'Top,Left'
    $TLR = [System.Windows.Forms.AnchorStyles]'Top,Left,Right'
    $TR  = [System.Windows.Forms.AnchorStyles]'Top,Right'

    function New-Ctl($type, $x, $y, $w, $h, $text, $anchor) {
        $c = New-Object "System.Windows.Forms.$type"
        $c.Location = New-Object System.Drawing.Point($x, $y)
        $c.Size = New-Object System.Drawing.Size($w, $h)
        if ($null -ne $text) { $c.Text = $text }
        $c.Anchor = $anchor
        $form.Controls.Add($c)
        return $c
    }

    [void](New-Ctl Label 12 16 150 20 'Package file(s):' $TL)
    $txtIn  = New-Ctl TextBox 165 13 680 23 '' $TLR
    $btnIn  = New-Ctl Button 855 11 100 27 'Browse...' $TR
    [void](New-Ctl Label 12 50 150 20 'Output folder:' $TL)
    $txtOut = New-Ctl TextBox 165 47 680 23 '' $TLR
    $btnOut = New-Ctl Button 855 45 100 27 'Browse...' $TR
    $hint   = New-Ctl Label 165 72 680 18 'Leave empty to create "<package name>_split" next to the package. Tip: drag and drop package files onto this window.' $TLR
    $hint.ForeColor = [System.Drawing.Color]::DimGray

    [void](New-Ctl Label 12 101 150 20 'Max size per part (MB):' $TL)
    $numSize = New-Ctl NumericUpDown 165 98 80 23 $null $TL
    $numSize.DecimalPlaces = 1; $numSize.Minimum = 0; $numSize.Maximum = 5000; $numSize.Increment = 1; $numSize.Value = [decimal]10
    [void](New-Ctl Label 262 101 215 20 'Max entities per part (0 = no limit):' $TL)
    $numCnt = New-Ctl NumericUpDown 480 98 80 23 $null $TL
    $numCnt.Minimum = 0; $numCnt.Maximum = 1000000; $numCnt.Value = 0
    $chkJs = New-Ctl CheckBox 585 99 360 22 'Create target clash-check script (check_target_clashes.js)' $TL
    $chkJs.Checked = $true

    $btnAnalyze = New-Ctl Button 165 132 150 32 'Analyze (no changes)' $TL
    $btnSplit   = New-Ctl Button 325 132 150 32 'Split' $TL
    $btnSplit.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $btnOpen    = New-Ctl Button 485 132 150 32 'Open output folder' $TL
    $btnOpen.Enabled = $false

    $log = New-Ctl TextBox 12 176 943 455 '' ([System.Windows.Forms.AnchorStyles]'Top,Bottom,Left,Right')
    $log.Multiline = $true; $log.ScrollBars = 'Both'; $log.WordWrap = $false; $log.ReadOnly = $true
    $log.BackColor = [System.Drawing.Color]::White
    $log.Font = New-Object System.Drawing.Font('Consolas', 9)
    $script:LogBox = $log
    $script:LastOut = $null

    if ($Preset) { $txtIn.Text = ($Preset -join '; ') }
    Write-Log "Adobe Campaign Package Splitter v$($script:Version)"
    Write-Log 'Choose a package, then Analyze (report only) or Split. Parts are never written over the source file.'

    $getFiles = { @($txtIn.Text -split ';' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ }) }
    $setBusy = { param($busy) foreach ($b in @($btnAnalyze, $btnSplit, $btnIn, $btnOut)) { $b.Enabled = -not $busy }; $form.Cursor = $(if ($busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }) }

    $btnIn.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Package XML (*.xml)|*.xml|All files (*.*)|*.*'
        $dlg.Multiselect = $true
        $dlg.Title = 'Select Adobe Campaign package file(s)'
        if ($dlg.ShowDialog($form) -eq 'OK') { $txtIn.Text = ($dlg.FileNames -join '; ') }
    })
    $btnOut.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Folder for the split packages'
        if ($dlg.ShowDialog($form) -eq 'OK') { $txtOut.Text = $dlg.SelectedPath }
    })
    $form.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' } })
    $form.Add_DragDrop({ $txtIn.Text = (@($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) -join '; ') })
    $btnOpen.Add_Click({ if ($script:LastOut -and [System.IO.Directory]::Exists($script:LastOut)) { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:LastOut) } })

    $btnAnalyze.Add_Click({
        $files = & $getFiles
        if ($files.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show($form, 'Choose at least one package file.', 'Package splitter'); return }
        & $setBusy $true
        try {
            $log.Clear()
            $rep = Invoke-Analyze -Paths $files -MaxMB ([double]$numSize.Value) -MaxEnt ([int]$numCnt.Value)
            $script:LastOut = [System.IO.Path]::GetDirectoryName($rep); $btnOpen.Enabled = $true
        } catch { Write-Log $_.Exception.Message 'ERROR'; [void][System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Analyze failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
        finally { & $setBusy $false }
    })
    $btnSplit.Add_Click({
        $files = & $getFiles
        if ($files.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show($form, 'Choose a package file to split.', 'Package splitter'); return }
        if ($files.Count -gt 1 -and $txtOut.Text.Trim()) {
            [void][System.Windows.Forms.MessageBox]::Show($form, 'With several files, leave the output folder empty: each file gets its own "<name>_split" folder.', 'Package splitter'); return
        }
        & $setBusy $true
        try {
            $log.Clear()
            $allOk = $true
            foreach ($f in $files) {
                $res = Invoke-Split -Path $f -OutDir $txtOut.Text.Trim() -MaxMB ([double]$numSize.Value) -MaxEnt ([int]$numCnt.Value) -TargetCheck $chkJs.Checked
                $script:LastOut = $res.OutDir; $btnOpen.Enabled = $true
                if (-not $res.Passed) { $allOk = $false }
            }
            if ($files.Count -gt 1) { Write-Log 'Note: each file was split on its own. Links BETWEEN the input files are not managed - run Analyze on them together to see any.' 'WARN' }
            $msg = if ($allOk) { "Split complete and verified.`n`nImport the parts in the order listed in 00_IMPORT_ORDER.txt." } else { 'Split finished but verification FAILED. Do not import - see the log and 00_IMPORT_ORDER.txt.' }
            [void][System.Windows.Forms.MessageBox]::Show($form, $msg, 'Package splitter', [System.Windows.Forms.MessageBoxButtons]::OK, $(if ($allOk) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }))
        } catch { Write-Log $_.Exception.Message 'ERROR'; [void][System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Split failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
        finally { & $setBusy $false }
    })

    [void]$form.ShowDialog()
    $script:LogBox = $null
}

# ----------------------------------------------------------------------------------------------
#  Entry point
# ----------------------------------------------------------------------------------------------
$all = @(); if ($InputFile) { $all += $InputFile }; if ($MoreFiles) { $all += $MoreFiles }
$all = @($all | Where-Object { $_ })

try {
    if ($Gui -or $all.Count -eq 0) {
        Show-Gui $all
        return
    }
    if ($Analyze) {
        [void](Invoke-Analyze -Paths $all -MaxMB $MaxSizeMB -MaxEnt $MaxEntities)
        exit 0
    }
    if ($all.Count -gt 1 -and $OutputFolder) { throw 'With several input files, omit -OutputFolder: each file gets its own "<name>_split" folder.' }
    $ok = $true
    foreach ($f in $all) {
        $r = Invoke-Split -Path $f -OutDir $OutputFolder -MaxMB $MaxSizeMB -MaxEnt $MaxEntities -TargetCheck (-not $NoTargetCheck)
        if (-not $r.Passed) { $ok = $false }
    }
    if ($ok) { exit 0 } else { exit 2 }
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    Write-Log ('Where: ' + ((@($_.ScriptStackTrace -split "`n") | Select-Object -First 3) -join ' <- ')) 'ERROR'
    if ($Gui -or $all.Count -eq 0) {
        try { Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Package splitter', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) } catch { }
    }
    exit 1
}
