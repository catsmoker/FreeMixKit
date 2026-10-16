# ==============================================================================
# NAVIGATION AND CONTENT CARDS
# ==============================================================================
# Every card is generated from $ModuleMeta, so the registry stays the single
# source of truth for what appears in the UI.
#
# A page is a stack of sections. Each section is a heading plus a body panel,
# and the body is either
#
#   * a WrapPanel of tiles  -- categories in $FmkGridCategories, where each
#     entry is a standalone application and a logo reads better than a row of
#     text; or
#   * a StackPanel of list cards -- everything else, where a one-line
#     description next to the action is the right shape.
#
# Both bodies hold children carrying the module key in Uid and the search
# haystack in Tag (a Border for list rows, a Button for tiles), so filtering
# treats them identically.
#
# Click handlers read their inputs from the sender ($this) and from
# $global:sync. They deliberately never close over function-locals, because
# PowerShell scriptblocks do not capture the scope they were created in -- a
# `$key` referenced inside Add_Click would resolve to nothing at click time.
#
# The StackPanel holding a page's cards is called $cardPanel, not $host. "Host"
# is the obvious name for it and that is exactly the problem: variable names are
# case-insensitive, so $host is $Host, the read-only automatic variable, and the
# assignment throws at runtime. Compile.ps1 now rejects assignments to automatic
# variables outright.

# Sidebar order. Each name must match a Category in registry.ps1 and a
# x:Name="page_<name>" element in ui.xaml; Compile.ps1 enforces this.
$FmkCategories = @('System', 'Utilities', 'Software', 'Activation', 'About')

# Categories drawn as a logo grid. Software is where the standalone applications
# live, which is exactly the set a user browses by name and icon; the other
# categories are task lists ("clean this", "fix that") where a compact row per
# entry is easier to scan.
$FmkGridCategories = @('Software')

# Tile metrics. Every tile sets its own Width/Height to these, so all tiles come
# out the same size, the columns line up regardless of how long a name or
# description is, and the WrapPanel reflows to the page width on its own (the
# page's ScrollViewer has horizontal scrolling disabled, which is what makes the
# panel wrap instead of growing a scrollbar).
#
# The size sits on the tile rather than on the panel as WrapPanel.ItemWidth /
# ItemHeight, and that is load-bearing. ItemWidth/ItemHeight are applied to every
# child unconditionally, Collapsed ones included: a tile the search has filtered
# out keeps a full-size empty slot, so the survivors stay in their original
# columns and the grid fills up with holes. A Collapsed Border measures to
# nothing, so a tile carrying its own size drops out of the layout entirely and
# the rest close up behind it. Never set ItemWidth/ItemHeight here.
#
# The margin lies outside Width/Height: it is the gutter between neighbours and
# at the panel edges.
#
# Compact launcher sizing: the tile is a single clickable surface with no inner
# button, so the height budget covers the logo, the name, and the description
# only. The mark is deliberately large against the tile (26px art in a 38px box
# on a 152px tile) so icons stay recognizable at five columns.
$FmkTileWidth = 152
$FmkTileHeight = 140
$FmkTileMargin = 5
$FmkLogoSize = 26

# Content height budget inside a tile, in device-independent pixels, with the
# TileButton template's 8px padding on each side already subtracted -- the tile
# is 140 tall, so 124 is available:
#
#     38  logo box
#      6  gap
#     32  name          (two lines at 12px)
#      4  gap
#     32  description   (two lines at 11px)
#    ---
#    112, against 124 available -- the remainder is slack for descenders and
#    the risk-badge overlay, which floats above the content and shifts nothing.
#
# A name longer than two lines or a description longer than two is clipped
# (with an ellipsis) rather than allowed to stretch the tile, which is what
# keeps the grid regular. Keep Software descriptions under about 48 characters
# and labels under about 24.
$FmkTileNameHeight = 32
$FmkTileDescriptionHeight = 32

function Get-FmkModuleKeys {
    <# Module keys for a category, in registration order. #>
    param([Parameter(Mandatory)][string]$Category)

    $keys = @()
    foreach ($key in $ModuleMeta.Keys) {
        if ($ModuleMeta[$key].Category -eq $Category) { $keys += $key }
    }
    return $keys
}

function Test-FmkGridCategory {
    param([Parameter(Mandatory)][string]$Category)
    return ($FmkGridCategories -contains $Category)
}

function Test-FmkLowContrast {
    <#
        True when two colors are too close together to tell apart.

        Used to decide whether a brand mark's own colour is legible against the
        card it is drawn on. The threshold is deliberately loose: a brand colour
        is a deliberate choice and worth keeping even when it is low contrast, so
        this only fires when a mark is effectively invisible against the card.
    #>
    param(
        [Parameter(Mandatory)][string]$Foreground,
        [Parameter(Mandatory)][string]$Background,
        [double]$Minimum = 1.8
    )

    $lumA = Get-FmkLuminance -Hex $Foreground
    $lumB = Get-FmkLuminance -Hex $Background
    $high = [Math]::Max($lumA, $lumB)
    $low = [Math]::Min($lumA, $lumB)
    return ((($high + 0.05) / ($low + 0.05)) -lt $Minimum)
}

function Get-FmkLuminance {
    <# WCAG relative luminance of a #rrggbb color, 0 (black) to 1 (white). #>
    param([Parameter(Mandatory)][string]$Hex)

    $hexValue = $Hex.TrimStart('#')
    if ($hexValue.Length -ne 6) { return 0 }

    $linear = foreach ($offset in 0, 2, 4) {
        $channel = [Convert]::ToInt32($hexValue.Substring($offset, 2), 16) / 255
        if ($channel -le 0.03928) { $channel / 12.92 }
        else { [Math]::Pow((($channel + 0.055) / 1.055), 2.4) }
    }

    return (0.2126 * $linear[0]) + (0.7152 * $linear[1]) + (0.0722 * $linear[2])
}

function Get-FmkInitials {
    <# Fallback mark for a module with no logo: initials on a neutral tile. #>
    param([Parameter(Mandatory)][string]$Text)

    $words = @([regex]::Matches($Text, '[A-Za-z0-9]\S*') | ForEach-Object { $_.Value })
    if ($words.Count -eq 0) { return '?' }
    if ($words.Count -eq 1) { return $words[0].Substring(0, [Math]::Min(2, $words[0].Length)).ToUpperInvariant() }
    return (($words[0].Substring(0, 1)) + ($words[1].Substring(0, 1))).ToUpperInvariant()
}

function New-FmkLogoPath {
    <#
        One vector mark at tile-icon size. $null when the data does not parse,
        so a corrupt entry degrades to the monogram path instead of throwing
        during card construction.
    #>
    param(
        [Parameter(Mandatory)][string]$Data,
        [Parameter(Mandatory)]$Sync
    )

    try {
        $path = New-Object System.Windows.Shapes.Path
        # "F1 " is the path mini-language's fill-rule prefix (non-zero). SVG
        # defaults to non-zero and the marks are authored for it, but WPF's
        # parser defaults to even-odd -- which punches holes through the parts of
        # a mark that are meant to be solid, such as Node.js's hexagon.
        $path.Data = [System.Windows.Media.Geometry]::Parse("F1 $Data")
        # Uniform stretch scales each mark's own coordinate space to the box, so
        # an icon authored on a 24x24 grid and one on a 128x128 grid come out the
        # same size.
        $path.Stretch = [System.Windows.Media.Stretch]::Uniform
        $path.Width = $FmkLogoSize
        $path.Height = $FmkLogoSize
        $path.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
        $path.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        return $path
    }
    catch {
        return $null
    }
}

function New-FmkLogoBox {
    <#
        The square mark at the top of a tile.

        Four cases, in order: a vendored brand mark drawn as vector path data;
        a vendored raster icon (base64 PNG from the publisher's own repo) drawn
        as an Image with aspect preserved; a shape set (one Path per element,
        for marks that are not a single flat shape); and a monogram where no
        trustworthy mark exists. A module with no entry in $FmkLogos at all
        falls back to the label's initials, so a new registry entry -- or a
        missing asset file -- never breaks the grid.
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Meta,
        [Parameter(Mandatory)]$Sync
    )

    $box = New-Object System.Windows.Controls.Border
    $box.Style = $Sync.FmkWindow.FindResource('LogoBox')

    $entry = $null
    $logoKey = [string]$Meta.Logo
    if ($logoKey -and $FmkLogos.ContainsKey($logoKey)) { $entry = $FmkLogos[$logoKey] }

    if ($entry -and $entry.ContainsKey('Png')) {
        # Raster icons are vendored base64 at build time, so the tile draws
        # with no network access and no image files, exactly like the marks.
        try {
            $bytes = [Convert]::FromBase64String([string]$entry.Png)
            $stream = New-Object System.IO.MemoryStream($bytes, $false)
            try {
                $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                $bmp.BeginInit()
                $bmp.StreamSource = $stream
                $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $bmp.EndInit()
            }
            finally {
                $stream.Close()
            }
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $bmp
            # Uniform keeps the aspect and the transparency; a wide banner
            # (SpotX) letterboxes instead of stretching.
            $img.Stretch = [System.Windows.Media.Stretch]::Uniform
            $img.Width = $FmkLogoSize
            $img.Height = $FmkLogoSize
            $img.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
            $img.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($img, [System.Windows.Media.BitmapScalingMode]::HighQuality)

            $box.Child = $img
            return $box
        }
        catch {
            Write-Log "Logo for '$Key' is corrupt ($($_.Exception.Message)); using initials." "Warn"
        }
    }

    if ($entry -and $entry.ContainsKey('Shapes')) {
        $layer = New-Object System.Windows.Controls.Grid
        foreach ($shape in @($entry.Shapes)) {
            $sub = New-FmkLogoPath -Data ([string]$shape.Data) -Sync $Sync
            if (-not $sub) { continue }
            if ($shape.ContainsKey('Fill') -and $shape.Fill) {
                $sub.Fill = (New-Object System.Windows.Media.BrushConverter).ConvertFromString([string]$shape.Fill)
            }
            else {
                $sub.Fill = [System.Windows.Media.Brushes]::Transparent
            }
            if ($shape.ContainsKey('Stroke') -and $shape.Stroke) {
                $sub.Stroke = (New-Object System.Windows.Media.BrushConverter).ConvertFromString([string]$shape.Stroke)
                $sub.StrokeThickness = [double]$shape.StrokeThickness
            }
            [void]$layer.Children.Add($sub)
        }
        if ($layer.Children.Count -gt 0) {
            $box.Child = $layer
            return $box
        }
    }

    if ($entry -and $entry.ContainsKey('Data')) {
        $fill = [string]$entry.Color

        # A mark drawn in the card's own colour is invisible. Substituting the
        # theme foreground rather than a hardcoded white is what keeps the
        # fallback correct on the light theme too, where the foreground is
        # dark and a dark mark must stay dark.
        $palette = $FmkThemes[$Sync.Theme]
        if ($palette -and (Test-FmkLowContrast -Foreground $fill -Background $palette['FmkCard'])) {
            $fill = $palette['FmkFg']
        }

        $path = New-FmkLogoPath -Data ([string]$entry.Data) -Sync $Sync
        if ($path) {
            $path.Fill = (New-Object System.Windows.Media.BrushConverter).ConvertFromString($fill)
            $box.Child = $path
            return $box
        }
    }

    $text = New-Object System.Windows.Controls.TextBlock
    $text.Text = if ($entry -and $entry.ContainsKey('Mono')) {
        [string]$entry.Mono
    }
    else {
        Get-FmkInitials -Text $Meta.Label
    }
    $text.FontSize = 15
    $text.FontWeight = [System.Windows.FontWeights]::Bold
    $text.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $text.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    if ($Sync.FmkWindow.Resources.Contains('FmkLabel')) {
        $text.Foreground = $Sync.FmkWindow.Resources['FmkLabel']
    }

    $box.Child = $text
    return $box
}

function New-FmkRiskBadge {
    <#
        Low, Medium, and High each get a badge in their own theme color; None
        (and anything unknown) gets nothing. Visibility goes through
        Test-FmkRiskBadgeVisible so the rule lives in one testable place.
    #>
    param([Parameter(Mandatory)]$Meta, [Parameter(Mandatory)]$Sync)

    if (-not (Test-FmkRiskBadgeVisible -Risk ([string]$Meta.Risk))) { return $null }

    $resources = $Sync.FmkWindow.Resources
    $badge = New-Object System.Windows.Controls.Border
    $badge.Style = $Sync.FmkWindow.FindResource('Pill')

    $riskKey = if ($Meta.Risk -eq 'High') { 'FmkRiskHigh' }
    elseif ($Meta.Risk -eq 'Medium') { 'FmkRiskMed' }
    else { 'FmkRiskLow' }
    if ($resources.Contains($riskKey)) { $badge.Background = $resources[$riskKey] }

    $badgeText = New-Object System.Windows.Controls.TextBlock
    $badgeText.Text = $Meta.Risk
    $badgeText.FontSize = 10
    $badgeText.FontWeight = [System.Windows.FontWeights]::SemiBold
    # FmkBg rather than white: the risk fills are bright in the dark theme and
    # dark in the light one, and FmkBg is the inverse of both, so one lookup
    # gives readable text either way. Hardcoded white was unreadable on the
    # dark theme's amber (#FFC83D).
    if ($resources.Contains('FmkBg')) { $badgeText.Foreground = $resources['FmkBg'] }
    $badge.Child = $badgeText

    # Tooltip carries the full breakdown (dimensions, evidence, review date).
    # Get-FmkRiskTooltip lives in 08-risk.ps1; resolved at click time, so the
    # file order does not matter.
    try {
        $tip = Get-FmkRiskTooltip -Meta $Meta
        if ($tip) { $badge.ToolTip = $tip }
    }
    catch { }

    return $badge
}

function New-FmkActionButton {
    <#
        The action button on a list-card row. (Grid tiles are buttons
        themselves and carry no inner button.) Uid holds the module key and the
        click handler lives here, so every row dispatches identically; Tag
        is left free for the search haystack on the card itself.

        Every button is added to $Sync.RunButtons, which is what Set-FmkBusy
        walks to disable the whole UI while a module is in flight -- that is the
        first of the two things stopping a repeated click from starting a second
        install. The second is the re-entrancy guard in Invoke-FmkModule.
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Meta,
        [Parameter(Mandatory)]$Sync,
        [Parameter(Mandatory)][string]$Caption
    )

    $run = New-Object System.Windows.Controls.Button
    $run.Content = $Caption
    $run.Style = $Sync.FmkWindow.FindResource('RunButton')
    $run.ToolTip = "$Caption $($Meta.Label)"
    $run.Uid = $Key
    # [void]: this must not put anything on the pipeline ahead of the caller's
    # return. Both ArrayList.Add and Add_Click return values.
    [void]$run.Add_Click({
            # Named in the log so it is always clear which component a click is
            # installing, including when the click is refused as a duplicate.
            # Muted: dispatch chatter, distinct from the operation's own output.
            Write-Log "Requested: $($this.ToolTip)" "Muted"
            [void](Invoke-FmkModule -ActionKey $this.Uid -Sync $global:sync)
        })
    [void]$Sync.RunButtons.Add($run)

    return $run
}

function New-FmkCard {
    <# A row card: title and description on the left, badge and action on the right. #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Sync
    )

    $meta = $ModuleMeta[$Key]
    $resources = $Sync.FmkWindow.Resources

    $card = New-Object System.Windows.Controls.Border
    $card.Style = $Sync.FmkWindow.FindResource('Card')
    $card.Uid = $Key
    # Search haystack. Lowercased once here so filtering stays cheap.
    $card.Tag = ("$($meta.Label) $($meta.Description)").ToLowerInvariant()

    $grid = New-Object System.Windows.Controls.Grid
    $colText = New-Object System.Windows.Controls.ColumnDefinition
    $colText.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $colActions = New-Object System.Windows.Controls.ColumnDefinition
    $colActions.Width = [System.Windows.GridLength]::Auto
    [void]$grid.ColumnDefinitions.Add($colText)
    [void]$grid.ColumnDefinitions.Add($colActions)

    # --- left: title + description -------------------------------------------
    $left = New-Object System.Windows.Controls.StackPanel

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = $meta.Label
    # 13 matches WinUtil's AppEntryFontSize (13.2); the base size is 12.
    $title.FontSize = 13
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.TextWrapping = [System.Windows.TextWrapping]::Wrap
    [void]$left.Children.Add($title)

    $desc = New-Object System.Windows.Controls.TextBlock
    $desc.Text = $meta.Description
    $desc.FontSize = 12
    $desc.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $desc.Margin = New-Object System.Windows.Thickness(0, 4, 0, 0)
    if ($resources.Contains('FmkFgMuted')) { $desc.Foreground = $resources['FmkFgMuted'] }
    [void]$left.Children.Add($desc)

    [System.Windows.Controls.Grid]::SetColumn($left, 0)
    [void]$grid.Children.Add($left)

    # --- right: risk badge + action ------------------------------------------
    $right = New-Object System.Windows.Controls.StackPanel
    $right.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $right.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
    $right.Margin = New-Object System.Windows.Thickness(16, 0, 0, 0)

    $badge = New-FmkRiskBadge -Meta $meta -Sync $Sync
    if ($badge) { [void]$right.Children.Add($badge) }

    [void]$right.Children.Add((New-FmkActionButton -Key $Key -Meta $meta -Sync $Sync -Caption $meta.ActionLabel))

    [System.Windows.Controls.Grid]::SetColumn($right, 1)
    [void]$grid.Children.Add($right)

    $card.Child = $grid
    return $card
}

function New-FmkTile {
    <#
        A grid tile: logo, name, description. The tile toggles selection -- a
        click selects, a second click deselects -- and never installs on its
        own. The shared Install Selected button runs the queue in selection
        order. No inner Install button, no checkbox; selection rides on
        Tag ("selected") so the TileButton style draws the colored rounded
        border, with the search haystack in DataContext (Tag is taken).

        Being a real Button buys the interaction states for free: hover and
        pressed come from the TileButton template, keyboard Enter/Space works
        natively, the tile participates in tab order with a visible focus ring
        (IsKeyboardFocused trigger), and the queue disables the action bar
        while running on top of its own re-entrancy guard.

        The risk badge sits in a Grid overlay rather than in the flow, so it can
        appear on a Medium or High tile without shifting the name down and making
        that one tile's contents sit differently from its neighbours'.
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Sync
    )

    $meta = $ModuleMeta[$Key]
    $resources = $Sync.FmkWindow.Resources

    $tile = New-Object System.Windows.Controls.Button
    $tile.Style = $Sync.FmkWindow.FindResource('TileButton')
    # Focus adorner lives OUTSIDE the tile (TileFocusVisual): keyboard focus
    # must never recolor the tile border itself, or a deselected-but-focused
    # tile renders a stale highlight. See the TileButton trigger note in XAML.
    $tile.FocusVisualStyle = $Sync.FmkWindow.FindResource('TileFocusVisual')
    # Local values, so they win over the style's setters; kept here rather than
    # in the XAML so the tile metrics live in one file. Fixed rather than
    # content-sized so the columns and rows line up, and set on the tile rather
    # than on the panel so that filtering a tile out actually removes its slot.
    $tile.Width = $FmkTileWidth
    $tile.Height = $FmkTileHeight
    $tile.Margin = New-Object System.Windows.Thickness($FmkTileMargin)
    $tile.Uid = $Key
    # Haystack lives in DataContext so Tag stays free for the "selected" marker
    # read by the TileButton style. Lowercased once here so filtering stays cheap.
    $tile.DataContext = ("$($meta.Label) $($meta.Description)").ToLowerInvariant()
    $tile.Tag = $null
    $tile.Cursor = [System.Windows.Input.Cursors]::Hand
    Update-FmkTileTooltip -Tile $tile -Meta $meta
    # [void]: Add_Click returns a value that must not reach the pipeline ahead
    # of the caller's return.
    [void]$tile.Add_Click({
            [void](Toggle-FmkSoftwareSelection -ActionKey $this.Uid -Sync $global:sync)
        })
    # Tiles are NOT in RunButtons: Set-FmkBusy disables single-run actions, and
    # the queue manages its own busy state plus the action bar. Click handlers
    # read $this/$global:sync, never function-locals, because PowerShell
    # scriptblocks do not capture the scope they were made in.

    $root = New-Object System.Windows.Controls.Grid

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.VerticalAlignment = [System.Windows.VerticalAlignment]::Top

    [void]$stack.Children.Add((New-FmkLogoBox -Key $Key -Meta $meta -Sync $Sync))

    $name = New-Object System.Windows.Controls.TextBlock
    $name.Text = $meta.Label
    $name.FontSize = 12
    $name.FontWeight = [System.Windows.FontWeights]::SemiBold
    $name.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $name.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
    $name.TextAlignment = [System.Windows.TextAlignment]::Center
    $name.Height = $FmkTileNameHeight
    $name.Margin = New-Object System.Windows.Thickness(0, 6, 0, 0)
    [void]$stack.Children.Add($name)

    $desc = New-Object System.Windows.Controls.TextBlock
    $desc.Text = $meta.Description
    $desc.FontSize = 11
    $desc.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $desc.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
    $desc.TextAlignment = [System.Windows.TextAlignment]::Center
    $desc.Height = $FmkTileDescriptionHeight
    $desc.Margin = New-Object System.Windows.Thickness(0, 4, 0, 0)
    if ($resources.Contains('FmkFgMuted')) { $desc.Foreground = $resources['FmkFgMuted'] }
    [void]$stack.Children.Add($desc)

    [void]$root.Children.Add($stack)

    $badge = New-FmkRiskBadge -Meta $meta -Sync $Sync
    if ($badge) {
        $badge.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        $badge.VerticalAlignment = [System.Windows.VerticalAlignment]::Top
        [void]$root.Children.Add($badge)
    }

    $tile.Content = $root
    return $tile
}

function New-FmkSectionBody {
    <#
        Wraps one group's entries in the panel its category calls for. Returns the
        panel; the caller pairs it with a heading.
    #>
    param([Parameter(Mandatory)][bool]$Grid)

    if ($Grid) {
        # No ItemWidth/ItemHeight: those would apply to Collapsed tiles too and
        # keep an empty slot for every filtered-out card, so a search would leave
        # the survivors in their original columns among holes. The tiles carry
        # their own size instead, which is what lets the panel close up the gaps
        # and reflow the filtered set from the top-left. See the metrics above.
        return (New-Object System.Windows.Controls.WrapPanel)
    }

    return (New-Object System.Windows.Controls.StackPanel)
}

function Get-FmkSoftwareGroups {
    <# Group names on the Software page, in first-appearance order. #>
    $groups = @()
    foreach ($key in (Get-FmkModuleKeys -Category 'Software')) {
        $group = $ModuleMeta[$key].Group
        if ($group -and ($groups -notcontains $group)) { $groups += $group }
    }
    return $groups
}

function Build-FmkSoftwareChips {
    <#
        The All + group chip bar above the Software grid. Chips are Buttons so
        they tab-focus and trigger on Enter/Space natively; selection rides on
        Tag ("active"), the same convention as the sidebar.
    #>
    param([Parameter(Mandatory)]$Sync)

    if (-not $Sync.SoftwareChips) { return }
    $Sync.SoftwareChips.Children.Clear()

    $names = @('All') + @(Get-FmkSoftwareGroups)
    foreach ($name in $names) {
        $chip = New-Object System.Windows.Controls.Button
        $chip.Content = $name
        $chip.Style = $Sync.FmkWindow.FindResource('ChipButton')
        # Empty Uid means All; Set-FmkSoftwareGroup normalises it to $null.
        $chip.Uid = if ($name -eq 'All') { '' } else { $name }
        $chip.Tag = if ((-not $Sync.SoftwareGroup -and $name -eq 'All') -or
            ($Sync.SoftwareGroup -and $Sync.SoftwareGroup -eq $name)) { 'active' } else { $null }
        [void]$chip.Add_Click({
                Set-FmkSoftwareGroup -Group $this.Uid -Sync $global:sync
            })
        [void]$Sync.SoftwareChips.Children.Add($chip)
    }
}

function Set-FmkSoftwareGroup {
    <#
        Selects a Software chip (empty means All), refreshes the chip states,
        and re-applies the filter. Search text is untouched, so chips and the
        search box combine: a card must match both.
    #>
    param(
        $Group = $null,
        [Parameter(Mandatory)]$Sync
    )

    if ([string]::IsNullOrEmpty($Group)) { $Group = $null }
    $Sync.SoftwareGroup = $Group

    foreach ($chip in $Sync.SoftwareChips.Children) {
        $chipGroup = $chip.Uid
        if ([string]::IsNullOrEmpty($chipGroup)) { $chipGroup = $null }
        $chip.Tag = if ($chipGroup -eq $Group) { 'active' } else { $null }
    }

    Apply-FmkFilter -Sync $Sync
}

function Build-FmkCards {
    param([Parameter(Mandatory)]$Sync)

    $Sync.CardSections = @{}

    foreach ($category in $FmkCategories) {
        $cardPanel = $Sync.FmkWindow.FindName("cards_$category")
        if (-not $cardPanel) { continue }

        $cardPanel.Children.Clear()
        $sections = New-Object System.Collections.ArrayList
        $Sync.CardSections[$category] = $sections

        # Short risk intro on pages that carry badges. Hovering a badge shows
        # the full breakdown; the header and Software action bar link the
        # complete Low/Medium/High criteria.
        if (-not (Test-FmkGridCategory -Category $category)) {
            $hasRated = $false
            foreach ($k in (Get-FmkModuleKeys -Category $category)) {
                $r = [string]$ModuleMeta[$k].Risk
                if ($r -eq 'Medium' -or $r -eq 'High') { $hasRated = $true; break }
            }
            if ($hasRated) {
                $intro = New-Object System.Windows.Controls.TextBlock
                $intro.Text = 'Risk badges mark Low, Medium, and High entries. Hover a badge for the reason; see Risk ratings? for the criteria.'
                $intro.FontSize = 11
                $intro.TextWrapping = [System.Windows.TextWrapping]::Wrap
                $intro.Margin = New-Object System.Windows.Thickness(2, 0, 0, 8)
                try {
                    if ($Sync.FmkWindow.Resources.Contains('FmkFgMuted')) {
                        $intro.Foreground = $Sync.FmkWindow.Resources['FmkFgMuted']
                    }
                }
                catch { }
                [void]$cardPanel.Children.Add($intro)
            }
        }

        # About page lead-in: product description, implemented features, and
        # support actions. The registry cards below it (repository, website,
        # shortcut) are preserved unchanged.
        if ($category -eq 'About') {
            [void]$cardPanel.Children.Add((New-FmkAboutPanel -Sync $Sync))
        }

        $isGrid = Test-FmkGridCategory -Category $category
        $currentGroup = $null
        $body = $null

        foreach ($key in (Get-FmkModuleKeys -Category $category)) {
            $meta = $ModuleMeta[$key]

            if ($meta.Group -and $meta.Group -ne $currentGroup) {
                $currentGroup = $meta.Group

                $heading = New-Object System.Windows.Controls.TextBlock
                $heading.Text = $currentGroup.ToUpperInvariant()
                $heading.Style = $Sync.FmkWindow.FindResource('SectionHeading')
                [void]$cardPanel.Children.Add($heading)

                $body = New-FmkSectionBody -Grid $isGrid
                [void]$cardPanel.Children.Add($body)

                # Kept so Apply-FmkFilter can hide a heading whose entries are
                # all filtered out -- walking the page's children instead would
                # mean re-deriving which heading belongs to which body.
                [void]$sections.Add(@{ Heading = $heading; Body = $body })
            }

            if (-not $body) { continue }

            $card = if ($isGrid) {
                New-FmkTile -Key $key -Sync $Sync
            }
            else {
                New-FmkCard -Key $key -Sync $Sync
            }
            [void]$body.Children.Add($card)
        }
    }

    Build-FmkSoftwareChips -Sync $Sync
}

function Build-FmkNav {
    param([Parameter(Mandatory)]$Sync)

    $Sync.NavPanel.Children.Clear()

    foreach ($category in $FmkCategories) {
        # @() because Get-FmkModuleKeys unrolls on return: one key comes back as
        # a bare string, none comes back as $null.
        if (@(Get-FmkModuleKeys -Category $category).Count -eq 0) { continue }

        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = $category
        $btn.Style = $Sync.FmkWindow.FindResource('NavButton')
        # Uid = which page to show; Tag = selection marker read by the style.
        $btn.Uid = $category
        $btn.Add_Click({
                Set-FmkPage -Category $this.Uid -Sync $global:sync
            })

        [void]$Sync.NavPanel.Children.Add($btn)
    }
}

function Set-FmkPage {
    param(
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)]$Sync
    )

    foreach ($name in $FmkCategories) {
        $page = $Sync.FmkWindow.FindName("page_$name")
        if ($page) {
            $page.Visibility = if ($name -eq $Category) {
                [System.Windows.Visibility]::Visible
            }
            else {
                [System.Windows.Visibility]::Collapsed
            }
        }
    }

    foreach ($btn in $Sync.NavPanel.Children) {
        $btn.Tag = if ($btn.Uid -eq $Category) { 'active' } else { $null }
    }

    $Sync.ActiveCategory = $Category
    $Sync.SearchBox.Text = ''
    Apply-FmkFilter -Sync $Sync
}

function Apply-FmkFilter {
    <#
        Hides entries on the active page that do not match the search box and,
        on the Software page, the selected group chip. A card must match both.

        Hiding is all this does -- the panels reflow on their own. A Collapsed
        child is not measured, so the grid closes up around it and the matches
        move up to the first free slot without any element being recreated or
        reordered. That is also why the tiles carry their own size instead of the
        panel carrying an item size; see the tile metrics at the top of this file.

        Also resets the page's scroll offset when the query or the chip changes,
        so a filtered grid starts at the top rather than at whatever offset the
        unfiltered list was left at. The Software page scrolls an inner viewer
        (the chips stay put), so that viewer is preferred when it exists.
    #>
    param([Parameter(Mandatory)]$Sync)

    $category = $Sync.ActiveCategory
    if (-not $category) { return }

    $sections = $null
    if ($Sync.CardSections) { $sections = $Sync.CardSections[$category] }
    if (-not $sections) { return }

    $term = ''
    if ($Sync.SearchBox.Text) { $term = $Sync.SearchBox.Text.Trim().ToLowerInvariant() }

    $group = $null
    if ($Sync.SoftwareGroup -and (Test-FmkGridCategory -Category $category)) {
        $group = $Sync.SoftwareGroup
    }

    $visible = 0

    foreach ($section in $sections) {
        $hits = 0

        foreach ($card in $section.Body.Children) {
            $inGroup = (-not $group) -or ($ModuleMeta[$card.Uid].Group -eq $group)
            # Tiles keep the haystack in DataContext (Tag is the selection
            # marker); list cards keep it in Tag. Filtering only flips
            # Visibility, so selection survives search and chip changes.
            $haystack = $null
            try {
                if ($card -is [System.Windows.Controls.Button]) { $haystack = [string]$card.DataContext }
                else { $haystack = [string]$card.Tag }
            }
            catch { $haystack = $null }
            $match = $inGroup -and (($term -eq '') -or ($haystack -and $haystack.Contains($term)))
            $card.Visibility = if ($match) {
                [System.Windows.Visibility]::Visible
            }
            else {
                [System.Windows.Visibility]::Collapsed
            }
            if ($match) { $hits++ }
        }

        # A heading with nothing under it is noise, so it goes with its entries.
        # Collapsed rather than hidden so it takes no space in the StackPanel.
        $section.Heading.Visibility = if ($hits -gt 0) {
            [System.Windows.Visibility]::Visible
        }
        else {
            [System.Windows.Visibility]::Collapsed
        }

        $visible += $hits
    }

    if ((Test-FmkGridCategory -Category $category) -and $Sync.SoftwareCount) {
        $total = @(Get-FmkModuleKeys -Category $category).Count
        if (($term -eq '') -and (-not $group)) {
            $Sync.SoftwareCount.Text = "$total applications"
        }
        else {
            $Sync.SoftwareCount.Text = "$visible of $total applications"
        }
    }

    # Reset the scroll only when this page's filter state actually changed.
    # Comparing against the last applied state keeps a plain navigation click
    # from throwing away the offset a page was left at, while every keystroke,
    # every chip change -- and clearing the box -- starts the grid at the top.
    $terms = $Sync.FilterTerms
    if (-not $terms) {
        $terms = @{}
        $Sync.FilterTerms = $terms
    }

    $state = "$group`n$term"
    if (-not $terms.ContainsKey($category) -or $terms[$category] -ne $state) {
        $terms[$category] = $state
        $scroller = $Sync.FmkWindow.FindName("scroll_$category")
        if (-not $scroller -or -not $scroller.PSObject.Methods['ScrollToTop']) {
            $scroller = $Sync.FmkWindow.FindName("page_$category")
        }
        if ($scroller -and $scroller.PSObject.Methods['ScrollToTop']) {
            $scroller.ScrollToTop()
        }
    }

    if ($visible -eq 0) {
        # The grid pages list applications; the others list tasks.
        $Sync.EmptyPageNotice.Text = if (Test-FmkGridCategory -Category $category) {
            'No applications found'
        }
        else {
            'No modules match that filter.'
        }
        $Sync.EmptyPageNotice.Visibility = [System.Windows.Visibility]::Visible
    }
    else {
        $Sync.EmptyPageNotice.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# --- Software multi-select ----------------------------------------------------
# Selection order is the ArrayList order: entries are appended on select and
# removed on deselect, never sorted or re-derived. Filtering only flips
# Visibility, so the queue survives search and chip changes.

function Get-FmkSoftwareSelection {
    <# Ordered snapshot of selected Software keys. #>
    param([Parameter(Mandatory)]$Sync)
    if (-not $Sync.SoftwareSelection) { return @() }
    return @($Sync.SoftwareSelection)
}

function Update-FmkTileTooltip {
    param(
        [Parameter(Mandatory)]$Tile,
        [Parameter(Mandatory)]$Meta
    )
    $base = "Click to select $($Meta.Label)"
    try {
        if ($Tile.Tag -eq 'selected') { $base = "$($Meta.Label) (selected, click to deselect)" }
    }
    catch { }
    $riskTip = ''
    try { $riskTip = Get-FmkRiskTooltip -Meta $Meta } catch { $riskTip = '' }
    if ($riskTip) { $Tile.ToolTip = "$base`n$riskTip" }
    else { $Tile.ToolTip = $base }
}

function Update-FmkTileSelectionVisual {
    param(
        [Parameter(Mandatory)][string]$ActionKey,
        [Parameter(Mandatory)]$Sync
    )
    if (-not $Sync.CardSections) { return }
    $sections = $Sync.CardSections['Software']
    if (-not $sections) { return }
    $selected = $false
    if ($Sync.SoftwareSelection -and $Sync.SoftwareSelection.Contains($ActionKey)) { $selected = $true }
    foreach ($section in $sections) {
        foreach ($card in $section.Body.Children) {
            if ($card.Uid -ne $ActionKey) { continue }
            $card.Tag = if ($selected) { 'selected' } else { $null }
            try { Update-FmkTileTooltip -Tile $card -Meta $ModuleMeta[$ActionKey] } catch { }
        }
    }
}

function Update-FmkInstallButton {
    param([Parameter(Mandatory)]$Sync)
    $count = 0
    if ($Sync.SoftwareSelection) { $count = $Sync.SoftwareSelection.Count }
    if ($Sync.InstallSelectedButton) {
        $Sync.InstallSelectedButton.Content = if ($count -eq 0) { 'Install Selected (0)' } else { "Install Selected ($count)" }
        # Disabled when nothing is selected or a queue/single module is running.
        $busy = [bool]$Sync.Busy -or [bool]$Sync.QueueRunning
        $Sync.InstallSelectedButton.IsEnabled = ($count -gt 0) -and (-not $busy)
    }
    if ($Sync.ClearSelectionButton) {
        $busy = [bool]$Sync.Busy -or [bool]$Sync.QueueRunning
        $Sync.ClearSelectionButton.IsEnabled = ($count -gt 0) -and (-not $busy)
    }
    if ($Sync.SelectionStatus -and -not [bool]$Sync.QueueRunning) {
        if ($count -eq 0) {
            $Sync.SelectionStatus.Text = 'Click cards to select apps, then Install Selected.'
        }
        else {
            $names = @()
            foreach ($k in @(Get-FmkSoftwareSelection -Sync $Sync)) {
                if ($ModuleMeta.Contains($k)) { $names += $ModuleMeta[$k].Label }
            }
            $Sync.SelectionStatus.Text = "Selected ($count): " + ($names -join ', ')
        }
    }
}

function Toggle-FmkSoftwareSelection {
    param(
        [Parameter(Mandatory)][string]$ActionKey,
        [Parameter(Mandatory)]$Sync
    )
    if ([bool]$Sync.QueueRunning) {
        Write-Log 'A queue is already running; selection is locked until it finishes.' 'Warn'
        return $null
    }
    if ([bool]$Sync.Busy) {
        Write-Log "Already running: $($Sync.BusyLabel). Selection change ignored." 'Warn'
        return $null
    }
    if (-not $Sync.SoftwareSelection) {
        $Sync.SoftwareSelection = New-Object System.Collections.ArrayList
    }
    if ($Sync.SoftwareSelection.Contains($ActionKey)) {
        [void]$Sync.SoftwareSelection.Remove($ActionKey)
        Write-Log "Deselected: $($ModuleMeta[$ActionKey].Label)" 'Muted'
    }
    else {
        [void]$Sync.SoftwareSelection.Add($ActionKey)
        $pos = $Sync.SoftwareSelection.Count
        Write-Log "Selected (#$pos): $($ModuleMeta[$ActionKey].Label)" 'Muted'
    }
    Update-FmkTileSelectionVisual -ActionKey $ActionKey -Sync $Sync
    Update-FmkInstallButton -Sync $Sync
    return $null
}

function Clear-FmkSoftwareSelection {
    param([Parameter(Mandatory)]$Sync)
    if ([bool]$Sync.QueueRunning) { return }
    $keys = @(Get-FmkSoftwareSelection -Sync $Sync)
    if ($Sync.SoftwareSelection) { [void]$Sync.SoftwareSelection.Clear() }
    foreach ($k in $keys) { Update-FmkTileSelectionVisual -ActionKey $k -Sync $Sync }
    Update-FmkInstallButton -Sync $Sync
}
