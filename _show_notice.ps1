# Called by _auto_git_pull.bat after a successful "git pull" (either the
# first attempt or a successful retry via _pull_retry.ps1). If an admin has
# published a new NOTICE.txt (edited -> committed -> PR -> merged -> pulled
# here via git), shows it once per team member and waits for Enter so they
# actually read it, instead of letting the window scroll past it.
#
# Comparison uses ONLY the git commit hash of NOTICE.txt (plain ASCII hex),
# never the Korean notice text itself - a PowerShell string comparison
# (-notcontains) on Korean text was found to misbehave elsewhere in this
# project (see docs/legacy/cp949_조사결과_및_결정.md), so this sidesteps
# that whole class of bug by design.
#
# Any failure here (missing file, git error, etc.) is swallowed silently so
# it can never block the actual work (python report generation).
#
# Also syncs [Console]::OutputEncoding to the console's real active code page
# before writing any Korean text below -- same fix as _pull_retry.ps1, same
# root cause (Windows PowerShell 5.1 does not reliably auto-match this when
# launched via "powershell -File" from a .bat, which garbled Korean output on
# a team member's PC even though this file is saved as UTF-8).

try {
    try {
        Add-Type -Name NativeConsole -Namespace Amd -MemberDefinition @'
[DllImport("kernel32.dll")]
public static extern uint GetConsoleOutputCP();
'@
        $consoleCp = [Amd.NativeConsole]::GetConsoleOutputCP()
        if ($consoleCp -gt 0) {
            [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding([int]$consoleCp)
        }
    } catch {
    }

    $repoDir = $env:REPO_DIR
    $noticePath = Join-Path $repoDir "NOTICE.txt"

    if (-not (Test-Path $noticePath)) { exit 0 }

    # -Encoding UTF8 is required here: NOTICE.txt is plain UTF-8 without a BOM
    # (e.g. when edited with Notepad or written by common tooling), and
    # Windows PowerShell 5.1's Get-Content falls back to the system ANSI code
    # page (CP949 on Korean Windows) whenever no BOM is present to auto-detect
    # from, silently corrupting every Korean character read from the file --
    # confirmed by reproducing the exact garbled output this way. This is a
    # different bug from the console output encoding fix in _pull_retry.ps1:
    # this one corrupts the string in memory before Write-Host ever runs, so
    # no console-side fix could have caught it.
    $content = Get-Content -Path $noticePath -Raw -Encoding UTF8 -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($content)) { exit 0 }

    $hash = (git -C "$repoDir." log -1 --format=%H -- NOTICE.txt 2>$null)
    if ([string]::IsNullOrWhiteSpace($hash)) { exit 0 }
    $hash = $hash.Trim()

    $stateDir = Join-Path $env:APPDATA "AMD"
    $stateFile = Join-Path $stateDir ".last_seen_notice"

    $seenHash = $null
    if (Test-Path $stateFile) {
        $seenHash = (Get-Content -Path $stateFile -Raw -ErrorAction Stop).Trim()
    }

    if ($seenHash -eq $hash) { exit 0 }

    Write-Host ""
    Write-Host "====================================================="
    Write-Host $content
    Write-Host "====================================================="
    Read-Host "계속하려면 Enter를 누르세요"

    if (-not (Test-Path $stateDir)) {
        New-Item -ItemType Directory -Path $stateDir -Force | Out-Null
    }
    Set-Content -Path $stateFile -Value $hash -Encoding ascii

    exit 0
}
catch {
    exit 0
}
