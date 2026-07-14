# Called only by _auto_git_pull.bat when the first "git pull --ff-only" fails.
# Most failures on the NAS (network path) are git's "dubious ownership" safety
# check (CVE-2022-24765): the SMB share reports an owner SID that does not
# match the logged-in user's SID, so git refuses to operate until this exact
# path is explicitly marked as safe. See docs/legacy/cp949_조사결과_및_결정.md
# for the full investigation.
#
# This script adds ONLY this one NAS repository path to safe.directory
# (never "*", to avoid trusting every repository on the machine), then
# retries the pull once. If it still fails for some other reason, it shows
# a Korean notice and lets the caller continue with the existing code.
#
# git itself may not be installed on a team member's PC. In that case every
# "git ..." call below would fail with PowerShell's own "term not recognized"
# error, which is a different failure than "pull failed" and was leaking to
# the console as a confusing raw error instead of a clear Korean notice. This
# check must run before any other git invocation in this script.
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host '[알림] Git이 설치되어 있지 않아 최신 코드를 받지 못했습니다. IT 담당자에게 Git 설치를 요청해 주세요. 기존 코드로 계속 진행합니다.'
    exit 1
}

$repoDir = $env:REPO_DIR
$safeDirValue = '%(prefix)///nas/CN_AMD/판매수수료 관련자료/salary_system_claude'

# Ask git itself whether this exact value is already registered, instead of
# comparing strings in PowerShell (a plain -notcontains comparison here was
# found to add a duplicate entry on Korean text, likely a Unicode
# normalization mismatch between what PowerShell captured and the literal
# stored value).
git config --global --fixed-value --get-all safe.directory $safeDirValue *> $null
if ($LASTEXITCODE -ne 0) {
    git config --global --add safe.directory $safeDirValue
}

git -C "$repoDir." pull --ff-only *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host '[알림] 최신 코드 확인 실패, 기존 코드로 진행합니다.'
    exit 1
}

exit 0
