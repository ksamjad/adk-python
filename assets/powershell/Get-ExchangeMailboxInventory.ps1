#requires -Version 5.1
<#
.SYNOPSIS
Generates inventory reports for on-premises Exchange mailboxes, including categorization and folder statistics.

.DESCRIPTION
This script collects Exchange mailbox information and produces two CSV reports: a mailbox summary and folder-level details.
It includes the primary SMTP address, additional proxy addresses, mailbox size, and item counts. The script should be run
from the Exchange Management Shell or a PowerShell session where the Exchange snap-ins are loaded.

.PARAMETER OutputPath
Destination folder for the generated CSV reports. The folder is created when it does not exist.

.PARAMETER IncludeArchiveStatistics
When specified, archive mailbox statistics are included in the mailbox summary report.

.PARAMETER FolderScope
Controls the folder scope used with Get-MailboxFolderStatistics. The default of "All" collects every folder. Other common
values are "Inbox", "NonIPM", and "Archive".

.OUTPUTS
Creates two CSV files:
- MailboxInventory.csv: High-level mailbox details.
- MailboxFolderInventory.csv: Folder-level statistics for each mailbox.

.EXAMPLE
PS> .\Get-ExchangeMailboxInventory.ps1 -OutputPath C:\Reports\Exchange -IncludeArchiveStatistics

.LINK
Get-Mailbox
Get-MailboxStatistics
Get-MailboxFolderStatistics
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath 'MailboxInventory'),

  [Parameter(Mandatory = $false)]
  [switch]$IncludeArchiveStatistics,

  [Parameter(Mandatory = $false)]
  [ValidateSet('All', 'Archive', 'Inbox', 'NonIPM')]
  [string]$FolderScope = 'All'
)

function Get-MailboxCategory {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$RecipientTypeDetails
  )

  switch ($RecipientTypeDetails) {
    'UserMailbox' { return 'User' }
    'LinkedMailbox' { return 'Linked' }
    'SharedMailbox' { return 'Shared' }
    'RoomMailbox' { return 'Room Resource' }
    'EquipmentMailbox' { return 'Equipment Resource' }
    'DiscoveryMailbox' { return 'Discovery' }
    'LegacyMailbox' { return 'Legacy' }
    default { return $RecipientTypeDetails }
  }
}

try {
  if (-not (Test-Path -Path $OutputPath)) {
    Write-Verbose "Creating output directory at '$OutputPath'."
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
  }

  $mailboxes = Get-Mailbox -ResultSize Unlimited | Sort-Object DisplayName
  if (-not $mailboxes) {
    Write-Warning 'No mailboxes were returned by Get-Mailbox.'
    return
  }

  $mailboxSummary = @()
  $folderInventory = @()

  foreach ($mailbox in $mailboxes) {
    Write-Verbose "Processing mailbox: $($mailbox.DisplayName)"

    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity
    if (-not $mailboxStats) {
      Write-Warning "Mailbox statistics not found for $($mailbox.Identity)."
      continue
    }

    $primarySmtp = $mailbox.PrimarySmtpAddress.ToString()
    $proxyAddresses = $mailbox.EmailAddresses |
      ForEach-Object { $_.ToString() } |
      Where-Object {
        $_ -like 'smtp:*' -and -not $_.StartsWith(
          "SMTP:$primarySmtp", [System.StringComparison]::InvariantCultureIgnoreCase)
      } |
      ForEach-Object { $_.Substring(5) }

    $category = Get-MailboxCategory -RecipientTypeDetails $mailbox.RecipientTypeDetails

    $summaryEntry = [PSCustomObject]@{
      DisplayName            = $mailbox.DisplayName
      Alias                  = $mailbox.Alias
      OrganizationalUnit     = $mailbox.OrganizationalUnit
      Category               = $category
      RecipientTypeDetails   = $mailbox.RecipientTypeDetails
      PrimarySmtpAddress     = $primarySmtp
      AdditionalEmailAddresses = ($proxyAddresses -join '; ')
      TotalItemSize          = $mailboxStats.TotalItemSize.ToString()
      ItemCount              = $mailboxStats.ItemCount
      LastLogonTime          = $mailboxStats.LastLogonTime
      LastLogoffTime         = $mailboxStats.LastLogoffTime
      Database               = $mailboxStats.Database
    }

    if ($IncludeArchiveStatistics.IsPresent -and $mailboxStats.ArchiveDatabase) {
      $archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive
      if ($archiveStats) {
        $summaryEntry | Add-Member -NotePropertyName 'ArchiveTotalItemSize' -NotePropertyValue $archiveStats.TotalItemSize.ToString()
        $summaryEntry | Add-Member -NotePropertyName 'ArchiveItemCount' -NotePropertyValue $archiveStats.ItemCount
      }
    }

    $mailboxSummary += $summaryEntry

    $folderStats = Get-MailboxFolderStatistics -Identity $mailbox.Identity -FolderScope $FolderScope |
      Select-Object @{Name = 'MailboxIdentity'; Expression = { $mailbox.PrimarySmtpAddress.ToString() }},
        DisplayName,
        FolderPath,
        ItemsInFolder,
        UnreadItemCount,
        @{Name = 'FolderSize'; Expression = { $_.FolderSize.ToString() }},
        @{Name = 'FolderAndSubfolderSize'; Expression = { $_.FolderAndSubfolderSize.ToString() }},
        @{Name = 'FolderType'; Expression = { $_.FolderType }}

    if ($folderStats) {
      $folderInventory += $folderStats
    }
  }

  $summaryPath = Join-Path -Path $OutputPath -ChildPath 'MailboxInventory.csv'
  $folderPath = Join-Path -Path $OutputPath -ChildPath 'MailboxFolderInventory.csv'

  $mailboxSummary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
  Write-Verbose "Mailbox summary exported to $summaryPath"

  $folderInventory | Export-Csv -Path $folderPath -NoTypeInformation -Encoding UTF8
  Write-Verbose "Folder inventory exported to $folderPath"

  Write-Host "Mailbox inventory reports created:" -ForegroundColor Green
  Write-Host "  Summary: $summaryPath"
  Write-Host "  Folder details: $folderPath"
}
catch {
  Write-Error "Failed to generate mailbox inventory. $_"
  throw
}
