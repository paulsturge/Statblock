@{
  # Identity
  RootModule        = 'RQG.Authoring.psm1'
  ModuleVersion     = '1.0.0'
  GUID              = 'b2a3f6fe-7c15-4e98-9b73-6f8e1a5a9e9f'
  Author            = 'You'
  CompanyName       = '—'
  Copyright         = '© You'

  # Requirements
  PowerShellVersion = '7.0'            # you’re on 7.5.x
  RequiredModules   = @('ImportExcel') # already used in your tools

  # Metadata
  Description       = 'Authoring tools to add RuneQuest creatures (stats, hit locations, weapons) from clipboard/text.'

  # Exports (psm1 also controls this; here for clarity)
  FunctionsToExport = @(
    'New-StatDiceRow',
    'New-HitLocationSheet',
    'Import-CreatureWeaponsFromText',
    'New-Creature'
  )
  CmdletsToExport   = @()
  VariablesToExport = @()
  AliasesToExport   = @()

  PrivateData       = @{
    PSData = @{
      Tags = @('RQG','Authoring','ImportExcel','RuneQuest')
    }
  }
}
