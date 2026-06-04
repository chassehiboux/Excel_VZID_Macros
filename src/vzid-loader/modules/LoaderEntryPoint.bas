Attribute VB_Name = "LoaderEntryPoint"
Option Explicit

Public Sub LoaderEntryPoint_HandleWorkbookOpen()
    On Error GoTo failed

    LoaderPaths_EnsureBaseFolders
    LoaderConfig_EnsureConfig
    LoaderLogging_Write "Loader startup. Version " & VZID_LOADER_VERSION

    LoaderUpdates_ActivatePendingIfPresent
    LoaderAddinHost_LoadMainAddin LoaderPaths_MainAddinPath()
    LoaderUpdates_CheckStartup
    Exit Sub

failed:
    LoaderLogging_Write "LoaderEntryPoint_HandleWorkbookOpen failed: " & Err.Number & " - " & Err.Description
End Sub
