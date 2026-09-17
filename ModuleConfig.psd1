# ---------------------------------------------------------------------------
# EndpointGuy Intune Toolkit - module configuration
#
# Place this file next to Toolkit.ps1. It overrides settings baked into
# Toolkit.ps1, so the same build can be shipped to every site with only this
# file changing.
#
# Any key you delete or omit falls back to whatever Toolkit.ps1 has set.
# ---------------------------------------------------------------------------
@{

    # -----------------------------------------------------------------------
    # Modules  -  which actions are available
    #
    # $true  = module is loaded and its button is shown
    # $false = module is not loaded, its button is hidden, and its Graph
    #          scopes are not requested at sign-in
    # -----------------------------------------------------------------------
    Modules = @{

        # Device Actions
        CopyDeviceGroups   = $true   # Copy assigned security groups to another device
        RemoveDeviceGroups = $true   # Remove the device from its assigned groups
        AssetStatus        = $true   # Set the Intune Management name to a lifecycle status

        # Bulk Actions
        BulkAddToGroup     = $true   # Add every device in a CSV to one group

        # Reporting
        AppDependencyCheck = $true   # Read-only Win32 app dependency report
    }

    # -----------------------------------------------------------------------
    # Asset Status  -  the statuses offered by the Update Asset Status module
    #
    # These become the drop-down, in the order listed here, and they are also
    # the approved list: a value that is not on this list is refused before
    # anything is written to Graph.
    #
    # Whatever you put here is written verbatim into the Intune Management
    # name of the device, so type it exactly as you want it to appear in the
    # console - 'In-Stock' and 'In Stock' are two different labels.
    #
    # Delete this whole key, or leave the list empty, to keep the six
    # built-in statuses.
    #
    # Reporting elsewhere keys off these strings, so renaming one does NOT
    # update devices already carrying the old label - they keep the old text
    # until each one is set again.
    # -----------------------------------------------------------------------
    AssetStatusValues = @(
        'In-Stock'
        'Retired'
        'Recycled'
        'Stolen'
        'Legalhold'
        'Lost'
    )
}
