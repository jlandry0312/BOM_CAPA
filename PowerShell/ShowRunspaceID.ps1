function ShowRunspaceID
{
    $id=[runspace]::DefaultRunspace.Id
    $app=[System.Diagnostics.Process]::GetCurrentProcess()
    [System.Windows.Forms.MessageBox]::Show("application: $($app.name)"+[Environment]::NewLine+"runspace ID: $id")
}