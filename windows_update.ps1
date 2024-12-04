"Initializing Windows Update"
Install-PackageProvider -Name NuGet -ForceBootstrap -force | Out-Null
Install-Module -Name PSWindowsUpdate -AllowClobber -SkipPublisherCheck -force | Out-Null
"Downloading and installing all available updates..."
Install-WindowsUpdate -AcceptAll -IgnoreReboot | Out-Null
"Windows Update finished."