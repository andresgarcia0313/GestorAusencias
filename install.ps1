$node = node --version;
if ($node) {
    Write-Output "Version de Node Instalada: $node";
}
else {
    Write-Output "No se encuentra instalado Node - Instalando Node";          
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'));# Esto permite instalar Chocolatey que sirve para instalar aplicacione como node
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")# Esto permite releer la variable de entorno Path
    choco install nvm.install;# Esto permite instalar NVM
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")# Esto permite releer la variable de entorno Path
    nvm install 14.20.0;# Esto permite instalar Node
    nvm use 14.20.0;# Esto permite usar Node
}

