# CARVAJAL

## Resumen

Short summary on functionality and used technologies.

Cambiar versión de node de desarrollo

nvm install 14.19.3; nvm use 14.19.3;npm install gulp-cli yo @microsoft/generator-sharepoint --global; gulp serve;npm install gulp-cli --global;npm install yo --global;npm install @microsoft/generator-sharepoint --global

# Volver De Otro Desarrollo

Abrir terminal con permisos de administrador y ejecute:

1. nvm use 14.20.0
2. npm run start

# Quick dev

Open PowerShell

nvm install 14.20.0; nvm use 14.20.0;npm install gulp-cli yo @microsoft/generator-sharepoint --global

## Instalar ambiente de desarrollo

Lea todo el presente documento antes de instalar el ambiente de desarrollo.
Windows

1. Instalar chocolatey desde su página web
2. Instalar nvm con chocolatey en un command prompt con permisos administrativos ejecutamos:
   choco install nvm.install -y
3. Instalar node después de finalizar lo anterior en una nueva ventana de command prompt:
   nvm install 14.20.0
4. Activamos la versión de node previamente instalada:
   nvm use 14.20.0
5. Instalamos librerias para ejecución de sharepoint
   npm install gulp-cli yo @microsoft/generator-sharepoint --global
   npm install gulp-cli --global
   npm install yo --global
   npm install @microsoft/generator-sharepoint --global
6. Instalar librerias del proyecto spfx desde command prompt, desde el directorio del código donde está el archivo package.json Ejecutamos:
   npm install
7. Instalar certificado de seguridad para conectarse con el espacio de trabajo online
   gulp trust-dev-cert

Linux

1. Instalar nvm: Descargar y ejecutar automaticamente el instalador:
   wget -qO- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.1/install.sh | bash

   cree un archivo .bashrc en la carpeta de usuario y agregue:
   export NVM_DIR="$([ -z "${XDG_CONFIG_HOME-}" ] && printf %s "${HOME}/.nvm" || printf %s "${XDG_CONFIG_HOME}/nvm")"
   [ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh" # This loads nvm
2. Valide la versión de nvm
   nvm --version
3. instale la versión 14
   nvm install 14.19.3 && nvm use 14.19.3 && node -v
4. Instalamos librerias para ejecución de sharepoint
   npm install gulp-cli --global && /
   npm install yo --global && /
   npm install  @microsoft/generator-sharepoint --global
5. Instalar librerias del proyecto spfx desde command prompt, desde el directorio del código donde está el archivo package.json Ejecutamos:
   npm install
6. Ejecute sin instalar el certificado
   gulp serve --nobrowser
   Antes de abrir el área de trabajo de SharePoint Online, acceda a la URL https://localhost:4321/temp/manifests.js en la advertensia que no es seguro, click en continuar.
7. Ahora abra el área de trabajo de SharePoint Online.
   "https://carvajal.sharepoint.com/sites/flujosprocesos/_layouts/workbench.aspx"  9. Crear certificado de seguridad autofirmado en el proyecto
   gulp trust-dev-cert
8. Instalar el certificado
   sudo apt install libnss3-tools -y
   sudo chmod +x installspfx.sh
   sudo mkdir /usr/local/share/ca-certificates/extra
   cp ~/.rushstack/rushstack-serve.pem /usr/local/share/ca-certificates/extra/rushstack-serve.crt
   sudo update-ca-certificates -f
9. Instale el certificado en el navegador "No funciona estos pasos"
   cp ~/.rushstack/rushstack-serve.pem ./rushstack-serve.pem
   cp ~/.rushstack/rushstack-serve.pem ./rushstack-serve.crt
   cp ~/.rushstack/rushstack-serve.key ./rushstack-serve.key
   sudo ./installspfx.sh
10. Ahora abra el área de trabajo de SharePoint Online.

En el archivo launch.json de la carpeta .vscode ya se dejo la ruta web para depurar con visual studio code fuente de información:

https://docs.microsoft.com/es-es/sharepoint/dev/spfx/debug-in-vscode

errores:
  No ejecutar con permisos adminsitrativos
  Ejecutar comandos en powershell

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
