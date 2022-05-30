import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'GestorDeAusenciasWebPartStrings';
import GestorDeAusencias from './components/GestorDeAusencias';
import { IGestorDeAusenciasProps } from './components/IGestorDeAusenciasProps';
import { getRandomString } from "@pnp/core";
import { spfi, SPFx } from "@pnp/sp";
export interface IGestorDeAusenciasWebPartProps {
  description: string;
}
export default class GestorDeAusenciasWebPart extends BaseClientSideWebPart<IGestorDeAusenciasWebPartProps> {
  private _isDarkTheme: boolean = false;
  private sp: any;
  /**
   * Inicia el elemento web
   * @returns 
   */
  protected async onInit(): Promise<void> {
    this.sp = spfi().using(SPFx(this.context));
    return await super.onInit();
  }
  public render(): void {
    console.clear();
    // Valida que este instalado la librería pnpjs
    (() => { console.log("PNPJS instalado: "+getRandomString(20)) })()
    //Crea el elemento react enviandole propiedades
    const element: React.ReactElement<IGestorDeAusenciasProps> = React.createElement(
      GestorDeAusencias, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      context: this.context
    });
    //Renderiza en elemento
    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
