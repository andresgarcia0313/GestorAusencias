import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import GestorDeAusencias from './components/GestorDeAusencias';
export default class GestorDeAusenciasWebPart extends BaseClientSideWebPart<any> {
  private _isDarkTheme: boolean = false;
  protected async onInit(): Promise<void> { return await super.onInit(); }
  public render(): void {
    console.clear();//limpia la consola
    const GestorDeAusenciasElement: React.ReactElement<any> = React.createElement(
      GestorDeAusencias, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      user:this.context.pageContext.user,
      context: this.context
    });//Creación de objeto react con sus props o propiedades    
    ReactDom.render(GestorDeAusenciasElement, this.domElement);//Renderiza en elemento    
  }

  /*private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }*/

  /*protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
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

  }*/

  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  /*protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              })]
            }
          ]
        }
      ]
    };
  }*/
}
