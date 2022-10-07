import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import GestorDeAusencias from './components/GestorDeAusencias';
export default class GestorDeAusenciasWebPart extends BaseClientSideWebPart<any> {
  protected async onInit(): Promise<void> {
    return await super.onInit();
  }
  public render(): void {
    console.clear();//limpia la consola
    var props: any = {
      userDisplayName: this.context.pageContext.user.displayName,
      user: this.context.pageContext.user,
      context: this.context
    }
    const GestorDeAusenciasElement: React.ReactElement<any> =
      React.createElement(
        GestorDeAusencias,
        props
      );//Crea el componente web con las propiedades
    ReactDom.render(GestorDeAusenciasElement, this.domElement);//Muestra el elemento web
  }
  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
