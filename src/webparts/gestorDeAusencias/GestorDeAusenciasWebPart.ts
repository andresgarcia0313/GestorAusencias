import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import GestorDeAusencias from './components/GestorDeAusencias';
export default class GestorDeAusenciasWebPart extends BaseClientSideWebPart<any> {
  protected async onInit(): Promise<void> { return await super.onInit(); }
  public render(): void {
    console.clear();//limpia la consola
    const GestorDeAusenciasElement: React.ReactElement<any> = React.createElement(
      GestorDeAusencias, {
      userDisplayName: this.context.pageContext.user.displayName,
      user:this.context.pageContext.user,
      context: this.context
    });//Creación de objeto react con sus props o propiedades    
    ReactDom.render(GestorDeAusenciasElement, this.domElement);//Renderiza en elemento    
  }
  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }

}
