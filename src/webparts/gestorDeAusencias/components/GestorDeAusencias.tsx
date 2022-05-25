import * as React from 'react';
import styles from './GestorDeAusencias.module.scss';
import { IGestorDeAusenciasProps } from './IGestorDeAusenciasProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './lib/models/IListItem';
export default class GestorDeAusencias extends React.Component<IGestorDeAusenciasProps, {}> {
  private personToBeAbsent: any;
  private endPoint: String;
  private result: any;
  private _getPeoplePickerItems(items: any[]) {
    if (items.length >= 1) {
      this.personToBeAbsent = items[0]
      console.log('Persona:', this.personToBeAbsent);
    }
  }
  /*private _getListData(): Promise<any> {
    this.endPoint = `/_api/web/lists?$filter=Hidden eq false`
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + this.endPoint, SPHttpClient.configurations.v1
    ).then(
      (response: SPHttpClientResponse) => {
        return response.json();
      }
    );
  }*/
  public render(): React.ReactElement<IGestorDeAusenciasProps> {
    console.log("Se consulta las listas");
    //this.result = this._getListData();
    console.log(this.result);
    console.log("Inicio Del Componente React");
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName, context
    } = this.props;
    console.log("Consultar usuarios");
    const users = this.props.sp.web.siteUsers();
    users.then(Response => {/*console.log(Response);*/ });
    var listItems: Array<IListItem> = new Array<IListItem>();
    listItems.push({Nombre:"Texto"});
    listItems.push({Nombre:"Texto1"});
    var listItems2:any=[{Nombre:"Texto"},{Nombre:"Texto1"}]
    return (
      <section className={`${styles.gestorDeAusencias} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Bienvenido, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Descripción del presente módulo : <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Bienvenido Al Gestor De Ausencias!</h3>
          <p>
            El gestor de ausencias es un elemento web que usted puede usar para delegar las actividades a otra persona en caso de ausencia de una persona. Siendo la forma más sencilla de realizar esto.
          </p>
          <p>
            <PeoplePicker
              context={this.props.context}
              titleText="Persona A Ausentar"
              personSelectionLimit={1}
              showtooltip={true}
              required={true}
              disabled={false}
              onChange={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
          </p>
          <ListView
            items={listItems2}
            viewFields={
              [
                {
                  name: "Nombre",
                  displayName: "Columna1",
                  isResizable: true,
                  sorting: true,
                  minWidth: 0,
                  maxWidth: 150
                }
              ]
            }
            iconFieldName="ServerRelativeUrl"
            compact={true}
            selectionMode={SelectionMode.multiple}
            //selection={this._getSelection}
            showFilter={true}
            defaultFilter=""
            filterPlaceHolder="Buscar..."
            // groupByFields={groupByFields}
            dragDropFiles={true}
            //onDrop={this._getDropFiles}
            stickyHeader={true}
          //className={styles.listWrapper}
          //listClassName={styles.list} 
          />
        </div>
      </section>
    );
  }
}
