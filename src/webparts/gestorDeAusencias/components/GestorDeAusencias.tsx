import * as React from 'react';
import styles from './GestorDeAusencias.module.scss';
import { IGestorDeAusenciasProps } from './IGestorDeAusenciasProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
/** variable de agrupamientos*/
const groupByFields: IGrouping[] = [{ name: "ListName", order: GroupOrder.ascending }];

export default class GestorDeAusencias extends React.Component<IGestorDeAusenciasProps, any> {
  private personToBeAbsent: any;//Variable para almacenar datos de persona a ausentar
  private sp = spfi().using(SPFx(this.props.context));//inicializa libreria pnpjs
  constructor(props: IGestorDeAusenciasProps) {
    super(props);
    this.state = {}
  }
  public componentDidMount(): void {
    this.setListItemsToStates("Tareas_FT-Facturacion Servicios");
    this.getTasksFromTaskLists();
  }
  public componentWillUnmount(): void { }
  private getPeoplePickerItems(items: any[]) {
    if (items.length >= 1) this.personToBeAbsent = items[0];
  }
  public async getListItemsByNameList(listname: string): Promise<any[]> {
    return new Promise<any[]>(async (resolve, reject) => {
      try { resolve(await this.sp.web.lists.getByTitle(listname).items()); } catch (error) { console.log(error); }
    });
  }
  public async getListNames() {
    return await this.props.context.spHttpClient.get(
      this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists?` + `$filter=(Hidden eq false) and (BaseTemplate eq 171)` + `&$select=Title`,
      SPHttpClient.configurations.v1
    ).then(async (response: SPHttpClientResponse) => { return await response.json(); });
  }
  public async getListNameFromService() {
    return await this.getListNames().then(r => { return r.value });
  }
  public async getTasksFromTaskLists() {
    var dataNameList = await this.getListNameFromService();
    var dataPromiseTaskAllList: any[] = new Array();
    var items: any;
    var userid = this.sp.web.siteUsers.getByEmail(this.props.user.email)();
    userid.then(r => { console.log("idUser:" + r.Id); })
    for (let num = 0; num < dataNameList.length; num++) {
      items = this.sp.web.lists.getByTitle(dataNameList[num].Title).items();
      dataPromiseTaskAllList.push(items);
    }
    var result = await Promise.all(dataPromiseTaskAllList);
    console.log(result);
  }
  public setListItemsToStates = (listName: string): void => {
    this.getListItemsByNameList(listName).then(res => { this.setState({ items: res }); });
  }
  public render(): React.ReactElement<IGestorDeAusenciasProps, any> {
    console.log("render() de Gestor Ausencias" + new Date().toISOString());
    const { description, isDarkTheme, hasTeamsContext, userDisplayName, context } = this.props;
    const users = this.sp.web.siteUsers().then(r => { console.log("Usuarios Del Sitio:"); console.dir(r) });
    console.log("Usuario");
    console.dir(this.props.user);
   
    return (
      <section className={`${styles.gestorDeAusencias} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <ListView items={this.state.items}
            viewFields={[{ name: "Title", displayName: "Columna1", isResizable: true, sorting: true, minWidth: 0, maxWidth: 150 }]}
            iconFieldName="ServerRelativeUrl" compact={true}
            selectionMode={SelectionMode.multiple}
            showFilter={true}
            defaultFilter=""
            filterPlaceHolder="Buscar..."
            dragDropFiles={true}
            //groupByFields={groupByFields} //selection={this._getSelection} //onDrop={this._getDropFiles} //className={styles.listWrapper} //listClassName={styles.list}
            stickyHeader={true} />
          <h2>Bienvenido, {escape(userDisplayName)}!</h2>

          <div>Descripción del presente módulo : <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>El presente elemento web es el Gestor De Ausencias</h3>
          <p>El gestor de ausencias es un elemento web que usted puede usar para delegar las actividades a otra persona en caso de ausencia de una persona. Siendo la forma más sencilla de realizar esto.</p>
          <p><PeoplePicker
            context={this.props.context} titleText="Persona A Ausentar"
            personSelectionLimit={1} showtooltip={false}
            required={true} disabled={false}
            onChange={this.getPeoplePickerItems} showHiddenInUI={false}
            principalTypes={[PrincipalType.User]} resolveDelay={1000} />
          </p>
        </div>
      </section>
    );
  }
}
