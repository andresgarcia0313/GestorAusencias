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
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
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
    console.log("componentDidMount()");
    this.setListItemsToStates("Tareas_FT-Facturacion Servicios");
    this.getTasksFromTaskLists();
  }
  public componentWillUnmount(): void {
    console.log("componentWillUnmount()");
  }
  private async getPeoplePickerItems(items: any[]) {
    if (items.length >= 1) 
    this.personToBeAbsent = items[0];
    var userId = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id
    console.log(items[0])
    debugger;
  }
  public async getListItemsByNameList(listname: string): Promise<any[]> {
    return new Promise<any[]>(async (resolve, reject) => {
      try { resolve(await this.sp.web.lists.getByTitle(listname).items()); } catch (error) { console.log(error); }
    });
  }
  public async getListNames() {
    console.log("getListNames()");
    return await this.props.context.spHttpClient.get(
      this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists?` + `$filter=(Hidden eq false) and (BaseTemplate eq 171)` + `&$select=Title`,
      SPHttpClient.configurations.v1
    ).then(async (response: SPHttpClientResponse) => { return await response.json(); });
  }
  public async getListNameFromService() {
    console.log("getListNameFromService()");
    return await this.getListNames().then(r => { return r.value });
  }
  public async getTasksFromTaskLists() {
    console.log("getTasksFromTaskLists()");//registro del metodo ejecutado    
    var dataNameList = await this.getListNameFromService();//obtiene listas tipo tarea
    var dataPromiseTaskAllList: any[] = new Array();//variable donde almacenara las promesas de obtener las tareas de las listas
    var items: any;//elementos donde se almacenara el nombre de las listas
    var userId = (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id;//obtiene el id del usuario de contexto instanciadose y ejecutandose en el acto
    console.dir(this.state);//imprime el state para saber que variables obtiene en este instante
    for (let num = 0; num < dataNameList.length; num++) {//recorre los nombres de las listas tipo tareas para obtener sus registros
      items = this.sp.web.lists.getByTitle(dataNameList[num].Title).items();//almacena los registros de cada lista de tareas en items
      for (let item of items) {
        item.List = dataNameList[num].Title;//agrega el titulo de la lista de cada item
      }
      dataPromiseTaskAllList.push(items);//Almacena las promesas en un array para esperar que le llegue los datos prometidos
    }
    var result = await Promise.all(dataPromiseTaskAllList);//espera que se resuelva los contenidos de cada lista
    for (let num = 0; num < dataNameList.length; num++) {//recorre los nombres de las listas tipo tareas para obtener sus registros
      for (let item of result) {
        item.List = dataNameList[num].Title;//agrega el titulo de la lista de cada item dentro de result
      }
      
    }
    console.log("Tareas con nombres de listas:");
    console.log(result);
    for (var list of result){
      for(var item of list){
        console.dir(item.Title+":"+item.AssignedToStringId[0]);
      }
    }
    //Mostrando tareas
    //Obteniendo tareas de la persona
    debugger;
  }
  public setListItemsToStates = (listName: string): void => {
    console.log("setListItemsToStates()"+listName);
    this.getListItemsByNameList(listName).then(res => { this.setState({ items: res }); });
  }
  public render(): React.ReactElement<IGestorDeAusenciasProps, any> {
    console.log("render() de Gestor Ausencias" + new Date().toISOString());
    const { description, isDarkTheme, hasTeamsContext, userDisplayName, context } = this.props;
    const users = this.sp.web.siteUsers().then(r => { console.log("Usuarios Del Sitio:"); console.dir(r) });
    console.log("Usuario");
    console.dir(this.props.user.email);

    return (
      <section className={`${styles.gestorDeAusencias} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Bienvenido Usuario: {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>El presente formulario es el Gestor De Ausencias</h3>
          <p>Aquí usted delega actividades a otra persona en caso de ausentarse.</p>
          <p><PeoplePicker
            context={this.props.context} titleText="Persona a ausentar"
            personSelectionLimit={1} showtooltip={false}
            required={true} disabled={false}
            onChange={this.getPeoplePickerItems} showHiddenInUI={false}
            principalTypes={[PrincipalType.User]} resolveDelay={1000} />
          </p>
          <p>
            <PeoplePicker
              context={this.props.context} titleText="Persona a delegar actividades"
              personSelectionLimit={1} showtooltip={false}
              required={true} disabled={false}
              onChange={this.getPeoplePickerItems} showHiddenInUI={false}
              principalTypes={[PrincipalType.User]} resolveDelay={1000} />
          </p>
          <DateTimePicker label="Fecha inicial de ausentismo"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12} />
          <DateTimePicker label="Fecha final de ausentismo"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12} />
          <p>Actualmente Posee Las Siguientes Actividades Asignadas</p>
          <ListView items={this.state.items}
            viewFields={[
              {
                name: "Title",
                displayName: "Actividad",
                isResizable: true,
                sorting: true,
                minWidth: 0,
                maxWidth: 150
              },
              {
                name: "List",
                displayName: "Lista",
                isResizable: true,
                sorting: true,
                minWidth: 0,
                maxWidth: 150
              }
            ]}
            iconFieldName="ServerRelativeUrl" compact={true}
            selectionMode={SelectionMode.multiple}
            showFilter={true}
            defaultFilter=""
            filterPlaceHolder="Buscar..."
            dragDropFiles={true}
            //groupByFields={groupByFields} //selection={this._getSelection} //onDrop={this._getDropFiles} //className={styles.listWrapper} //listClassName={styles.list}
            stickyHeader={true} />
          <p>
            <button type="button">Guardar y Generar Ausencia</button>
          </p>
        </div>
      </section>
    );
  }
}
