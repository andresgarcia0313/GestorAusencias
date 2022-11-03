import * as React from 'react';//Importar React para crear componentes react que son componentes web que se muestran en la pagina
import styles from './GestorDeAusencias.module.scss';//Importamos el css para darle estilo a la webpart o elemento web
import { escape } from '@microsoft/sp-lodash-subset';//importamos el paquete de escape para evitar inyecciones de código
import { spfi, SPFx } from "@pnp/sp";//importamos el paquete pnp para obtener las funciones de sharepoint
import "@pnp/sp/webs";//importamos el modulo de webs para poder obtener las webs de sharepoint
import "@pnp/sp/lists";//importamos el modulo de listas para poder obtener los datos de listas de sharepoint
import "@pnp/sp/items";//importamos el modulo de items para poder obtener los datos de items de sharepoint que son los datos de las listas
import "@pnp/sp/site-users/web";//importamos el modulo de usuarios para poder obtener los datos de usuarios de sharepoint
import "@pnp/sp/fields";//importamos el modulo de campos para poder obtener los datos de campos de sharepoint
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';//importamos el modulo de llamadas de datos servicios web
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";//importamos el modulo de people picker para seleccionar usuarios
import { ContentType, IContentType } from "@pnp/sp/content-types";//importamos el modulo de content types para obtener los tipos de contenido de las listas
import { IField, IFieldAddResult, FieldTypes } from "@pnp/sp/fields/types";
import {
  ListView, //Importamos el modulo de listview para mostrar los datos de la tabla
  IViewField, //Importamos la interface de campos de vista para mostrar los datos de la tabla y que asegure el tipo de dato
  SelectionMode,//Importamos el modo de seleccion de la tabla para que solo se pueda seleccionar una fila o varias
  GroupOrder,//Importamos el orden de los grupos de la tabla para que se muestren los grupos en la tabla
  IGrouping //Importamos la interface de agrupacion de la tabla para que se muestren los grupos en la tabla
} from "@pnp/spfx-controls-react/lib/ListView";
import {
  DateTimePicker,//Importamos el modulo de date time picker para seleccionar fechas y horas
  DateConvention,//Importamos la convencion de fecha para que se muestre la fecha en el formato que queramos
  TimeConvention//Importamos la convencion de hora para que se muestre la hora en el formato que queramos
} from '@pnp/spfx-controls-react/lib/DateTimePicker';//importamos el modulo de datetimepicker para seleccionar fechas
export default class GestorDeAusencias extends React.Component<any, any> {//Clase Gestiornar Ausencias que sirve para crear el componente web y extiende de React.Component que es la clase base de react para crear componentes web y tiene dos parametros que son las propiedades y el estado
  private sp = spfi().using(SPFx(this.props.context));//Variable sp para poder usar las funciones de pnp y obtener los datos de sharepoint
  constructor(props: any) { //Constructor de la clase para inicializar las variables para usarlas en la clase
    super(props); //Llamamos al constructor de la clase padre
    this.state = {//Estado inicial de la webpart
      listSelectedViewTask: [],//Lista de tareas seleccionadas
      listViewTask: [],//Lista de tareas mostradas en la tabla
      showPeoplePickerAusente: false,//Mostrar people picker de ausente para el grupo de sharepoint de propietarios flujos de trabajo cts
      userIdAusente: null, //Id del usuario por ausentarse
      userIdDelegado: null, //Id del usuario a delegar actividades
    };
  }
  public async componentDidMount(): Promise<void> {//Funcion que se ejecuta cuando se carga la webpart y actualmente obtiene las tareas del usuario loqueado
    var groups = await this.sp.web.currentUser.groups();//Obtenemos los grupos del usuario actual
    var group: any = groups.filter(g => g.Title == "Propietarios Flujos de Trabajo CTS")[0];//Filtramos el grupo de propietarios de flujos de trabajo
    if (group != undefined) this.setState({ showPeoplePickerAusente: true });//Si el usuario pertenece al grupo de propietarios de flujos de trabajo se muestra el people picker
    var userId = (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id;//Obtenemos el id del usuario que esta logueado
    this.setState({ userIdAusente: userId });//Guardamos el id de la persona a ausentar en el estado
    var tasks = await this.getTasksFromTaskListsByUserId(userId);//Obtenemos las tareas del usuario que esta logueado
    this.setState({ listViewTask: tasks });//Guardamos las tareas en el estado
    this.changeOwnerOfCurrentAbsenceTasks();
  }
  private setStartDate = async (date: Date) => { this.setState({ startDate: date }); };
  private setEndDate = async (date: Date) => { this.setState({ endDate: date }); };
  private getPeoplePickerItemsAusente = async (items: any[]) => {//Obtiene la persona a ausentar  y se ejecuta cuando se selecciona un usuario en el people picker de ausente
    this.setState({ listViewTask: [] });//Borra el listado de tareas del state y a su vez de la tabla de tareas que ve el usuario
    if (items.length > 0) {//Validar que exista una persona seleccionada de
      var userIdAusente = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id//obtener el id del usuario
      var tasksAusente = await this.getTasksFromTaskListsByUserId(userIdAusente);// Obtener tareas del usuario
      this.setState({ userIdAusente: userIdAusente });//Guardamos el id de la persona a ausentar en el estado
      this.setState({ listViewTask: tasksAusente });// Establecer en state las actividades para que se presenten en la tabla de tareas
    }
  }
  private getPeoplePickerItemsDelegado = async (items: any[]) => {//Obtiene la persona a delegar actividades y se ejecuta cuando se selecciona un usuario en el people picker de delegado
    if (items.length > 0) {//validar que exista una persona seleccionada de delegado
      var userIdDelegado = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id;//obtener el id del usuario delegado
      this.setState({ userIdDelegado: userIdDelegado });//Establecer en state el id del usuario ausente para que se presente en la tabla de tareas
    }
  }

  public getTasksFromTaskListsByUserId = async (userId: number) => {
    var inicio: any = (new Date()); //Obtener momento actual
    var lists = (await this.props.context.spHttpClient.get(//Obtener listas de tareas
      this.props.context.pageContext.web.absoluteUrl +//Url del sitio
      `/_api/web/lists?$filter=(Hidden eq false) and (BaseTemplate eq 171)&$select=Title`,//Filtro para obtener solo las listas de tareas
      SPHttpClient.configurations.v1)//Configuracion de la peticion
      .then(async (r: SPHttpClientResponse) => {//Obtener respuesta de la peticion
        return (await r.json()).value;//Obtener listas de tareas
      }));//Obtener listas de tareas
    var _taskLists: any[] = new Array();//Lista de tareas
    for (let l of lists) _taskLists.push(this.sp.web.lists.getByTitle(l.Title).items());//Obtener tareas de cada lista de tareas
    var taskLists = await Promise.all(_taskLists);//Obtener tareas de cada lista de tareas    
    for (var i = 0; i < lists.length; i++)//Recorrer listas de tareas
      for (var x = 0; x < taskLists.length; x++)//Recorrer tareas de cada lista de tareas
        if (x == i)//Si el indice de la lista de tareas es igual al indice de las tareas de la lista de tareas
          for (var task of taskLists[x])//Recorrer tareas de la lista de tareas
            task.List = lists[i].Title;//Agregar propiedad de lista a la tarea
    var taskByUserId: any[] = new Array();//Lista de tareas del usuario
    for (var taskList of taskLists) for (var task of taskList)//Recorrer listas de tareas y tareas de cada lista de tareas
      if (task.PercentComplete < 1 && task.AssignedToId[0] == userId)//Si la tarea no esta completada y el usuario es el asignado
        taskByUserId.push(task);//Agregar tarea a la lista de tareas del usuario
    var fin: any = (new Date());//Obtener momento actual
    console.log("GetTasks: " + ((fin - inicio) / 1000) + "s");//Mostrar tiempo de ejecucion
    return taskByUserId;//Retornar lista de tareas del usuario
  }
  //Funcion para obtener la lista de elementos seleccionados en la lista de tareas
  private _getSelectionOfListView = (tasks: any[]) => {
    this.setState({ listSelectedViewTask: tasks });
  }
  public changeOwner = async () => {//metodo que cambia de propietario de la tarea por tarea, usuario ausente, usuario delegado
    try {
      var inicio: any = (new Date()); //Obtener momento de inicio      
      if (this.state.peoplePickerAusente != undefined)
        var userIdAusente = (await this.sp.web.siteUsers.getByEmail(this.state.peoplePickerAusente[0].secondaryText)()).Id;
      else
        var userIdAusente = (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id;
      var tasksToChangeOwner: any[] = new Array();//declarar array de promesas de cambio de propietario de tarea 
      for (var task of this.state.listSelectedViewTask) //Recorre las tareas seleccionadas para cambiar de propietario
        tasksToChangeOwner[task.Id] = this.sp.web.lists.getByTitle(task.List).items.getById(task.Id).update({ AssignedToId: [this.state.userIdDelegado], });//Cambia el propietario de la tarea
      await Promise.all(tasksToChangeOwner);//Espera a que se resuelvan todas las promesas de cambio de propietario de tarea para continuar
      var tasks = await this.getTasksFromTaskListsByUserId(userIdAusente);//Obtenemos las tareas del usuario que esta logueado
      this.setState({ listViewTask: [] });//Guardamos las tareas en el estado
      this.setState({ listViewTask: tasks });//Guardamos las tareas en el estado
      var fin: any = (new Date()); console.log("Cambiado Asignación De Tareas  " + ((fin - inicio) / 1000) + "s");//Obtener momento de fin
      this.saveAbsence();//Guarda la ausencia del usuario
    } catch (e) { console.log(e); }//Captura errores
  }
  public changeOwnerOfCurrentAbsenceTasks = async () => {
    try {
      console.log("changeOwnerOfCurrentAbsenceTasks");
      for (let item of (await this.sp.web.lists.getByTitle("ListForDelegationOfAbsences").items())) {//Recorre las ausencias para obtener las tareas de los usuarios ausentes
        if ((new Date(item.Inicio)) < (new Date()) && (new Date()) < (new Date(item.Fin))) {//Si la fecha de inicio de la ausencia es menor a la fecha actual y es menor a la fecha de fin de la ausencia entonces es una ausencia actual que se esta presentando
          this.getTasksFromTaskListsByUserId(item.AusenteId).then(//Obtiene las tareas del usuario ausente para cambiar el propietario de las tareas
            async tasks => {//Recibe las tareas del usuario ausente
              for (var task of tasks) {//Recorre las tareas del usuario ausente
                this.sp.web.lists.getByTitle(task.List).items.getById(task.Id).update({ AssignedToId: [item.DelegadoId], });//Cambia el propietario de la tarea
              }
            }
          );
        }
      }
      console.log("Finalizado changeOwnerOfCurrentAbsenceTasks");
    } catch (e) { console.log(e); }//Captura errores
  }
  //funcion para guardar en la lista de ListForDelegationOfAbsences los datos de persona ausente, persona delegada, fecha de ausencia y fecha de regreso
  public saveAbsence = async () => {
    try {
      var userIdAusente = this.state.userIdAusente;//Obtener id del usuario ausente
      var userIdDelegado = this.state.userIdDelegado;//Obtener id del usuario delegado
      var inicio = this.state.startDate.toISOString();//Obtener fecha de inicio de la ausencia
      var fin = this.state.endDate.toISOString();//Obtener fecha de fin de la ausencia
      var reg = {//Objeto con los datos de la ausencia
        AusenteId: userIdAusente,//Id del usuario ausente
        AusenteStringId: userIdAusente.toString(),//Id del usuario ausente en string
        DelegadoId: userIdDelegado,//Id del usuario delegado
        DelegadoStringId: userIdDelegado.toString(),//Id del usuario delegado en string
        Inicio: inicio,//Fecha de inicio de la ausencia
        Fin: fin//Fecha de fin de la ausencia
      }
      await this.sp.web.lists.getByTitle("ListForDelegationOfAbsences").items.add(reg);//Agregar registro a la lista de ausencias
    } catch (e) { console.log(e); }//Captura errores
  }
  public render(): React.ReactElement<any, any> {//Renderiza el componente
    const { userDisplayName } = this.props;//captura propiedades
    var jsx = (//html del componente
      <section className={styles.gestorDeAusencias}>
        <div className={styles.welcome}>
          <h3>Gestor y creador de ausencias</h3>
          <h4>Bienvenido señor(a) {escape(userDisplayName)}!</h4>
          <DateTimePicker label="Inicio"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12}
            onChange={this.setStartDate} />
          <DateTimePicker label="Fin"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12}
            onChange={this.setEndDate} />
          {this.state.showPeoplePickerAusente
            ? <p><PeoplePicker context={this.props.context}
              placeholder="Persona a ausentarse..."
              personSelectionLimit={1} showtooltip={false} required={true}
              disabled={false} onChange={this.getPeoplePickerItemsAusente}
              showHiddenInUI={true} principalTypes={[PrincipalType.User]}
              resolveDelay={50} /> </p>
            : ""}
          <p><PeoplePicker context={this.props.context}
            placeholder="Persona a delegarle las actividades..."
            personSelectionLimit={1} showtooltip={false} required={true}
            disabled={false} onChange={this.getPeoplePickerItemsDelegado}
            showHiddenInUI={false} principalTypes={[PrincipalType.User]}
            resolveDelay={50} /></p>
          <p>Actividades</p>
          <ListView
            viewFields={[{
              name: "Title", displayName: "Actividad", isResizable: true,
              sorting: true, minWidth: 0, maxWidth: 150
            }, {
              name: "List", displayName: "Lista de tarea", isResizable: true,
              sorting: true, minWidth: 0, maxWidth: 150
            }]}
            filterPlaceHolder="Busque lista o actividad..."
            selection={this._getSelectionOfListView}
            selectionMode={SelectionMode.multiple}
            iconFieldName="ServerRelativeUrl" items={this.state.listViewTask}
            dragDropFiles={false} stickyHeader={true}
            showFilter={true} defaultFilter="" compact={false} />
          <p>
            <button type="button" onClick={this.changeOwner}>
              Delegar Actividades</button>
          </p>
        </div>
      </section>
    );
    var url = window.location.href;//Obtiene la url actual
    var urlblock = "https://carvajal.sharepoint.com/sites/flujosprocesos";//url a bloquear la visualización del componente
    
    if (url == urlblock) jsx = (<div></div>);//Si la url actual es igual a la url a bloquear entonces no se muestra el componente pero muestra un campo vacio
    return jsx;
  }
}