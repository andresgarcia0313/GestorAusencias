import * as React from 'react';//Importar React para crear componentes react que son componentes web que se muestran en la pagina
import styles from './GestorDeAusencias.module.scss';//Importamos el css para darle estilo a la webpart o elemento web
import { escape } from '@microsoft/sp-lodash-subset';//importamos el paquete de escape para evitar inyecciones de código
import { spfi, SPFx } from "@pnp/sp";//importamos el paquete pnp para obtener las funciones de sharepoint
import "@pnp/sp/webs";//importamos el modulo de webs para poder obtener las webs de sharepoint
import "@pnp/sp/lists";//importamos el modulo de listas para poder obtener los datos de listas de sharepoint
import "@pnp/sp/items";//importamos el modulo de items para poder obtener los datos de items de sharepoint que son los datos de las listas
import "@pnp/sp/site-users/web";//importamos el modulo de usuarios para poder obtener los datos de usuarios de sharepoint
import "@pnp/sp/fields";//importamos el modulo de campos para poder obtener los datos de campos de sharepoint
import "@pnp/sp/items/get-all";//importamos el modulo de items para poder obtener los datos de items de sharepoint que son los datos de las listas
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
//import { PopupWindowPosition } from '@microsoft/sp-property-pane';
var log = console.log;//creamos una variable para hacer console.log mas corto
export default class GestorDeAusencias extends React.Component<any, any> {//Clase Gestiornar Ausencias que sirve para crear el componente web y extiende de React.Component que es la clase base de react para crear componentes web y tiene dos parametros que son las propiedades y el estado
  private sp = spfi().using(SPFx(this.props.context));//Variable sp para poder usar las funciones de pnp y obtener los datos de sharepoint
  private PeoplePickerDelegado;
  constructor(props: any) { //Constructor de la clase para inicializar las variables para usarlas en la clase
    super(props); //Llamamos al constructor de la clase padre
    this.state = {//Estado inicial de la webpart
      listSelectedViewTask: [],//Lista de tareas seleccionadas
      listViewTask: [],//Lista de tareas mostradas en la tabla
      showPeoplePickerAusente: false,//Mostrar people picker de ausente para el grupo de sharepoint de propietarios flujos de trabajo cts
      userIdAusente: null, //Id del usuario por ausentarse
      userIdDelegado: null, //Id del usuario a delegar actividades
      loading: false,//Mostrar loading
    };
  }
  public async componentDidMount(): Promise<void> {//Funcion que se ejecuta cuando se carga la webpart y actualmente obtiene las tareas del usuario loqueado
    var group;
    var groups = await this.sp.web.currentUser.groups();//Obtenemos los grupos del usuario actual
    group = groups.filter(g => g.Title == "Propietarios Flujos de Trabajo CTS")[0];
    if (group == undefined) group = groups.filter(g => g.Title == "Propietarios Flujos de trabajo CTS MX")[0];
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
    console.log("Personas que llegan cuando eligo el ausente");
    console.dir(items);
    this.setState({ listViewTask: [] });//Borra el listado de tareas del state y a su vez de la tabla de tareas que ve el usuario
    if (items.length > 0) {//Validar que exista una persona seleccionada de
      var userIdAusente = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id//obtener el id del usuario
      var tasksAusente = await this.getTasksFromTaskListsByUserId(userIdAusente);// Obtener tareas del usuario
      this.setState({ userIdAusente: userIdAusente });//Guardamos el id de la persona a ausentar en el estado
      this.setState({ listViewTask: tasksAusente });// Establecer en state las actividades para que se presenten en la tabla de tareas
    }
    else {
      alert("Persona no se ha podido consultar no selecciono persona o si esta seleccionada por favor actualice la página con Ctrl+F5");
    }
  }
  private getPeoplePickerItemsDelegado = async (items: any[]) => {//Obtiene la persona a delegar actividades y se ejecuta cuando se selecciona un usuario en el people picker de delegado
    if (items.length > 0) {//validar que exista una persona seleccionada de delegado
      var userIdDelegado = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id;//obtener el id del usuario delegado
      this.setState({ userIdDelegado: userIdDelegado });//Establecer en state el id del usuario ausente para que se presente en la tabla de tareas
    }
  }
  //userid, columnas, filtro, 
  public getTasksFromTaskListsByUserId = async (userId: number) => {
    log("Obteniendo tareas del usuario con id: " + userId + "Para consultarlas o mostrarlas en la tabla");
    this.setState({ loading: true });//Mostrar loading
    let listTasks = [];//Lista de tareas global
    const promesas = [];//Lista de promesas para obtener las tareas de las listas de tareas
    for (let l of (await this.sp.web.lists.select("Title,ItemCount").filter("BaseTemplate eq 171")()))//Recorremos las listas de tareas para obtener las tareas de cada una
      for (let i = 0; i < (Math.ceil(l.ItemCount / 1600)); i++)//Recorremos las paginas de las listas de tareas para obtener las tareas de cada pagina sin exceder el limite de 5000 items por consulta en este caso la dejamos de 1600
        promesas.push(this.sp.web.lists.getByTitle(l.Title).items//Agregamos la consulta al listado de consultas para esperar y obtener las tareas de cada lista de forma paralela
          .select("Id", "Title", "AssignedToStringId")//Seleccionamos los campos que se van a obtener
          .filter(`AssignedToStringId eq ${userId} and PercentComplete lt 1 and Id ge ${i * 1600} and Id le ${(i * 1600) + 1600 - 1}`)//Filtramos las tareas por el id del usuario y que no esten completadas y que están en el rango de paginado
          .getAll().then(
            r => {
              listTasks = listTasks.concat(r.map(r => ({
                Id: r.Id,
                List: l.Title,
                Title: r.Title
              })));
            }))//Agregamos las tareas a la lista de tareas global
    await Promise.all(promesas);
    console.log("Tareas obtenidas y capturadas ya sea para enviarlas o mostrarlas en la tabla");
    console.dir(listTasks);
    this.setState({ loading: false });//Ocultar loading
    return listTasks;
  }
  //Funcion para obtener la lista de elementos seleccionados en la lista de tareas
  private _getSelectionOfListView = (tasks: any[]) => { this.setState({ listSelectedViewTask: tasks }); }
  public changeOwner = async () => {//metodo que cambia de propietario de la tarea por tarea, usuario ausente, usuario delegado
    try {
      //valide que se haya seleccionado una persona a ausentar
      if (this.state.userIdDelegado == null) alert("Debe seleccionar persona a delegar actividades");
      //Si no tiene persona seleccionada asigne la que esta autenticada
      if (this.state.userIdAusente == null) this.setState({ userIdAusente: (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id });
      var promisesUpdate: any[] = new Array();//declarar array de promesas de cambio de propietario de tarea 
      for (var task of this.state.listSelectedViewTask) //Recorre las tareas seleccionadas para cambiar de propietario
        promisesUpdate[task.Id] = this.sp.web.lists.getByTitle(task.List).items.getById(task.Id)
          .update({ AssignedToId: [this.state.userIdDelegado] });
      await Promise.all(promisesUpdate);//Espera a que se resuelvan todas las promesas de cambio de propietario de tarea para continuar
      //si promisesUpdate tiene alguna promesa muestre mensaje de tareas cambiadas
      if (promisesUpdate.length > 0) { alert("Se ha cambiado el propietario de las tareas seleccionadas") }
      this.setState({ listViewTask: [] })//Limpiar tabla de tareas
      var actividades = await this.getTasksFromTaskListsByUserId(this.state.userIdAusente).then((t) => { this.setState({ listViewTask: [] }); this.setState({ listViewTask: t }); return t; });//Obtener las tareas del usuario ausente y establecer en state las actividades para que se presenten en la tabla de tareas
      if (actividades.length == 0) {
        alert("Sin actividades por delegar");
        if (this.state.startDate == undefined || this.state.endDate == undefined) {
          alert("Sin fechas, para delegar actividades futuras elige fechas y delega nuevamente en especial para vacaciones");
        } else {
          alert("Las actividades futuras se reasignaran al delegado");
          this.saveAbsence();
        }
      } else {
        alert("Existen " + actividades.length + " actividades por delegar");
      }
      //actualizar la web
      window.location.href = window.location.href;
    } catch (e) { log(e); }//Captura errores
  }
  public changeOwnerOfCurrentAbsenceTasks = async () => {
    try {
      log("Cambiando propietario de tareas de ausencia");
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
      log("Finalizado Cambiando propietario de tareas de ausencia");
    } catch (e) { log(e); }//Captura errores
  }
  //funcion para guardar en la lista de ListForDelegationOfAbsences los datos de persona ausente, persona delegada, fecha de ausencia y fecha de regreso
  public saveAbsence = async () => {
    try {
      await this.sp.web.lists.getByTitle("ListForDelegationOfAbsences").items.add({//Objeto con los datos de la ausencia
        AusenteId: this.state.userIdAusente,//Id del usuario ausente
        AusenteStringId: this.state.userIdAusente.toString(),//Id del usuario ausente en string
        DelegadoId: this.state.userIdDelegado,//Id del usuario delegado
        DelegadoStringId: this.state.userIdDelegado.toString(),//Id del usuario delegado en string
        Inicio: this.state.startDate.toISOString(),//Fecha de inicio de la ausencia
        Fin: this.state.endDate.toISOString()//Fecha de fin de la ausencia
      });
    } catch (e) { log(e); }//Captura errores

  }
  public render(): React.ReactElement<any, any> {//Renderiza el componente
    const { userDisplayName } = this.props;//captura propiedades
    var jsx = (//html del componente
      <section className={styles.gestorDeAusencias}>
        <div className={styles.welcome}>
          <h3>Delegador de actividades y creador de ausencias</h3>
          <h4>Bienvenido señor(a) {escape(userDisplayName)}!</h4>
          <DateTimePicker label="Inicio de ausencia (Campo opcional)"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12}
            onChange={this.setStartDate} />
          <DateTimePicker label="Fin de ausencia (Campo opcional)"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12}
            onChange={this.setEndDate} />
          {this.state.showPeoplePickerAusente
            ? <p><PeoplePicker context={this.props.context}
              placeholder="Persona que delega o se ausentara..."
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
            resolveDelay={50} ref={c => (this.PeoplePickerDelegado = c)} /></p>
          <p>Actividades</p>
          {this.state.loading == false ?
            <ListView
              viewFields={[
                { name: "Title", displayName: "Actividad", isResizable: true, sorting: true, minWidth: 0, maxWidth: 150 },
                { name: "List", displayName: "Lista o Flujo De Tareas", isResizable: true, sorting: true, minWidth: 0, maxWidth: 150 }
              ]}
              filterPlaceHolder="Busque por lista, flujo o actividad..."
              selection={this._getSelectionOfListView}
              selectionMode={SelectionMode.multiple}
              items={this.state.listViewTask}
              dragDropFiles={false} stickyHeader={true}
              showFilter={true} defaultFilter="" compact={true}
            /> :
            "Cargando datos..."
          }
          <p>
            <button type="button" onClick={this.changeOwner}>Delegar Actividades</button>
          </p>
        </div>
      </section>
    );
    var url = window.location.href;//Obtiene la url actual
    if (
      url == "https://carvajal.sharepoint.com/sites/flujosprocesos" ||
      url == "https://carvajal.sharepoint.com/sites/flujosprocesos/" ||
      url == "https://carvajal.sharepoint.com/sites/FlujosdetrabajoCTSMX" ||
      url == "https://carvajal.sharepoint.com/sites/FlujosdetrabajoCTSMX/" ||
      url == "https://carvajal.sharepoint.com/sites/flujosprocesos/SitePages/Tareas.aspx" ||
      url == "https://carvajal.sharepoint.com/sites/flujosprocesos/SitePages/Tareas.aspx/"
    )
      jsx = (<div></div>);//Si la url actual es igual a la url a bloquear entonces no se muestra el componente pero muestra un campo vacio
    return jsx;
  }
}