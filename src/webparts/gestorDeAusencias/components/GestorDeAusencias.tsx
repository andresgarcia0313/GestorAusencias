import * as React from 'react';//Importar React para crear componentes react que son componentes web que se muestran en la pagina
import styles from './GestorDeAusencias.module.scss';//Importamos el css para darle estilo a la webpart o elemento web
import { escape } from '@microsoft/sp-lodash-subset';//importamos el paquete de escape para evitar inyecciones de código
import { spfi, SPFx } from "@pnp/sp";//importamos el paquete pnp para obtener las funciones de sharepoint
import "@pnp/sp/webs";//importamos el modulo de webs para poder obtener las webs de sharepoint
import "@pnp/sp/lists";//importamos el modulo de listas para poder obtener los datos de listas de sharepoint
import "@pnp/sp/items";//importamos el modulo de items para poder obtener los datos de items de sharepoint que son los datos de las listas
import "@pnp/sp/site-users/web";//importamos el modulo de usuarios para poder obtener los datos de usuarios de sharepoint
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';//importamos el modulo de llamadas de datos servicios web
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";//importamos el modulo de people picker para seleccionar usuarios
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
  }
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

  public getTasksFromTaskListsByUserId = async (userId: number) => {//Funcion que obtiene las tareas de las listas de tareas por el id del usuario

    var inicio: any = (new Date());  //Obtener momento de inicio

    var nameLists = (await this.props.context.spHttpClient.get(
      this.props.context.pageContext.web.absoluteUrl +
      `/_api/web/lists?$filter=(Hidden eq false) and (BaseTemplate eq 171)&$select=Title`,
      SPHttpClient.configurations.v1)
      .then(async (r: SPHttpClientResponse) => {
        return (await r.json()).value;
      }));//obtiene listas tipo tarea

    var _taskLists: any[] = new Array();//variable donde almacenara las promesas de obtener las tareas de las listas    

    for (let l of nameLists) _taskLists.push(this.sp.web.lists.getByTitle(l.Title).items());//recorre los nombres de las listas tipo tareas para obtener sus registros //Almacena las promesas en un array 

    var taskLists = await Promise.all(_taskLists);//Obtiene los registros de las listas tipo tarea
    //Cuando la lista de tareas encuentre su nombe asignar el nombre de la lista a cada tarea dentro de cada lista de tareas
    for (var i = 0; i < nameLists.length; i++)
      for (var x = 0; x < taskLists.length; x++)
        if (x == i)
          for (var task of taskLists[x])
            task.List = nameLists[i].Title;
    //Variable donde se almacenaran las tareas del usuario
    var taskByUserId: any[] = new Array();
    //Recorre tareas en cada lista de tareas para obtener las tareas incompletas por id usuario y las almacena en un array taskByUserId
    for (var taskList of taskLists) for (var task of taskList)
      if (task.PercentComplete < 1 && task.AssignedToId[0] == userId)
        taskByUserId.push(task);
    var fin: any = (new Date());
    console.log("Obtenido las tareas del Usuario en " + ((fin - inicio) / 1000) + "s");
    //Retorna las tareas del usuario
    return taskByUserId;
  }
  //Funcion para obtener la lista de elementos seleccionados en la lista de tareas
  private _getSelectionOfListView = (tasks: any[]) => {
    this.setState({ listSelectedViewTask: tasks });
  }
  public changeOwner = async () => {//metodo que cambia de propietario de la tarea por tarea, usuario ausente, usuario delegado
    try {
      var inicio: any = (new Date()); //Obtener momento de inicio      
      if (this.state.peoplePickerAusente != undefined) {
        var userIdAusente = (await this.sp.web.siteUsers.getByEmail(this.state.peoplePickerAusente[0].secondaryText)()).Id
      }
      else {
        var userIdAusente = (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id;
      }//obtener el id del usuario ausente del people picker o del usuario logueado
      var tasksToChangeOwner: any[] = new Array();//declarar array de promesas de cambio de propietario de tarea 
      for (var task of this.state.listSelectedViewTask) //Recorre las tareas seleccionadas para cambiar de propietario
        tasksToChangeOwner[task.Id] = this.sp.web.lists.getByTitle(task.List).items.getById(task.Id).update({ AssignedToId: [this.state.userIdDelegado], });//Cambia el propietario de la tarea
      await Promise.all(tasksToChangeOwner);//Espera a que se resuelvan todas las promesas de cambio de propietario de tarea para continuar
      var tasks = await this.getTasksFromTaskListsByUserId(userIdAusente);//Obtenemos las tareas del usuario que esta logueado
      this.setState({ listViewTask: tasks });//Guardamos las tareas en el estado
      var fin: any = (new Date()); console.log("Cambiado Asignación De Tareas  " + ((fin - inicio) / 1000) + "s");//Obtener momento de fin
    } catch (e) { console.log(e); }//Captura errores
  }
  public render(): React.ReactElement<any, any> {
    const { userDisplayName } = this.props;//captura propiedades
    return (
      <section className={styles.gestorDeAusencias}>
        <div className={styles.welcome}>
          <h3>Gestor y creador de ausencias</h3>
          <h4>Bienvenido señor(a) {escape(userDisplayName)}!</h4>          
          <DateTimePicker label="Inicio"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12} />
          <DateTimePicker label="Fin"
            dateConvention={DateConvention.DateTime}
            timeConvention={TimeConvention.Hours12} />
          {this.state.showPeoplePickerAusente ?
            <p><PeoplePicker context={this.props.context}
              placeholder="Persona a ausentarse..."
              personSelectionLimit={1} showtooltip={false} required={true}
              disabled={false} onChange={this.getPeoplePickerItemsAusente}
              showHiddenInUI={true} principalTypes={[PrincipalType.User]}
              resolveDelay={50} /> </p> : ""}
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
            },
            {
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
  }
}