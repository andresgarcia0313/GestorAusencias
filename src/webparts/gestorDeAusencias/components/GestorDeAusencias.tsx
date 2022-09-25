import * as React from 'react';
import styles from './GestorDeAusencias.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";//importamos el modulo de webs
import "@pnp/sp/lists";//importamos el modulo de listas
import "@pnp/sp/items";//importamos el modulo de items
import "@pnp/sp/site-users/web";//importamos el modulo de usuarios
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';//importamos el modulo de llamadas de datos servicios web
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListView, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
export default class GestorDeAusencias extends React.Component<any, any> {//Clase GestorDeAusencias que hereda de React.Component
  private sp = spfi().using(SPFx(this.props.context));//inicializa libreria pnpjs
  constructor(props: any) {
    super(props);//ejecuta el constructor de la clase padre
    this.state = { showHidePeoplePickerAusente: false }//Estable la variable estado del componente    
  }
  public async componentDidMount(): Promise<void> {//Se ejecuta después de que el componente react se monta o se muestra
    console.log("componentDidMount");
    var group = (await this.sp.web.currentUser.groups()).filter(g => g.Title == "Propietarios del sitio")[0];//Identifica si el usuario pertenece al grupo "Propietarios del sitio"
    debugger;
    /*let groups = await this.sp.web.currentUser.groups();//Obtener grupos del usuario que inicio sesión
    for (let group of groups) {//Recorre los grupos del usuario
      console.log("Grupo:" + group.Title); // Mostrar en consola los titulos de los grupos del usuario que inicio sesión
      debugger;//Pausa la ejecución del código
      if (group.Title == "") {//Si el grupo es igual a "Gestor de Ausencias"
        this.setState({ showHidePeoplePickerAusente: true }); //Muestra el PeoplePicker
      }//Fin del if
    }//Fin del for
    */
    var userId = (await this.sp.web.siteUsers.getByEmail(this.props.user.email)()).Id;//obtiene el id del usuario de contexto instanciadose y ejecutandose en el acto
    var userTask = await this.getTasksFromTaskListsByUserId(userId);//Obtener tareas del usuario
    this.setState({ items: userTask });//Establecer en state las actividades para que se presenten en la tabla de tareas       
  }//Fin del método componentDidMount
  private getPeoplePickerItems = async (items: any[]) => {//Obtiene los elementos del selector de personas
    this.setState({ items: [] });//Borra el listado de tareas del state y a su vez de la tabla de tareas que ve el usuario
    if (items.length >= 1) {//Validar que exista una persona seleccionada de
      var userId = (await this.sp.web.siteUsers.getByEmail(items[0].secondaryText)()).Id//obtener el id del usuario
      var userTask = await this.getTasksFromTaskListsByUserId(userId);// Obtener tareas del usuario
      this.setState({ items: userTask });// Establecer en state las actividades para que se presenten en la tabla de tareas
    }
  }
  public getTasksFromTaskListsByUserId = async (userId: number) => {
    var inicio: any = (new Date());//Obtener momento de inicio
    console.log("Iniciado Obtener Tareas Del Usuario");//Muestre el metodo que se ejecuta  
    var lists = (await this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=(Hidden eq false) and (BaseTemplate eq 171)&$select=Title`, SPHttpClient.configurations.v1)
      .then(async (r: SPHttpClientResponse) => { return (await r.json()).value; }));//obtiene listas tipo tarea
    var promiseTaskLists: any[] = new Array();//variable donde almacenara las promesas de obtener las tareas de las listas    
    for (let list of lists) promiseTaskLists.push(this.sp.web.lists.getByTitle(list.Title).items());//recorre los nombres de las listas tipo tareas para obtener sus registros //Almacena las promesas en un array 
    var taskLists = await Promise.all(promiseTaskLists);//Espera que todas las promesas de listas de tareas almacenen datos
    for (var numList = 0; numList < lists.length; numList++)//Para cada nombre de lista haga AGREGAR A CADA TAREA EL NOMBRE DE LA LISTA A LA CUAL PERTENECE
      for (var numTaskList = 0; numTaskList < taskLists.length; numTaskList++) //Para cada lista de tareas haga
        if (numTaskList == numList) // Donde cada nombre de lista con su respectiva lista haga
          for (var task of taskLists[numTaskList]) //Para cada tarea de cada lista haga
            task.List = lists[numList].Title; //Agregale a la tarea el nombre de la lista a la cuál pertenece
    var taskByUserId: any[] = new Array();//Array donde almacenara todas las tareas del usuario autenticado
    for (var taskList of taskLists) // Para cada lista de tareas //CAPTURAR EN UN ARRAY LAS TAREAS DEL USUARIO
      for (var task of taskList) // Para cada tarea dentro de cada lista de tareas
        if (task.PercentComplete < 1 && task.AssignedToId[0] == userId)// Si esta completada y pertenece al usuario logueado
          taskByUserId.push(task)//Agregue todas las tareas que cumple la condicion al array creado
    var fin: any = (new Date());//Registrar finalizado del proceso
    console.log("Finalizado Obtener Tareas Del Usuario:" + ((fin - inicio) / 1000) + "s"); // Mostrar duración de ejecución
    return taskByUserId;// Retorna las tareas del usuario
  }
  public render(): React.ReactElement<any, any> {
    const { description, isDarkTheme, hasTeamsContext, userDisplayName, context } = this.props;//Establecen las propiedades del componente react
    return (
      <section className={`${styles.gestorDeAusencias} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Bienvenido Usuario: {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <h3>El presente formulario es el Gestor De Ausencias</h3>
          <p>Aquí usted delega actividades a otra persona en caso de ausentarse.</p>
          <p><PeoplePicker context={this.props.context}
            titleText="Persona a ausentar" personSelectionLimit={1}
            showtooltip={false} required={true} disabled={false}
            onChange={this.getPeoplePickerItems} showHiddenInUI={false}
            principalTypes={[PrincipalType.User]} resolveDelay={1000} />
          </p>
          <p><PeoplePicker context={this.props.context}
            titleText="Persona a delegar actividades"
            personSelectionLimit={1} showtooltip={false} required={true}
            disabled={false} onChange={this.getPeoplePickerItems}
            showHiddenInUI={false} principalTypes={[PrincipalType.User]}
            resolveDelay={1000} />
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
                name: "Title", displayName: "Actividad",
                isResizable: true, sorting: true, minWidth: 0, maxWidth: 150
              },
              {
                name: "List", displayName: "Lista",
                isResizable: true, sorting: true, minWidth: 0, maxWidth: 150
              }
            ]}
            iconFieldName="ServerRelativeUrl" compact={true}
            selectionMode={SelectionMode.multiple} showFilter={true}
            defaultFilter="" filterPlaceHolder="Buscar..."
            dragDropFiles={true} stickyHeader={true}
          />
          <p>
            <button type="button">Guardar y Generar Ausencia</button>
          </p>
        </div>
      </section>
    );
  }
}
