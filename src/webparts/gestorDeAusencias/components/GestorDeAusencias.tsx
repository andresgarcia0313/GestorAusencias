import * as React from 'react';
import styles from './GestorDeAusencias.module.scss';
import { IGestorDeAusenciasProps } from './IGestorDeAusenciasProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class GestorDeAusencias extends React.Component<IGestorDeAusenciasProps, {}> {
  public render(): React.ReactElement<IGestorDeAusenciasProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

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
        </div>
      </section>
    );
  }
}
