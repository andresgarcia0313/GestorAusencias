import * as React from 'react';
import styles from './Componente.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListView, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/DateTimePicker';
//componente de ejemplo de react spfx
export default class Component extends React.Component<any, any> {
    render() {
        return (
            <div className="ms-Grid-row">   
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                    <h1>Componente de ejemplo</h1>
                </div>
            </div>
        );
    }
}