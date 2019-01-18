import * as React from 'react';
import styles from './MsGraphCalls.module.scss';
import { IMsGraphCallsProps } from './IMsGraphCallsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http';
import { IUsersResponse } from './IUsersReponse';
import { IUser } from './IUser';

import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn, IObjectWithKey } from 'office-ui-fabric-react/lib/DetailsList';
import { IMsGraphCallsState } from './IMsGraphCallsState';

export default class MsGraphCalls extends React.Component<IMsGraphCallsProps, IMsGraphCallsState> {
  private _selection: Selection;
  public state: IMsGraphCallsState;
  constructor(props: IMsGraphCallsProps) {
    super(props);

    const columns: IColumn[] = [
      {
        key: 'displayName',
        name: 'Full Name',
        fieldName: 'displayName',
        minWidth: 300,
      },
      {
        key: 'mail',
        name: 'E-mail',
        fieldName: 'mail',
        minWidth: 300,
      },
      {
        key: 'userPrincipalName',
        name: 'Login name',
        fieldName: 'userPrincipalName',
        minWidth: 300,
      }
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        debugger;
        this._setSelection()
      }
    });
    this.state = {
      columns: columns,
      items: null,
      selectedItems: null,
      //selectionDetails: this._getSelectionDetails(),
    };
  }

  private _setSelection(): void {
    debugger;
    const selectionCount = this._selection.getSelectedCount();

    const selection: IObjectWithKey[] = this._selection.getSelection();

    this.setState({
      selectedItems: this._selection.getSelection() as IUser[]
    })
  }




  public render(): React.ReactElement<IMsGraphCallsProps> {


    return (
      <div className={styles.msGraphCalls}>
        <button type="button" onClick={this.getUsersFromO365.bind(this)}>Get Users From O365</button>

        {this.state.items == null ? null :
          <DetailsList
            items={this.state.items}
            //compact={isCompactMode}
            columns={this.state.columns}
            //selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            //onItemInvoked={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />





        }



        <button type="button" onClick={this.getUserFromO365.bind(this)}>Get Dusan From O365</button>
      </div>
    );
  }


  private async getUsersFromO365() {
    //debugger;
    const msGraphClient: MSGraphClient = await this.props.spfxContext.msGraphClientFactory.getClient();


    const allUsers: IUsersResponse = await msGraphClient
      .api('/users')
      .get((error, response: any, rawResponse?: any) => {
        //debugger;
        this.setState({
          items: response.value
        })
        // handle the response
      });



  }

  private async getUserFromO365() {
    //debugger;
    const msGraphClient: MSGraphClient = await this.props.spfxContext.msGraphClientFactory.getClient();

    const dusan: IUser = await msGraphClient
      .api('/users/b3cdc539-a8d1-47ab-a010-2ddb5aafaec3')
      .get((error, response: any, rawResponse?: any) => {
        //debugger;
        // handle the response
      });


  }
}
