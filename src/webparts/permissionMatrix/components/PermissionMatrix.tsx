import * as React from 'react';
import styles from './PermissionMatrix.module.scss';
import { IPermissionMatrixWebPartProps } from '../PermissionMatrixWebPart';
import { escape } from '@microsoft/sp-lodash-subset';
import DropPermissionItem from './DropPermission';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface IDetailsListNavigatingFocusExampleState {
  items: string[];
  initialFocusedIndex?: number;
  parentFile: string;
  key: number;
  fileItems?: MicrosoftGraph.DriveItem[];
  userColumn?: MicrosoftGraph.Permission[];
}
export interface defColumn {
  id: string;
  displayName: string;
}

export default class PermissionMatrix extends React.Component<IPermissionMatrixWebPartProps, {}> {
  public state: IDetailsListNavigatingFocusExampleState = {
    items: generateItems(''),
    parentFile: "root",
    key: 0,
  };

  private _columns: IColumn[] = [
    {
      key: 'file',
      name: "File",
      minWidth: 10,
      onRender: item => (
        // tslint:disable-next-line:jsx-no-lxambda
        <Link key={item} onClick={() => this._navigate(item)}>
          {item}
        </Link>
      ),
    }
  ];

  private _addcolumns(column: IColumn[]): IColumn[] {
    if (this.state.userColumn == null){
        // return column;
        this._loadUser();
      } else
      {
        for (let col of this.state.userColumn) {
          column.push({
              key: col.id,
              name: col.grantedTo.user.displayName,
              minWidth: 10,
              onRender: item => <DropPermissionItem/>,
            }
          );
        }
        return column;
      }
  }

  private displayColumns: IColumn[] = this._addcolumns(this._columns);

  public render(): JSX.Element {
    // By default, when the list is re-rendered on navigation or some other event,
    // focus goes to the list container and the user has to tab back into the list body.
    // Setting initialFocusedIndex makes focus go directly to a particular item instead.

    return (
      <DetailsList
        key={this.state.key}
        items={this.state.items}
        columns={this.displayColumns}
        onItemInvoked={this._navigate}
        initialFocusedIndex={this.state.initialFocusedIndex}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="Row checkbox"
      />
    );
  }

  private _navigate = (name: string) => {
    this.setState({
      items: generateItems(name + ' / '),
      initialFocusedIndex: 0,
      key: this.state.key + 1,
    });
  }
  private _loadUser(): void {
    this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      let apiUrl: string = '/groups/'+this.props.group+'/drive/items/root/permissions';
      client
        .api(apiUrl)
        .version('v1.0')
        .get((error, result: MicrosoftGraph.Permission[]) => {
          // handle the response
          if(error){
            console.error(error);
          } else {
            // result.forEach(element => {

            // });
            this.setState({userColumn: result});
          }}
      );
    });
  }
  private _loadFiles(): void {
    let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+this.state.parentFile+'/children';
    this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      client
        .api(apiUrl, )
        .version("v1.0")
        .get((error, response: MicrosoftGraph.DriveItem) => {
          // handle the response
          if(error){
            console.error(error);
          } else {
            this.setState({fileItems: response});
          }}
      );
    });
  }
}


function generateItems(parent: string): string[] {
  return Array.prototype.map.call('ABCDEFGHI', (name: string) => parent + 'Folder ' + name);
}
