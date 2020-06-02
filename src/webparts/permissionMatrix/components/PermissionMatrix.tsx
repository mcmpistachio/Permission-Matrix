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
import { concatStyleSets, ThemeSettingName } from 'office-ui-fabric-react/lib/Styling';

export interface IDetailsListNavigatingFocusExampleState {
  items: string[];
  initialFocusedIndex?: number;
  parentFile: string;
  key: number;
  fileItems?: MicrosoftGraph.DriveItem[];
  userColumn?: MicrosoftGraph.Permission[];
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
      minWidth: 60,
      onRender: item => (
        // tslint:disable-next-line:jsx-no-lxambda
        <Link key={item} onClick={() => this._navigate(item)}>
          {item}
        </Link>
      )
    }
  ];

  private _addcolumns(column: IColumn[]): IColumn[] {
    let _loadColumn: MicrosoftGraph.Permission[] = this._loadUser();
    if (_loadColumn == null){
        column.push({
          key: 'permission',
          name: 'Permission',
          minWidth: 60,
          onRender: item => (<DropPermissionItem/>),
          }
        );
        return column;
      } else {
        _loadColumn.forEach(element =>{
          column.push({
            key: 'permission',
            name: element.grantedTo.user.displayName,
            minWidth: 60,
            onRender: item => (<DropPermissionItem/>),
        });
        });
        return column;
      }
  }

  private displayColumns: IColumn[] = this._addcolumns(this._columns);

  public render(): JSX.Element {
    // By default, when the list is re-rendered on navigation or some other event,
    // focus goes to the list container and the user has to tab back into the list body.
    // Setting initialFocusedIndex makes focus go directly to a particular item instead.

    return (
      <div>
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
      </div>
    );
  }

  private _navigate = (name: string) => {
    this.setState({
      items: generateItems(name + ' / '),
      initialFocusedIndex: 0,
      key: this.state.key + 1,
    });
  }

  private _loadUser(): MicrosoftGraph.Permission[] {
    let response: MicrosoftGraph.Permission[];
    this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient):any => {
      let apiUrl: string = '/groups/'+this.props.group+'/drive/items/root/permissions';
      client
        .api(apiUrl)
        .version("v1.0")
        .get((error?, result?: any, rawResponse?: any):any => {
          // handle the response
          if(error){
            console.error(error);
          }
          if (result) {
            console.log("Reached the Graph");
            console.log(result.value.length);
            result.value.forEach(element =>{
              console.log(element.grantedTo.user.displayName);
            });
          }
        }
      );
    });
    return response;
  }
  private _apiUser():MicrosoftGraph.Permission[] {
    let apiUser = new Promise<any>((resolve, reject)=>{
      this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient):any => {
        let apiUrl: string = '/groups/'+this.props.group+'/drive/items/root/permissions';
        client
          .api(apiUrl)
          .version("v1.0")
          .get((error?, result?: any, rawResponse?: any):any => {
            // handle the response
            if(error){
              console.error(error);
            }
            if (result) {
              resolve(result)
            }
          }
        );
      });
    })
    var _mappedUser: MicrosoftGraph.Permission[];
    apiUser.value.map((item: any)=>{
      _mappedUser.push({
        id: item.id,
        roles: item.roles,
        grantedTo: item.grantedTo
      })
    })
    return ;
  }

  private _loadFiles(): MicrosoftGraph.DriveItem[] {
    let driveFile: MicrosoftGraph.DriveItem[];
    this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): any => {
      let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+this.state.parentFile+'/children';
      client
        .api(apiUrl)
        .version("v1.0")
        .get((err, res: MicrosoftGraph.DriveItem[]) => {
          // handle the response
          if(err){
            console.error(err);
          } else {
            return res;
          }
        }
      );
    });
    return driveFile;
  }
}


function generateItems(parent: string): string[] {
  return Array.prototype.map.call('ABCDEFGHI', (name: string) => parent + 'Folder ' + name);
}
