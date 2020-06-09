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
  apiColumn?: IColumn[];
  apiFiles?: MicrosoftGraph.DriveItem[];
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
      fieldName: 'name'
    }
  ];

  private _addcolumns(column: IColumn[]): void {
    this._apiUser().then(element => {
      // console.log(element.length);
      // console.log(element);
      if (element == null){
        column.push({
          key: 'permission',
          name: 'Permission',
          minWidth: 60,
          onRender: item => (<DropPermissionItem/>),
          }
        );
        this.setState({dispColumn:column});
      } else {
        for (let each of element){
          column.push({
            key: each.grantedTo.user.id,
            name: each.grantedTo.user.displayName,
            minWidth: 60,
            onRender: item => (<DropPermissionItem/>),
          });
        }
        this.setState({apiColumn:column});
      }
    });
  }

  //Need to finish!!! Handle the API push to state
  private _getFiles():void {
    this._apiFiles().then(element =>{
      console.log(element);
      if (element == null) {
        //what to do?
      } else {
        let file:MicrosoftGraph.DriveItem[];
        for (let each of element){
          file.push(each);
        }
        this.setState({apiFiles:file});
      }
    });
  }

  public componentDidMount() {
    this._addcolumns(this._columns);
    this._getFiles();
  }

  public render(): JSX.Element {
    // By default, when the list is re-rendered on navigation or some other event,
    // focus goes to the list container and the user has to tab back into the list body.
    // Setting initialFocusedIndex makes focus go directly to a particular item instead.

    return (
      <div>
        <DetailsList
        key={this.state.key}
        items={this.state.apiFiles}
        columns={this.state.apiColumn}
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

  //Promising the API result
  private _apiUser(): any {
    return new Promise<any>((resolve, reject)=>{
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
              resolve(result.value);
            }
          }
        );
      });
    });
  }

  private _apiFiles(): any {
    return new Promise<any>((resolve, reject)=>{
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient):any => {
        let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+this.state.parentFile+'/children';
        client
          .api(apiUrl)
          .version("v1.0")
          .get((error?, result?: any, rawResponse?: any):any => {
            // handle the response
            if(error){
              console.error(error);
            }
            if (result) {
              resolve(result.value);
            }
          }
        );
      });
    });
  }
}


function generateItems(parent: string): string[] {
  return Array.prototype.map.call('ABCDEFGHI', (name: string) => parent + 'Folder ' + name);
}
