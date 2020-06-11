import * as React from 'react';
import styles from './PermissionMatrix.module.scss';
import { IPermissionMatrixWebPartProps } from '../PermissionMatrixWebPart';
import { escape } from '@microsoft/sp-lodash-subset';
import { DropPermissionItem, IDropdownPermissionProps } from './DropPermission';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton } from 'office-ui-fabric-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { concatStyleSets, ThemeSettingName } from 'office-ui-fabric-react/lib/Styling';


export interface IDetailsListNavigatingFocusExampleState {
  initialFocusedIndex?: number;
  key: number;
  apiColumn?: IColumn[];
  apiFiles?: MicrosoftGraph.DriveItem[];
}


export default class PermissionMatrix extends React.Component<IPermissionMatrixWebPartProps, {}> {
  public state: IDetailsListNavigatingFocusExampleState = {
    key: 0,
    apiFiles: loadItem(),
  };

  private _columns: IColumn[] = [
    {
      key: 'file',
      name: "File",
      minWidth: 60,
      onRender: item => (
        // tslint:disable-next-line:jsx-no-lambda
        <Link key={item.id} onClick={() => this._getFiles(item.id)}>
          {item.name}
        </Link>
      )
    }
  ];

  private _addcolumns(column: IColumn[]): void {
    this._apiUser('root').then(element => {
      // console.log(element.length);
      // console.log(element);
      if (element == null){
        // column.push({
        //   key: 'permission',
        //   name: 'Permission',
        //   minWidth: 60,
        //   onRender: item => (<DropPermissionItem />),
        //   }
        // );
        this.setState({dispColumn:column});
      } else {
        for (let each of element){
          let perm:MicrosoftGraph.Permission = each;
          column.push({
            key: each.grantedTo.user.id,
            name: each.grantedTo.user.displayName,
            minWidth: 60,
            onRender: item => (<DropPermissionItem column={perm} context={this.props.context} group={this.props.group} item={item}/>),
          });
        }
        this.setState({apiColumn:column});
      }
    });
  }

  private _getFiles(parent:string):void {
    this._apiFiles(parent).then(element =>{
      // console.log(element);
      let file = new Array<MicrosoftGraph.DriveItem>();
      for (let each of element){
        file.push(each);
      }
      this.setState({apiFiles:file});
      console.log(this.state.apiFiles);
    });
  }

  public componentDidMount() {
    this._addcolumns(this._columns);
    this._getFiles('root');
  }

  public render(): JSX.Element {
    // By default, when the list is re-rendered on navigation or some other event,
    // focus goes to the list container and the user has to tab back into the list body.
    // Setting initialFocusedIndex makes focus go directly to a particular item instead.

    return (
      <div>
        <PrimaryButton text="Home" onClick={()=> this._getFiles('root')}/>
        <DetailsList
        key={this.state.key}
        items={this.state.apiFiles}
        columns={this.state.apiColumn}
        initialFocusedIndex={this.state.initialFocusedIndex}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="Row checkbox"
      />
      </div>
    );
  }

  //Promising the API result
  private _apiUser(file:string): any {
    return new Promise<any>((resolve, reject)=>{
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient):any => {
        let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+file+'/permissions';
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

  private _apiFiles(parent:string): any {
    return new Promise<any>((resolve, reject)=>{
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient):any => {
        let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+parent+'/children';
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


function loadItem():MicrosoftGraph.DriveItem[]{
  let item = new Array<MicrosoftGraph.DriveItem>();
  item.push({
      id: 'Loading',
      name: 'Loading'

  });
  return item;
}
