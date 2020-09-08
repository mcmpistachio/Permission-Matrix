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
import { WebPartContext } from "@microsoft/sp-webpart-base";


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
        <Link key={item.id} onClick={() => this._getFiles(item.id)}>
          {item.name}
        </Link>
      )
    }
  ];

  private _addcolumns(column: IColumn[]): void {
    this._apiUser('root').then(element => {
      if (element == null){
        this.setState({dispColumn:column});
      } else {
        for (let each of element){
          column.push({
            key: each.grantedTo.user.id,
            name: each.grantedTo.user.displayName,
            minWidth: 60,
            onRender: item => (
              <DropPermissionItem context={this.props.context} file={item} groupColumn={each}/>
            ),
          });
        }
        this.setState({apiColumn:column});
      }
    });
  }

  private _getFiles(parent:string):void {
    this.setState({apiFiles:loadItem()});
    this._apiFile(parent).then(element =>{
      let file = new Array<MicrosoftGraph.DriveItem>();
      for (let each of element){
        file.push(each);
      }
      console.log(file);
      this.setState({apiFiles:file});

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
        let apiUrl: string = '/groups/'+this.props.context.pageContext.site.group.id+'/drive/items/'+file+'/permissions';
        client
          .api(apiUrl)
          .version("v1.0")
          .get((error?, result?: any, rawResponse?: any): any => {
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

  private _apiFile(parent:string): any {
    return new Promise<any>((resolve, reject)=>{
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient):any => {
        let apiUrl: string = '/groups/'+this.props.context.pageContext.site.group.id+'/drive/items/'+parent+'/children';
        client
          .api(apiUrl)
          .version("v1.0")
          .get(async (error?, result?: any, rawResponse?: any) => {
            // handle the response
            if(error){
              console.error(error);
            }
            if (result) {
              let file = new Array<MicrosoftGraph.DriveItem>();
              for (let each of result.value){
                let eachItem: MicrosoftGraph.DriveItem = each;
                eachItem.permissions = await this._apiUser(eachItem.id);
                file.push(eachItem);
              }
              resolve(file);
            }
          }
        );
      });
    });
  }
}

function loadItem(): MicrosoftGraph.DriveItem[] {
  let fileLoad = new Array<MicrosoftGraph.DriveItem>();
  fileLoad.push({
    id: 'Loading',
    name: 'Loading'
  });
  return fileLoad;
}
