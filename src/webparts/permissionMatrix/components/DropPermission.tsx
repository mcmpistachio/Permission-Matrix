import * as React from 'react';
import { IPermissionMatrixWebPartProps } from '../PermissionMatrixWebPart';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDropdownPermissionProps {
  context: WebPartContext;
  file: MicrosoftGraph.DriveItem;
  perm: MicrosoftGraph.Permission[];
  group: MicrosoftGraph.Identity;
}

export interface IDropdownPermissionState {
  role?: string[];
}

const exampleOptions: IDropdownOption[] = [
  { key: 'noaccess', text: 'No Access', data: { icon: 'Blocked12' } },
  { key: 'read', text: 'Read', data: { icon: 'View' } },
  { key: 'write', text: 'Write', data: { icon: 'Edit' } }
];


export class DropPermissionItem extends React.Component<IDropdownPermissionProps,{}> {
  public state: IDropdownPermissionState = {
    // role: defaultRole(this.props.group.toString(), this.props.item.id.toString(), this.props.column.grantedTo.user.id.toString(), this.props.context),
    role: ['noaccess'],
  };

  public componentDidMount(){
    this.dropDownDefault();
  }

  public render(): JSX.Element {
    return (
      <Dropdown
      defaultSelectedKey={this.state.role}
      onRenderTitle={this._onRenderTitle}
      onRenderOption={this._onRenderOption}
      styles={{ dropdown: { width: 60 } }}
      options={exampleOptions}
    />
    );
  }

  private _onRenderOption = (option: IDropdownOption): JSX.Element => {
    return (
      <div>
        {option.data && option.data.icon && (
          <Icon style={{ marginRight: '8px' }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
        )}
      </div>
    );
  }

  private _onRenderTitle = (options: IDropdownOption[]): JSX.Element => {
    const option = options[0];

    return (
      <div>
        {option.data && option.data.icon && (
          <Icon style={{ marginRight: '8px' }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
        )}
      </div>
    );
  }

  private dropDownDefault(){
    for (let each of this.props.perm){
      if (each.grantedTo.user.id == this.props.group.id){
        this.setState({role:each.roles});
      }
    }
  }

  // private _getPermission(){
  //   this._filePermission(this.props.item.id).then(element =>{
  //     for (let each of element){
  //       if (each.grantedTo.user.displayName == this.props.column.grantedTo.user.displayName){
  //         this.setState({
  //           role: each.roles
  //         });
  //       }
  //     }
  //   });
  // }

  // private _filePermission(file:string): any {
  //   return new Promise<any>((resolve, reject)=>{
  //   this.props.context.msGraphClientFactory
  //     .getClient()
  //     .then((client: MSGraphClient):any => {
  //       let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+file+'/permissions';
  //       client
  //         .api(apiUrl)
  //         .version("v1.0")
  //         .get((error?, result?: any, rawResponse?: any):any => {
  //           // handle the response
  //           if(error){
  //             console.error(error);
  //           }
  //           if (result) {
  //             let response: any;
  //             for (let each of result.value){
  //               if (each.grantedTo.user.id == this.props.column.grantedTo.user.id){
  //                 response = each.roles;
  //               }
  //             }
  //             if (response == null){
  //               response = ['noaccess'];
  //             }
  //             console.log(response);
  //             resolve(response);
  //           }
  //         }
  //       );
  //     });
  //   });
  // }
}

// function defaultRole(groupID:string, file:string, columnID:string, property: WebPartContext): any {
//   return new Promise<any>((resolve, reject)=>{
//     property.msGraphClientFactory
//       .getClient()
//       .then((client: MSGraphClient):any => {
//         let apiUrl: string = '/groups/'+groupID+'/drive/items/'+file+'/permissions';
//         client
//           .api(apiUrl)
//           .version("v1.0")
//           .get((error?, result?: any, rawResponse?: any):any => {
//             // handle the response
//             if(error){
//               console.error(error);
//             }
//             if (result) {
//               let response: any;
//               for (let each of result.value){
//                 if (each.grantedTo.user.id == columnID){
//                   console.log(each.grantedTo.user.displayName);
//                   console.log(each.roles);
//                   response = each.roles;
//                 }
//               }
//               if (response == null){
//                 response = ['noaccess'];
//               }
//               console.log(response);
//               resolve(response);
//             }
//           }
//         );
//       });
//     });
//   }
