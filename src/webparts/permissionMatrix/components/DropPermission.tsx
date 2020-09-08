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
  perm?: MicrosoftGraph.Permission;
  groupColumn?: MicrosoftGraph.Identity;
}

export interface IDropdownPermissionState {
  role?: string[];
}

const exampleOptions: IDropdownOption[] = [
  { key: 'noaccess', text: 'No Access', data: { icon: 'Blocked12' } },
  { key: 'read', text: 'Read', data: { icon: 'View' } },
  { key: 'write', text: 'Write', data: { icon: 'Edit' } },
  { key: 'owner', text: 'Owner', data: { icon: 'Teamwork' } }
];


export class DropPermissionItem extends React.Component<IDropdownPermissionProps,{}> {
  public state: IDropdownPermissionState = {
    // role: defaultRole(this.props.group.toString(), this.props.item.id.toString(), this.props.column.grantedTo.user.id.toString(), this.props.context),
    role: this.dropDownDefault(),
  };

  public componentDidMount(){
    this._dropDownDefault();
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

  private dropDownDefault(): string[] {
    if (this.props.file.permissions == null){
      // console.log(this.props.file.name);
      return ['noaccess'];
    } else {
      for (let each of this.props.file.permissions){
        // console.log(this.props.file.name);
        if (each.grantedTo.user.displayName == this.props.groupColumn.displayName){
          let role = new Array<string>();
          role.concat(each.roles);
          console.log(each.roles);
          return each.roles;
        }
      }
    }
  }

  private _dropDownDefault(): void {
    if (this.props.file.permissions == null){
      this.setState({role:['noaccess']});
    } else {
      for (let each of this.props.file.permissions){
        if (each.grantedTo.user.displayName == this.props.groupColumn.displayName){
          this.setState({role:each.roles});
        }
      }
    }
  }
}
