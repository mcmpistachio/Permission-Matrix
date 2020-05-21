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
  fileItems?: FileItem[];
  userColumn?: permUser[];
}
export interface permUser {
  displayName: string;
  objectID: string;
}
export interface FileItem {
  displayName: string;
  objectID: string;
}

export default class PermissionMatrix extends React.Component<IPermissionMatrixWebPartProps, {}> {
  public state: IDetailsListNavigatingFocusExampleState = {
    items: generateItems(''),
    parentFile: "root",
    key: 0,
  };

  private _loadGroups(): any {
    this.props.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {

      let apiUrl: string = '/groups/'+this.props.group+'/drive/items/root/permissions';

      client
        .api(apiUrl)
        .version("v1.0")
        .get((err, res: any) => {
          // handle the response
          if(err){
            console.error(err);
            return err;
          }

            let group: Array<permUser> = new Array<permUser>();

              res.value.map((item:MicrosoftGraph.Permission) => {
                group.push({
                  displayName: item.grantedTo.user.displayName,
                  objectID: item.grantedTo.user.id
                });
              });
            this.setState({userColumn: group});
          }
      );
    });
  }

  private _loadFiles(): any {
    this.props.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient): void => {

      let apiUrl: string = '/groups/'+this.props.group+'/drive/items/'+this.state.parentFile+'/children';

      client
        .api(apiUrl)
        .version("v1.0")
        .get((err, res) => {
          // handle the response
          if(err){
            console.error(err);
            return err;
          }

            let file: Array<FileItem> = new Array<FileItem>();

              res.value.map((item:MicrosoftGraph.DriveItem) => {
                file.push({
                  displayName: item.name,
                  objectID: item.id
                });
              });
              this.setState({fileItems: file});
          }
      );
    });
  }

  private _columns: IColumn[] = [
    {
      key: 'file',
      name: 'File',
      minWidth: 90,
      onRender: item => (
        // tslint:disable-next-line:jsx-no-lxambda
        <Link key={item} onClick={() => this._navigate(item)}>
          {item}
        </Link>
      ),
    },
    {
      key: 'permission',
      name: 'permission',
      minWidth: 60,
      onRender: item => (<DropPermissionItem/>),
    }
  ];

  private _addcolumns(column: IColumn[]): IColumn[] {
      if (this.props.people == null){
        return column;
      } else {

        for (let user of this.props.people) {
          column.push({
              key: 'permission',
              name: user.fullName,
              minWidth: 60,
              onRender: item => (<DropPermissionItem/>),
            }
          );
        }
        return column;
      }
  }

  private displayColumns: IColumn[] = this._addcolumns(this._columns);

    public componentDidMount(): void{
      this._loadGroups();
      this._loadFiles();
    }
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
}


function generateItems(parent: string): string[] {
  return Array.prototype.map.call('ABCDEFGHI', (name: string) => parent + 'Folder ' + name);
}
