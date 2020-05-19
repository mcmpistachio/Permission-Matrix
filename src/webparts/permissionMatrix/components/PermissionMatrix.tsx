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
import { getItemStyles } from 'office-ui-fabric-react/lib/components/ContextualMenu/ContextualMenu.classNames';

export interface IDetailsListNavigatingFocusExampleState {
  items: string[];
  initialFocusedIndex?: number;
  parentFile: string;
  key: number;
}

export default class PermissionMatrix extends React.Component<IPermissionMatrixWebPartProps, {}> {
  public state: IDetailsListNavigatingFocusExampleState = {
    items: generateItems(''),
    parentFile: "root",
    key: 0,
  };

  private _columns: IColumn[] = [
    {
      key: 'name',
      name: 'File',
      minWidth: 90,
      onRender: item => (
        // tslint:disable-next-line:jsx-no-lxambda
        <Link key={item} onClick={() => this._navigate(item)}>
          {item}
        </Link>
      ),
    }
  ];

  private _addcolumns(column: IColumn[]): IColumn[] {
      if (this.props.people == null){
        return column;
      } else{
        for (let user of this.props.people) {
          column.push({
              key: 'permissionset',
              name: 'Permission',
              minWidth: 60,
              onRender: item => (<DropPermissionItem/>),
            }
          );
        }
        return column;
      }
  }

  private displayColumns: IColumn[] = this._addcolumns(this._columns);

  private getGroupID: string = this.context;

  public render(): JSX.Element {
    // By default, when the list is re-rendered on navigation or some other event,
    // focus goes to the list container and the user has to tab back into the list body.
    // Setting initialFocusedIndex makes focus go directly to a particular item instead.

    return (
      <div>
        <Label>{escape(this.props.group)}</Label>
        <DetailsList
          key={this.state.key}
          items={getLibraryItems(this.state.parentFile)}
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
}

function generateItems(parent: string): string[] {
  return Array.prototype.map.call('ABCDEFGHI', (name: string) => parent + 'Folder ' + name);
}

function getLibraryItems(parent:string): any {
  this.context.msGraphClientFactory
  .getClient()
  .then((client: MSGraphClient): void => {
    client
      .api("/groups/"+this.props.group+"/drive/items/"+parent+"/children")
      .get((error, response: MicrosoftGraph.DriveItem[], rawResponse?: any) => {
        // handle the response
        if(error){
          return generateItems('');
        } else if (response){
            var file: Array<MicrosoftGraph.DriveItem> = new Array<MicrosoftGraph.DriveItem>();

            response.map((item:MicrosoftGraph.DriveItem) => {
              file.push({
                id: item.id,
                name: item.name,
                children: item.children
              });
            });
          return file;
        }
    });
  });
}
