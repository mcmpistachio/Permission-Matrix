import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDynamicField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'PermissionMatrixWebPartStrings';
import PermissionMatrix from './components/PermissionMatrix';
import { IPermissionMatrixProps } from './components/IPermissionMatrixProps';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { selectProperties } from 'office-ui-fabric-react/lib/Utilities';
import PropertyFieldPeoplePickerHost from '@pnp/spfx-property-controls/lib/propertyFields/peoplePicker/PropertyFieldPeoplePickerHost';
import { CompactPeoplePicker } from 'office-ui-fabric-react/lib/Pickers';

export interface IPermissionMatrixWebPartProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  groups: string[];
}

export default class PermissionMatrixWebPart extends BaseClientSideWebPart <IPermissionMatrixProps> {

  public render(): void {
    const element: React.ReactElement<IPermissionMatrixProps> = React.createElement(
      PermissionMatrix,
      {
        description: this.properties.description,
        people: this.properties.people,
        groups: this.properties.groups
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Select Groups to Grant Permissions',
              groupFields: [
                PropertyFieldPeoplePicker('people', {
                  label: 'PropertyFieldPeoplePicker',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
