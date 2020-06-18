import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDynamicField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import * as strings from 'PermissionMatrixWebPartStrings';
import PermissionMatrix from './components/PermissionMatrix';
import { IPermissionMatrixProps } from './components/IPermissionMatrixProps';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { selectProperties } from 'office-ui-fabric-react/lib/Utilities';

export interface IPermissionMatrixWebPartProps {
  context: WebPartContext;
}

export default class PermissionMatrixWebPart extends BaseClientSideWebPart <IPermissionMatrixProps> {

  public render(): void {
    const element: React.ReactElement<IPermissionMatrixProps> = React.createElement(
      PermissionMatrix,
      {
        context: this.context
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
    return ;
  }
}
