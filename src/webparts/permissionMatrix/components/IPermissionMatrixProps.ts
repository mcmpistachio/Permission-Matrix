import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { PageContext } from "@microsoft/sp-page-context";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPermissionMatrixProps {
  group: string;
  context: WebPartContext;
}

//  export interface IPropertyControlsTestWebPartProps {
//    people: IPropertyFieldGroupOrPerson[];
//  }
