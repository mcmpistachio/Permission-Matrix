import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { PageContext } from "@microsoft/sp-page-context";

export interface IPermissionMatrixProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  group: string;
}

//  export interface IPropertyControlsTestWebPartProps {
//    people: IPropertyFieldGroupOrPerson[];
//  }
