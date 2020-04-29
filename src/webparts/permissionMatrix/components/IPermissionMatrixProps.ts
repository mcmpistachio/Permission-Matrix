import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IPermissionMatrixProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  groups: string[];
}

//  export interface IPropertyControlsTestWebPartProps {
//    people: IPropertyFieldGroupOrPerson[];
//  }
