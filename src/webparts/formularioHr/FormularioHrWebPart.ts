import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "FormularioHrWebPartStrings";
import FormularioHr from "./components/FormularioHr";
import { IFormularioHrProps } from "./components/IFormularioHrProps";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

export interface IFormularioHrWebPartProps {
  description: string;
  DatosHR: any;
  lab: any;
}

export default class FormularioHrWebPart extends BaseClientSideWebPart<IFormularioHrWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IFormularioHrProps> = React.createElement(
      FormularioHr,
      {
        description: this.properties.description,
        context: this.context,
        DatosHR: this.properties.DatosHR,
        lab: this.properties.lab,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      spfi().using(SPFx(this.context));
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker("DatosHR", {
                  label: "Selecciona la lista de DatosHR",
                  selectedList: this.properties.DatosHR,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("lab", {
                  label: "Selecciona la lista de lab",
                  selectedList: this.properties.lab,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId1",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
