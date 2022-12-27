import BuscaroHr from "./components/BuscaroHr";
import { IBuscaroHrProps } from "./components/IBuscaroHrProps";

import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "BuscaroHrWebPartStrings";

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";

export interface IBuscaroHrWebPartProps {
  description: string;
  docType: any;
  Ano: any;
  DatosHR: any;
}

export default class BuscaroHrWebPart extends BaseClientSideWebPart<IBuscaroHrWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IBuscaroHrProps> = React.createElement(
      BuscaroHr,
      {
        description: this.properties.description,
        context: this.context,
        docType: this.properties.docType,
        Ano: this.properties.Ano,
        DatosHR: this.properties.DatosHR,
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
                PropertyFieldListPicker("docType", {
                  label: "Selecciona la lista de Tipos de documento",
                  selectedList: this.properties.docType,
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
                PropertyFieldListPicker("Ano", {
                  label: "Selecciona la lista de Años",
                  selectedList: this.properties.Ano,
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
                  key: "listPickerFieldId2",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
