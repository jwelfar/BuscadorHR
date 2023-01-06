import * as React from "react";
import styles from "./BuscaroHr.module.scss";
import { IBuscaroHrProps } from "./IBuscaroHrProps";
import { DetailsHeader, IColumn, IDetailsHeaderProps, Link } from "office-ui-fabric-react";
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  TextField,
} from "office-ui-fabric-react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";

let _sp: SPFI = null;

export interface IDetailsListItem {
  Id: number;
  field_4: string;
  field_5: string;
  field_7: string;
  field_8: string;
  field_9: string;
  field_14: string;
}

export interface IBuscaroHrState {
  columns: IColumn[];
  identification: string;
  docType: IDropdownOption[];
  selectDocType: IDropdownOption;
  Ano: IDropdownOption[];
  selectAno: IDropdownOption;
  DatosHR: IDetailsListItem[];
}

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

export default class BuscaroHr extends React.Component<
  IBuscaroHrProps,
  IBuscaroHrState
> {
  constructor(props: IBuscaroHrProps) {
    
    super(props);
    this.renderDetailsHeader = this.renderDetailsHeader.bind(this);
    const newURL = this.props.context.pageContext.web.absoluteUrl;
    const columnas: IColumn[] = [
      {
        key: 'column1',
        name: 'Cédula',
        ariaLabel: 'Column operations for File type, Press to sort on File type',
        fieldName: 'field_4',
         isResizable: true,
        minWidth: 150,
        maxWidth: 200,
        onRender: (global: any) => {   return <span>{global.field_4}</span>;
      }
    },
    {
      key: 'column2',
      name: 'Nombre',
      ariaLabel: 'Column operations for File type, Press to sort on File type',
      fieldName: 'field_5',
         isResizable: true,
      minWidth: 150,
      maxWidth: 200,
      onRender: (global: any) => {         return <span>{global.field_5}</span>;
    }
  },
    {
      key: 'column3',
      name: 'Tipo Documento',
      ariaLabel: 'Column operations for File type, Press to sort on File type',
      fieldName: 'field_8',
      isResizable: true,
      minWidth: 150,
      maxWidth: 200,
      onRender: (global: any) => {         return <span>{global.field_8}</span>;
    }
  },
    {
      key: 'column4',
      name: 'Documento',
      ariaLabel: 'Column operations for File type, Press to sort on File type',
      fieldName: 'field_7',
      isResizable: true,
      minWidth: 150,
      maxWidth: 200,
      onRender: (global: any) => {         return <span>{global.field_7}</span>;
    }
  },
    {
      key: 'column5',
      name: 'Ruta',
      ariaLabel: 'Column operations for File type, Press to sort on File type',
      fieldName: 'field_14',
         isResizable: true,
      minWidth: 150,
      maxWidth: 200,
      onRender: (global: any) => {         
        return <Link href={newURL+"/"+this.props.LaborH.title+"/"+global.field_14} target="_blank">{global?.field_14}</Link>;;
    },
  }
      
    ];
   
    this.state = {
      columns:columnas,
      identification: "",
      docType: [],
      selectDocType: { key: "", text: "" },
      Ano: [],
      selectAno: { key: "", text: "" },
      DatosHR: [],
    };
  }
  
  private renderDetailsHeader(detailsHeaderProps: IDetailsHeaderProps) {
    return (
      <DetailsHeader
        {...detailsHeaderProps}

      />
    );
  }
  private async getListContent(): Promise<void> {
    let items: [];
    if (
      typeof this.props.DatosHR !== "undefined" &&
      this.props.DatosHR.id?.length > 0
    ) {
      try {
        // let filtro: string;
        // if (
        //   this.state.identification &&
        //   this.state.selectDocType.key &&
        //   this.state.selectAno.key
        // ) {
        //   filtro = `field_4 eq '${this.state.identification}' and field_8 eq '${this.state.selectDocType.key}' and field_9 eq '${this.state.selectAno.key}'`;
        // } else if (
        //   this.state.identification &&
        //   this.state.selectDocType.key &&
        //   !this.state.selectAno.key
        // ) {
        //   filtro = `field_4 eq '${this.state.identification}' and field_8 eq '${this.state.selectDocType.key}'`;
        // } else if (
        //   this.state.identification &&
        //   !this.state.selectDocType.key &&
        //   this.state.selectAno.key
        // ) {
        //   filtro = `field_4 eq '${this.state.identification}' and field_9 eq '${this.state.selectAno.key}'`;
        // } else if (
        //   this.state.identification &&
        //  !this.state.selectDocType.key &&
        //   !this.state.selectAno.key
        // ) {
        //   filtro = `field_4 eq '${this.state.identification}'`;
        // }
        items = await getSP(this.props.context)
          .web.lists.getById(this.props.DatosHR.id)
          .items.select(
            "Id",
            "field_4",
            "field_5",
            "field_7",
            "field_8",
            "field_9",
            "field_14"
          )();
        // .filter(filtro)();

        this.setState({
          DatosHR: items,
        });
      } catch (err) {
        console.log("Error", err);
        err.res.json().then(() => {
          console.log("Failed to get list items!", err);
        });
      }
    }
  }

  private async getDocType(): Promise<void> {
    try {
      const data = this.state.DatosHR.map((a) => {
        return {
          key: a.field_8,
          text: a.field_8,
        } as IDropdownOption;
      });
      const docOptions = new Set(data.map((a) => JSON.stringify(a)));
      const array: IDropdownOption[] = [];
      docOptions.forEach((ele) => {
        array.push(JSON.parse(ele));
      });
      this.setState({
        docType: array,
      });
    } catch (e) {
      e.res.json().then(() => {
        console.log("Failed to get list items!", e);
      });
    }
  }

  private async getYear(): Promise<void> {
    try {
      const data = this.state.DatosHR.map((a) => {
        return {
          key: a.field_9,
          text: a.field_9,
        } as IDropdownOption;
      });

      const yearOptions = new Set(data.map((a) => JSON.stringify(a)));
      const array: IDropdownOption[] = [];
      yearOptions.forEach((ele) => {
        array.push(JSON.parse(ele));
      });
      this.setState({
        Ano: array,
      });
    } catch (e) {
      e.res.json().then(() => {
        console.log("Failed to get list items!", e);
      });
    }
  }

  onSearch = async (): Promise<void> => {
    if (!this.state.identification) {
      alert("Debes llenar el campo requerido!");
    }
    await this.getListContent();
    await this.getDocType();
    await this.getYear();
  };

  onReset = async (): Promise<void> => {
    this.setState({
      identification: "",
      docType: [],
      selectDocType: { key: "", text: "" },
      Ano: [],
      selectAno: { key: "", text: "" },
      DatosHR: [],
    });
  };
  
  public render(): React.ReactElement<IBuscaroHrProps> {
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 150 },
    };

   
   

    return (
      <><section>
        <div>
          <div className={styles.fristRow__container}>
            <TextField
              label="Identificación"
              id="identification"
              value={this.state.identification}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  identification: newValue,
                });
              } } />

            <Dropdown
              placeholder="Select an option"
              label="Tipo de documento"
              options={this.state.docType}
              styles={dropdownStyles}
              onChange={(
                event: React.FormEvent<HTMLDivElement>,
                item: IDropdownOption
              ) => {
                this.setState({
                  selectDocType: item,
                });
              } } />

            <Dropdown
              placeholder="Select an option"
              label="Año"
              options={this.state.Ano}
              styles={dropdownStyles}
              onChange={(
                event: React.FormEvent<HTMLDivElement>,
                item: IDropdownOption
              ) => {
                this.setState({
                  selectAno: item,
                });
              } } />

            <DefaultButton
              text="Buscar"
              onClick={() => this.onSearch()}
              allowDisabledFocus />

            <DefaultButton
              text="Limpiar búsqueda"
              onClick={() => this.onReset()}
              allowDisabledFocus />
          </div>

          <br />
          <DetailsList
          layoutMode={DetailsListLayoutMode.justified}
          items={ this.state.DatosHR}
          columns={this.state.columns}
          compact={true}
          setKey="set"

           />
        </div>

      </section>
      </>
    );
  }
}
