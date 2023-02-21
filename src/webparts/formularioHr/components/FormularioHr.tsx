import * as React from "react";
// import styles from "./FormularioHr.module.scss";
import { IFormularioHrProps } from "./IFormularioHrProps";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/folders";
import "@pnp/sp/files";

import { DefaultButton } from "office-ui-fabric-react/lib/components/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
  FilePicker,
  IFilePickerResult,
} from "@pnp/spfx-controls-react/lib/FilePicker";
import { ComboBox, DatePicker, IComboBox, IComboBoxOption, IDatePicker, IDropdownOption } from "office-ui-fabric-react";
import Swal from "sweetalert2";
// import Swal from "sweetalert2";

let _sp: SPFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

export interface IDetailsListItem {
  Id: number;
  field_4: string; //cedula
  field_5: string; //nombre
  field_6: string; //empresa
  field_7: string; //Documento
  field_8: string; //tipo documento
  field_9: string; //fecha
  field_11: string; //nombre archivo
  field_13: number; //peso archivo
  field_14: string; //ruta archivo
}

export interface IFormularioHrState {

  datepickerref:IDatePicker;
  cedula: string;
  name: string;
  company: string;
  docTyp: string;
  emptype:string;
  doc: string;
  datep: Date;
  datepicker:any;
  filePicked: IFilePickerResult;
  fileUpl: File[];
  DatosHR: IDetailsListItem[];
  lab: File[];
  fileURL: string;
  selectDocType: IDropdownOption;
  tipodoc:   IComboBoxOption[];
  empresa:   IComboBoxOption[];
  selectedServicio: any,
  selectedempresa: any,
}
  

export default class FormularioHr extends React.Component<
  IFormularioHrProps,
  IFormularioHrState
> {
  constructor(props: IFormularioHrProps) {
    super(props);
    const today = new Date();
    this.state = {
      datepickerref:null,
      cedula: "",
      name: "",
      company: "",
      docTyp: "",
      doc: "",
      datep: new Date(),
      fileUpl: null,
      filePicked: null,
      DatosHR: [],
      lab: null,
      fileURL: "",
      selectDocType: { key: "", text: "" },
      tipodoc: [],
      selectedServicio: "",
      empresa:[],
      selectedempresa:"",
      emptype:"",
      datepicker:today.getMonth()+1 + "-"+ today.getFullYear(),
    };
       this.getDocType();
      this.getenterprise();
  }

  sendForm = async (): Promise<void> => {
    if (
      !this.state.cedula ||
      !this.state.name ||
      !this.state.emptype ||
      !this.state.docTyp ||
      !this.state.doc ||
      !this.state.datep
    ) {
      alert("Uno o más campos requeridos están vacíos.");
    } else {
      await this.onSave(this.state.fileUpl);
    }
  };


  private async getDocType(): Promise<void> {
    try {
      let items: any[]=[];
      console.log("vea");
      items = await getSP(this.props.context).web.lists
      .getById(this.props.doctype.id).items.select("Id", "Title","Orden").orderBy("Orden")()  
      items = items.map(a => {
        return { key: a.Id, text: a.Title, data: a };
      });
      this.setState({ tipodoc: items });
    } catch (e) {
        console.log("Failed to get list items!", e);
      }
    }

    private async getenterprise(): Promise<void> {
    try {
      let items: any[];
      items = await getSP(this.props.context).web.lists
      .getById(this.props.empresa.id).items.select("Id", "Title")()  
      items = items.map(a => {
        return { key: a.Id, text: a.Title, data: a };
      });
      this.setState({ empresa: items });
    } catch (e) {
        console.log("Failed to get list items!", e);
      }
    }
  

  public handleChangeFile = async (filePickerResult: IFilePickerResult[]) => {
    if (filePickerResult && filePickerResult.length > 0) {
      const results: File[] = [];
      for (let index = 0; index < filePickerResult.length; index++) {
        const item = filePickerResult[index];
        const fileResultContent = await item.downloadFileContent();
        // console.log("fileResultContent", fileResultContent);
        results.push(fileResultContent);
      }
      this.setState({ fileUpl: results });
      console.log("upl", this.state.fileUpl);
    } else {
      this.setState({ fileUpl: null });
    }
  };

  // private ensureFolder = async (): Promise<any> => {
  //   // console.log(await this.createPrincipalFolder());
  //   const folder = await getSP(this.props.context)
  //     .web.getFolderByServerRelativePath(await this.createPrincipalFolder())
  //     .select("Exists")();
  //   if (!folder.Exists) {
  //     await getSP(this.props.context).web.folders.addUsingPath(
  //       await this.createPrincipalFolder()
  //     );
  //     console.log("FOLDER", folder);
  //   }
  // }; 
 
  public createPrincipalFolder = async (): Promise<any> => {
    try {
      await (getSP(this.props.context).web.lists.getById(this.props.lab.id).rootFolder.folders).addUsingPath(this.state.cedula);
      console.log(`Folder "${this.state.cedula}" created successfully in lab library`);
    } catch (error) {
      console.error(`Error creating folder "${this.state.cedula}" in  lab library`, error);
    }
  
    try {
      // add a folder to site assets
 await getSP(this.props.context).web.rootFolder.folders.getByUrl(this.props.lab.url).addSubFolderUsingPath(this.state.cedula +"/"+this.state.docTyp);
   //   await (getSP(this.props.context).web.getFolderByServerRelativePath('/lab/'+`${this.state.cedula}`) as any)
        //  .folders.addUsingPath(this.state.docTyp);
        console.log(`Subfolder "${this.state.docTyp}" created successfully`)
      
      } catch (error) {
        console.error(`Error creating subfolder "${this.state.docTyp}"`, error);
       
      }
      await this.sendForm();
     
  };

  

  onSave = async (_files: File[]): Promise<void> => {
    try {
      
      if (_files !== null) {
        for (let index = 0; index < _files.length; index++) {
          const _file = _files[index];
          const splitda=  _file.name.split('.');
          const namefile = splitda[0]+"-"+this.state.datepicker+"."+splitda[1];
          const basicURL = this.props.lab.url;
          const addFolderURL = `${basicURL}/${this.state.cedula}/${this.state.docTyp}/${namefile}`;
          const routeFolder = `${this.state.cedula}/${this.state.docTyp}/${namefile}`;
          //Guarda en Biblioteca
          let folder: any;
          if (_file.size <= 10485760) {
            
            // small upload
            folder = await getSP(this.props.context)
              .web.getFolderByServerRelativePath(
                this.props.lab.title //biblioteca labor history
              )
              .files.addUsingPath(addFolderURL, _file, { Overwrite: false });
          } else {
            // large upload
            folder = await getSP(this.props.context)
              .web.getFolderByServerRelativePath(this.props.lab.title)
              .files.addChunked(
                addFolderURL,
                _file,
                (data) => {
                  console.log("progress", data);
                },
                false
              );
          }
          await folder.file.getItem();

          this.setState({ fileURL: routeFolder });
        }

        const rowData = {
          field_4: this.state.cedula,
          field_5: this.state.name,
          field_6: this.state.emptype,
          field_7: this.state.doc,
          field_8: this.state.docTyp,
          field_9: this.state.datepicker,
          field_11: this.state.fileUpl[0].name.substring(
            0,
            this.state.fileUpl[0].name.indexOf(".")
          ),
          field_13: this.state.fileUpl[0].size,
          field_14: this.state.fileURL,
        };
        const data = await getSP(this.props.context)
          .web.lists.getByTitle("DatosHR")
          .items.add(rowData)
         .then(() => {
           Swal.fire("Se ha registrado correctamente", "success");
        });
        console.log("data Item", data);
        // .then(() => {
        //   Swal.fire("Se ha registrado correctamente", "success");
        // })
        // .catch((err) => {
        //   console.log("Error", err);
        // });
        
        this.setState({
          cedula: "",
          name: "",
          company: "",
          docTyp: "",
          doc: "",
          datep: new Date(),
          datepicker: null,
          DatosHR: [],
          selectedempresa:[],
          selectedServicio:[],
          fileUpl: null,
          lab: null,
        });
      }
    } catch (err) {
      err.res.json().then(() => {
        console.log("Failed!", err);
      });
    }
  };
  
  setMaxCosto(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void {
      this.setState({
        docTyp: option.text,
        selectedServicio: option.key ? parseInt(option.key.toString()) : null,
      })
    }

    setempresa(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void {
      this.setState({
        emptype: option.text,
        selectedempresa: option.key ? parseInt(option.key.toString()) : null,
      })
    }

     onDayPicked(date: Date | undefined | undefined) {
      this.setState({
        datep: date,
        datepicker: date.getMonth()+1 + "-"+ date.getFullYear()
      })
    }
   

  public render(): React.ReactElement<IFormularioHrProps> {
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <FilePicker
              label="Seleccione el archivo"
              buttonLabel="Agregar archivo"
              onSave={this.handleChangeFile}
              onChange={this.handleChangeFile}
              context={this.props.context as any}
              hideStockImages={true}
              hideOrganisationalAssetTab={true}
              hideWebSearchTab={true}
              hideOneDriveTab={true}
              hideSiteFilesTab={true}
            />
            <span>
              <b>
                {this.state.fileUpl === null
                  ? "Nombre del Archivo"
                  : this.state.fileUpl[0].name}
              </b>
            </span>
          </div>
          </div>
          <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Identificación "
              id="cedula"
              value={this.state.cedula}
              onChange={(
                _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  cedula: newValue,
                });
              }}
            />
          </div>

          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Nombre Persona"
              id="name"
              value={this.state.name}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  name: newValue,
                });
              }}
            />
          </div>
        </div>

        <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
          <ComboBox
              selectedKey={this.state.selectedempresa}
              placeholder="Selecciona una opción"
              label="Empresa"
              options={this.state.empresa}
              onChange={this.setempresa.bind(this)}
              required
            />

            
          </div>

          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
          <ComboBox
              selectedKey={this.state.selectedServicio}
              placeholder="Selecciona una opción"
              label="Tipo Documento"
              options={this.state.tipodoc}
              onChange={this.setMaxCosto.bind(this)}
              required
            />

          </div>
        </div>

        <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Documento Nombre"
              id="doc"
              value={this.state.doc}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  doc: newValue,
                });
              }}
            />
          </div>

          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
          <DatePicker 
           
            label='Fecha registro' 
            isRequired={ false } 
            allowTextInput={ true } 
            ariaLabel="Seleccione fecha de carga del documento"
            value={ this.state.datep } 
            onSelectDate={this.onDayPicked} 
          
        /> 
          </div>
        </div>

        <div className="ms-Grid-row">
          <div
            className="ms-Grid-col ms-sm12 ms-md12 ms-lg12"
            style={{ textAlign: "center" }}
          >
            <br />

            <DefaultButton
              text="Crear"
              onClick={async () => {
                await this.createPrincipalFolder();
               
              }}
              allowDisabledFocus
            />
          </div>
        </div>
      </div>
    );
  }
}
