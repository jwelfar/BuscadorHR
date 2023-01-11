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
  cedula: string;
  name: string;
  company: string;
  docTyp: string;
  doc: string;
  date: string;
  filePicked: IFilePickerResult;
  fileUpl: File[];
  DatosHR: IDetailsListItem[];
  lab: File[];
  fileURL: string;
}

export default class FormularioHr extends React.Component<
  IFormularioHrProps,
  IFormularioHrState
> {
  constructor(props: IFormularioHrProps) {
    super(props);

    this.state = {
      cedula: "",
      name: "",
      company: "",
      docTyp: "",
      doc: "",
      date: "",
      fileUpl: null,
      filePicked: null,
      DatosHR: [],
      lab: null,
      fileURL: "",
    };
  }

  sendForm = async (): Promise<void> => {
    if (
      !this.state.cedula ||
      !this.state.name ||
      !this.state.company ||
      !this.state.docTyp ||
      !this.state.doc ||
      !this.state.date
    ) {
      alert("Uno o más campos requeridos están vacíos.");
    } else {
      await this.onSave(this.state.fileUpl);
    }
  };

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
    await getSP(this.props.context).web.folders.addUsingPath("lab/12324");

    await getSP(this.props.context)
      .web.rootFolder.folders.getByUrl("lab")
      .addSubFolderUsingPath("12324/subfolder");

    const finalSubFolderResult = await getSP(this.props.context)
      .web.rootFolder.folders.getByUrl("lab")
      .addSubFolderUsingPath(`12324/subfolder/${this.state.doc}`);
    return finalSubFolderResult;
  };

  onSave = async (_files: File[]): Promise<void> => {
    try {
      if (_files !== null) {
        for (let index = 0; index < _files.length; index++) {
          const _file = _files[index];
          const basicURL = `/sites/Desarrollo/pruebaJW/lab`;
          const addFolderURL = `${basicURL}/12324/subfolder/${this.state.doc}/${_file.name}`;
          const routeFolder = `12324/subfolder/${this.state.doc}/${_file.name}`;
          //Guarda en Biblioteca
          let folder: any;
          if (_file.size <= 10485760) {
            // small upload
            folder = await getSP(this.props.context)
              .web.getFolderByServerRelativePath(
                this.props.lab["title"] //biblioteca labor history
              )
              .files.addUsingPath(addFolderURL, _file, { Overwrite: true });
          } else {
            // large upload
            folder = await getSP(this.props.context)
              .web.getFolderByServerRelativePath(this.props.lab["title"])
              .files.addChunked(
                addFolderURL,
                _file,
                (data) => {
                  console.log("progress", data);
                },
                true
              );
          }
          await folder.file.getItem();

          this.setState({ fileURL: routeFolder });
        }

        const rowData = {
          field_4: this.state.cedula,
          field_5: this.state.name,
          field_6: this.state.company,
          field_7: this.state.docTyp,
          field_8: this.state.doc,
          field_9: this.state.date,
          field_11: this.state.fileUpl[0].name.substring(
            0,
            this.state.fileUpl[0].name.indexOf(".")
          ),
          field_13: this.state.fileUpl[0].size,
          field_14: this.state.fileURL,
        };
        const data = await getSP(this.props.context)
          .web.lists.getByTitle("DatosHR")
          .items.add(rowData);
        // .then(() => {
        //   alert("Se ha registrado correctamente");
        // });
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
          date: "",
          DatosHR: [],
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

  public render(): React.ReactElement<IFormularioHrProps> {
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
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

          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Cédula"
              id="cedula"
              value={this.state.cedula}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
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
              label="Nombre"
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
            <TextField
              label="Empresa"
              id="company"
              value={this.state.company}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  company: newValue,
                });
              }}
            />
          </div>

          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Tipo de Documento"
              id="docTyp"
              value={this.state.docTyp}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  docTyp: newValue,
                });
              }}
            />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
            <TextField
              label="Documento"
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
            <TextField
              label="Fecha"
              id="date"
              value={this.state.date}
              onChange={(
                event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newValue?: string
              ) => {
                this.setState({
                  date: newValue,
                });
              }}
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
                await this.sendForm();
              }}
              allowDisabledFocus
            />
          </div>
        </div>
      </div>
    );
  }
}
