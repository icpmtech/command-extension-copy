import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import { HttpRequestError } from "@pnp/odata";
import * as strings from 'CopyExtensionCommandSetStrings';
import { IFolder } from '@pnp/sp/folders';
import IProcessFolder from './IProcessFolder';
import { IContextInfo } from '@pnp/sp/sites';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICopyExtensionCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'CopyExtensionCommandSet';

export default class CopyExtensionCommandSet extends BaseListViewCommandSet<ICopyExtensionCommandSetProperties> {

  private readonly CONFIGURATION_LIST_NAME = "configcopyextension";
  private nestedFolders:boolean;
  private setBatchFolders:Function = new Function();
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CopyExtensionCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        // destination is a server-relative url of a new file
       this.initExecution(event);
       
        break;
      default:
        throw new Error('Unknown command');
    }

   
    }
    private initExecution(event: IListViewCommandSetExecuteEventParameters) {
     
      sp.web.select("Title").get().then(w => {
        // show our title
      console.log(w.Title);
        sp.web.lists.getByTitle(this.CONFIGURATION_LIST_NAME).items.get().then((items: any[]) => {
          console.log(items);
         let itemConfiguration= items.filter(e => e.Title==w.Title);
         console.log(itemConfiguration);
         if(itemConfiguration)
         {
           //Obter da lista de configurações o site origem e destino
         let title= itemConfiguration[0].Title;
         let origin=itemConfiguration[0].origin;
         let destiny= itemConfiguration[0].destiny;
         //Obtemos o destino e origem do ficheiro
           const origemUrl = event.selectedRows[0].getValueByName("FileRef");
           sp.site.getContextInfo().then(onfulfilled =>{

            if( onfulfilled.SiteFullUrl)
            {
              let tenantUrl=onfulfilled.SiteFullUrl.split("sites")[0];
              console.log(onfulfilled.SiteFullUrl);
              var { pathDestiny, siteUrl, pathRoot, destinationUrl } = this.builderParams(origemUrl, event, destiny,tenantUrl);
              //Criação e valição do processo de copia 
             //Passo 1- criar estrura
             this.copyProcessor(pathDestiny, siteUrl, origemUrl, pathRoot, destinationUrl);
            }
        });
         
         }
         else{
          Dialog.alert("Validar as configurações.");
         }
        
      });
    });
      
  }
  private builderParams(origemUrl: any, event: IListViewCommandSetExecuteEventParameters, destiny: any,tenantUrl:any) {
    let paths = origemUrl.split('/');
    let pathDestiny = "";
    let pathsDestinity = "";
    let folderName = paths[4];
    paths.forEach(function (path, i) {
      if (!(i == 0 || i == 1 || i == 2 || i == 3 || i == paths.length - 1))
        pathDestiny += `/${path}`;
    });
    pathDestiny = pathDestiny.slice(1);
    const filename = event.selectedRows[0].getValueByName("FileLeafRef");
    //Construimos os cainhos de destino do ficheiro  
    let destinationUrl = `/sites/${destiny}/Shared Documents/${pathDestiny}/${filename}`;
    if (pathDestiny=="") {
      destinationUrl = `/sites/${destiny}/Shared Documents/${filename}`;
    }
    const destinationUrlFolder = `sites/${destiny}`;
    const pathRoot = `${pathDestiny}`;
    //Construção do url do Site de destino
   

    debugger;
    let tenant =tenantUrl;
    let siteUrl = `${tenant}${destinationUrlFolder}`;
    return { pathDestiny, siteUrl, pathRoot, destinationUrl };
  }

  private copyProcessor(pathDestiny: string, siteUrl: string, origemUrl: any, pathRoot: string, destinationUrl: string) {
    if (pathDestiny!== "") {
      this.createStrutuctureFolders(siteUrl, origemUrl, pathRoot);
      //Passo 2- copiar ficheiro
      setTimeout(() => {
        this.copyFile(destinationUrl, origemUrl);
      }, 5000);
    }
    else
    {
      this.copyFile(destinationUrl, origemUrl);
    }
  }

  private  createStrutuctureFolders(siteUrl: string, origemUrl: any, path: string) {
   
   
     this.createDirectories(siteUrl,path);

    //web.getFileByServerRelativePath(origemUrl).copyByPath(`${destinationUrl}/${filename}`, true, true);
    //sp.web.rootFolder.folders.getByName("doc").folders.getByName("Test").copyByPath(destinationUrl, true);
    
  }
 
  private  createDirectories(siteUrl:any,root?:string,childFolder?:string,parent?:string,_pathDirectories?:any[])
  {

   
    const path = 'Shared Documents';
    let folderProecessing:IProcessFolder []=[];
      _pathDirectories=root.split("/");
      _pathDirectories.forEach((t)=> {
        let folder:IProcessFolder={
          key:t,
          value:t,
          nestedFolder:true
        };
        folderProecessing.push(folder);
      });
     let web = Web(siteUrl);
     this._addFolders(folderProecessing,path, web)
  }
  
  async  _addFolders(foldersToAdd: IProcessFolder[],root:any, web:any) {
    let currentFolderRelativeUrl =root;
    let batchAddFolders = null;
    let newFolder: IFolder;
    if (!this.nestedFolders) {
      batchAddFolders = web.createBatch();
    }

   this.setBatchFolders([] as IProcessFolder[]);

    try {
      for (let fol of foldersToAdd) {
        if (fol.nestedFolder) {
          try {
            if (currentFolderRelativeUrl) {
            
              newFolder = await web.getFolderByServerRelativePath("!@p1::" + currentFolderRelativeUrl).addSubFolderUsingPath(fol.value);
              currentFolderRelativeUrl = await newFolder.serverRelativeUrl.get();
              this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: true }]);
            }
            else {
              throw new Error("Current folder URL is empty");
            }
          } catch (nestedError) {
            if(await this.raiseException(nestedError)) {
              console.error(`Error during the creation of the folder [${fol.value}]`);
              this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: false }]);

              throw nestedError;
            }
            else {
            
              currentFolderRelativeUrl += "/" + fol.value;
              this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: true }]);
            }
          }
        }
        else {
            web.getFolderByServerRelativePath("!@p1::" + currentFolderRelativeUrl).inBatch(batchAddFolders).addSubFolderUsingPath(fol.value)
          .then(_ => {
            console.log(`Folder [${fol.value}] created`);
            this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: true }]);
          })
          .catch(async(nestedError: HttpRequestError) => {
          
            if(await this.raiseException(nestedError)) {
              console.error(`Error during the creation of the folder [${fol.value}]`);
              this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: false }]);

              throw nestedError;
            }
            else {
              currentFolderRelativeUrl += "/" + fol.value;
              this.setBatchFolders(oldBatchFolders => [...oldBatchFolders, { key: fol.key, value: fol.value, created: true }]);
            }
          });
        }
      }

      if (! this.nestedFolders) {
        await batchAddFolders.execute();
      }

    } catch (globalError) {
      console.log('Global error');
      console.log(globalError);
    }
  }

  async  raiseException(nestedError: HttpRequestError): Promise<boolean> {
    let raiseError: boolean = true;
    return new Promise<boolean>(async(resolve, reject) => {
      if (nestedError.isHttpRequestError) {
        try {
          const errorJson = await (nestedError).response.json();
          console.error(typeof errorJson["odata.error"] === "object" ? errorJson["odata.error"].message.value : nestedError.message);

          if (nestedError.status === 500) {
            // Don't raise an error if the folder already exists
            if (nestedError.message.indexOf('exist') > 0) {
              raiseError = false;
            }

            console.error(nestedError.statusText);
          }
        } catch (error) {
          console.error(error);
        }

      } else {
        console.log(nestedError.message);
      }

      resolve(raiseError);
    });
  }

private copyFile(destinationUrl: string, origemUrl: any) {
  const message = `Copiado para a seguinte directoria:${destinationUrl}`;

  sp.web.getFileByServerRelativePath(origemUrl).copyByPath(destinationUrl, true, false).then(function (res) {
    console.log(res);
    Dialog.alert(message);
  }).catch(onrejected=>{

    Dialog.alert(`Erro ao copiar validar com administrador:${onrejected.reason}`);
  });
  return message;
}

}

