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
import "@pnp/sp/webs";
import "@pnp/sp/files";
import * as strings from 'CopyExtensionCommandSetStrings';

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
        const origemUrl = event.selectedRows[0].getValueByName("FileRef");
        const filename = event.selectedRows[0].getValueByName("FileLeafRef");
        const destinationUrl =  `/sites/UnilabsDemo/docpublic/`+ filename;
        const message =  `Copiado para a directoria "docpublic"` + `o ficheiro: `+ filename;
         sp.web.getFileByServerRelativePath(origemUrl).copyTo(destinationUrl, true);
         Dialog.alert(message);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}