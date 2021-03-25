import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { BaseDialog, Dialog } from '@microsoft/sp-dialog';
import * as strings from 'ExternalFieldUpdaterCommandSetStrings';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "@pnp/sp/site-users/web";
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ISiteUserProps } from "@pnp/sp/site-users/";
import "@pnp/sp/fields";
import { List } from '@pnp/sp/lists';
import { Batch } from '@pnp/odata';
<<<<<<< HEAD
import { JSONParser } from "@pnp/odata";
=======
>>>>>>> 1d8ac18df09c439d5eabb2939eccc9dd2ed33289



/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExternalFieldUpdaterCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

 interface ListItem {
  // This is an example; replace with your own properties
  ExternalSite: boolean;
  ID: string;
}

const LOG_SOURCE: string = 'ExternalFieldUpdaterCommandSet';


export default class ExternalFieldUpdaterCommandSet extends BaseListViewCommandSet<IExternalFieldUpdaterCommandSetProperties> {
  private isInOwnersGroup: boolean = false;
  @override
  public async onInit(): Promise<void> {

    await super.onInit();
    let user = await sp.web.currentUser();
    await sp.setup({ spfxContext: this.context });
    this.isInOwnersGroup = user.IsSiteAdmin;

    return Promise.resolve<void>();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.tryGetCommand('COMMAND_1').visible = this.isInOwnersGroup && (event.selectedRows.length >= 1);
  }

  private async updateFile(itemID: any, list: any){
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    const parser = new JSONParser();
    let fileType = await list.items.getById(itemID).select('FileSystemObjectType').get(parser);
    let currentValue = list.items.getById(itemID).select('ExternalSite').get(parser);
    let newValue;
    let batch = sp.web.createBatch();
    console.log(fileType);
    if(fileType == 1){
      newValue = null;
      let files = await sp.web.getFolderById(itemID).files();
      for(let file of files){
        console.log(file);
        this.updateFile(file.ListId , list);
      }
    }
    else{
      (currentValue == 'No') ? newValue = true : newValue = false;
      console.log(newValue);
      list.items.getById(itemID).inBatch(batch).update({ ExternalSite: newValue }, "*", entityTypeFullName).then(b => {
        console.log(b);
      });
    }
    await batch.execute();
  }

  private updateListItems(Rows: ReadonlyArray<RowAccessor>) {
    // Update list item here
    let list = sp.web.lists.getByTitle("Documents");
    list.fields.getByTitle('External Site').update({
      ReadOnlyField: false
    });
    for(let item of Rows){
      let itemID =  item.getValueByName('ID');
      this.updateFile(itemID, list);
    }
    list.fields.getByTitle('External Site').update({
      ReadOnlyField: true
    });
    console.log("Done");
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let newValue: boolean = false;
    // console.log('ROW:', event.selectedRows[0].getValueByName('ExternalSite'));
    switch (event.itemId) {
      case 'COMMAND_1':
        this.updateListItems(event.selectedRows);
        Dialog.alert(`External Sync Updated`); //.then(() => {location.reload()});
        break;
      default:
        throw new Error('Unknown command');
    }

  }
}
