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
import { sp, SPBatch } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/folder";
import "@pnp/sp/site-users/web";
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ISiteUserProps } from "@pnp/sp/site-users/";
import "@pnp/sp/fields";
import { List } from '@pnp/sp/lists';
import { Batch } from '@pnp/odata';
import { JSONParser } from "@pnp/odata";
import { IFile } from '@pnp/sp/files/types';
import OptionDialog from './OptionsDialog';
import { IDropdownOption } from 'office-ui-fabric-react';


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
  private _selectedKey: IDropdownOption;
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

  private async updateFile(itemID: any, list: any, batch: SPBatch){
    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    const parser = new JSONParser();
    let fileType = await list.items.getById(itemID).get(parser);
    let currentValue = await list.items.getById(itemID).select('ExternalSite').get(parser);
    currentValue = currentValue.ExternalSite;
    let newValue= null;
    if(fileType.FileSystemObjectType == 1){
     // console.log('Folder: ');
     // console.log(fileType);
      let files = await list.items.getById(itemID).folder.files();
      let folders = await list.items.getById(itemID).folder.folders();
      //console.log(files);
      //console.log(folders);
      for(let i in files) {
        //console.log('File:');
        //console.log(files[i]);
        let url = files[i].ServerRelativeUrl;
        //console.log(url);
        let file = await sp.web.getFileByServerRelativeUrl(url).getItem();
        let id = file['Id'];
        //console.log('List Id:'+ id);
        batch = await this.updateFile(id, list, batch);
      }
      for(let i in folders) {
        //console.log('Folder:');
        //console.log(folders[i]);
        let url = folders[i].ServerRelativeUrl;
        // console.log(url);
        let folder = await sp.web.getFolderByServerRelativeUrl(url).getItem();
        // console.log(folder);
        let id = folder['Id'];
        console.log('List Id:'+ id);
        batch = await this.updateFile(id, list, batch);
      }
    }
    else{
      newValue =(this._selectedKey.key != 'No');
      // console.log(this._selectedKey.key != 'No');
      // console.log(newValue);
      list.items.getById(itemID).inBatch(batch).update({ ExternalSite: newValue }, "*", entityTypeFullName).then(b => {
        console.log(b);
      });
    }
    return batch;
  }

  private async updateListItems(Rows: ReadonlyArray<RowAccessor>) {
    // Update list item here
    if(this._selectedKey == undefined){
      return;
    }
    console.log(this._selectedKey);
    let list = sp.web.lists.getByTitle("Documents");
    let batch = sp.web.createBatch();
    list.fields.getByTitle('External Site').update({
      ReadOnlyField: false
    });
    for(let item of Rows){
      let itemID =  item.getValueByName('ID');
      batch = await this.updateFile(itemID, list, batch);
    }
    list.fields.getByTitle('External Site').update({
      ReadOnlyField: true
    });
    await batch.execute();
    console.log("Done");
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    // console.log('ROW:', event.selectedRows[0].getValueByName('ExternalSite'));
    switch (event.itemId) {
      case 'COMMAND_1':
        const diag:  OptionDialog = new OptionDialog();
        diag.show().then(()=>{
          this._selectedKey = diag.selectedKey;
          this.updateListItems(event.selectedRows);
        });

        break;
      default:
        throw new Error('Unknown command');
    }

  }
}
