import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'ExternalFieldUpdaterCommandSetStrings';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ISiteUserProps } from "@pnp/sp/site-users/";
import "@pnp/sp/fields";
import { List } from '@pnp/sp/lists';



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


  private async updateListItems(Rows: ReadonlyArray<RowAccessor>) {
    // Update list item here
    let list = sp.web.lists.getByTitle("Documents");
    list.fields.getByTitle('External Site').update({
      readOnlyField: false
    });

    const entityTypeFullName = await list.getListItemEntityTypeFullName();
    let batch = sp.web.createBatch();
    for(let item of Rows){
      let itemID =  item.getValueByName('ID');
      let newValue;
      (item.getValueByName('ExternalSite') == 'No') ? newValue = true : newValue = false;
      // note requirement of "*" eTag param - or use a specific eTag value as needed
      list.items.getById(itemID).inBatch(batch).update({ ExternalSite: newValue }, "*", entityTypeFullName).then(b => {
        console.log(b);
      });
    }
    await batch.execute();
    list.fields.getByTitle('External Site').update({
      readOnlyField: true
    });
    console.log("Done");
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let newValue: boolean = false;
    // console.log('ROW:', event.selectedRows[0].getValueByName('ExternalSite'));
    switch (event.itemId) {
      case 'COMMAND_1':
        this.updateListItems(event.selectedRows);
        Dialog.alert(`External Sync Updated`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
