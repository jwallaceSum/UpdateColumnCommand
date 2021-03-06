import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
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

export interface ListItem {
  // This is an example; replace with your own properties
  ExternalSite: string;
  ID: string;

}

const LOG_SOURCE: string = 'ExternalFieldUpdaterCommandSet';


export default class ExternalFieldUpdaterCommandSet extends BaseListViewCommandSet<IExternalFieldUpdaterCommandSetProperties> {
  private isInOwnersGroup: boolean = false;
  private list = sp.web.lists.getByTitle('Documents');
  private stat: boolean = true;

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
    this.stat = this.isInOwnersGroup && (event.selectedRows.length >= 1);
    this.tryGetCommand('COMMAND_1').visible = this.stat;
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let newValue: boolean = false;
    // console.log('ROW:', event.selectedRows[0].getValueByName('ExternalSite'));
    switch (event.itemId) {
      case 'COMMAND_1':
        this.list.fields.getByTitle('External Site').update({
          ReadOnlyField: false
        });
        for(let item of event.selectedRows) {
          console.log('Value', item.getValueByName('ExternalSite'));
          (item.getValueByName('ExternalSite') == 'No') ? newValue = true: newValue = false;
          console.log('New Value', newValue);
          this.list.items.getById(item.getValueByName('ID')).update({
              ExternalSite: newValue
          });
          console.log('Value:', this.list.items.getById(item.getValueByName('ID')).fieldValuesAsHTML());
        }
        this.list.fields.getByTitle('External Site').update({
          ReadOnlyField: true
        });
        console.log("Read:", (this.list.fields.getByTitle('External Site')));
        Dialog.alert(`External Sync Enabled`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}