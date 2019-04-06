import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { sp } from "@pnp/sp";
import * as strings from 'UnpromoteNewsCommandSetStrings';
import UnPromoteNewsComponent from './Components/UnpromoteNews';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IUnpromoteNewsCommandSetProperties {

}

const LOG_SOURCE: string = 'UnpromoteNewsCommandSet';

export default class UnpromoteNewsCommandSet extends BaseListViewCommandSet<IUnpromoteNewsCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized UnpromoteNewsCommandSet');
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
  private async getPageDetails(pageRelativeUrl: string): Promise<void> {
    try {
      const pageItem = await sp.web.getFileByServerRelativeUrl(pageRelativeUrl).listItemAllFields.select('PromotedState').get();
      return (pageItem["PromotedState"]);
    }
    catch (ex) {
      console.log(ex);
    }

  }
  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    try {
      switch (event.itemId) {
        case 'COMMAND_1':
          let relativeUrl = event.selectedRows[0].getValueByName('FileRef');
          let pageNameToolTip = event.selectedRows[0].getValueByName('FileLeafRef');
          //Gets short name of page
          let pageName = pageNameToolTip.length > 20 ? pageNameToolTip.substring(0, 8) + "..." +
            pageNameToolTip.substring(pageNameToolTip.length - 8, pageNameToolTip.length) :
            pageNameToolTip;
          const callout: UnPromoteNewsComponent = new UnPromoteNewsComponent();
          callout.pageName = pageName;
          callout.pageRelativeUrl = relativeUrl;
          callout.pageNameToolTip = pageNameToolTip;
          //Gets the page promoted state
          const promotedState = await this.getPageDetails(relativeUrl);
          callout.promotedState = Number(promotedState);
          callout.show();
          break;
        default:
          throw new Error('Unknown command');
      }
    }
    catch (ex) {
      console.log(ex);
    }

  }
}
