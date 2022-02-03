/****spfx imports *****/
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  ListViewAccessor,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import { sp } from "@pnp/sp/presets/all";
import * as xlsx from 'xlsx';
import "@pnp/sp/webs";    
import "@pnp/sp/lists";    
import "@pnp/sp/items";  
import "@pnp/sp/views"; 

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomButtonCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}


const LOG_SOURCE: string = "CustomButtonCommandSet";

/**** generating random number test code ****/

var dateObj = new Date();
var month = dateObj.getUTCMonth() + 1; //months from 1-12
var day = dateObj.getUTCDate();
var year = dateObj.getUTCFullYear();

const newRandNum  = year+""+""+ month+""+ day +"_"+ Math.floor(Math.random() * 999999) + 5;
//console.log('newRandNum',newRandNum);


export default class CustomButtonCommandSet extends BaseListViewCommandSet<ICustomButtonCommandSetProperties> {
  context: any;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized CustomButtonCommandSet");
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length ===1;
    }
  }

  private async viewData (){    
    const items =sp.web.lists.getByTitle("MRF");
    const result = await items.views.getById("FEB744C4-9F87-4CE8-A6E0-7399A2E42CBC").fields();
    console.log('viewResult',result);   
   
   
  } 
 
/**** function to update Status and UploadID ****/
private async updateListItem(itemID: any) {

    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Exported", //column to be updated in the list
      UploadID: newRandNum,
    });

}
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {
  /****----------------------------------------------------------------------------------------------------- ****/    
      case "COMMAND_1":
       /* generating random number test code 
        var dateObj = new Date();
        var month = dateObj.getUTCMonth() + 1; //months from 1-12
        var day = dateObj.getUTCDate();
        var year = dateObj.getUTCFullYear();
        
        var randID2  = year+""+""+ month+""+ day +"_"+Math.floor(Math.random() * 16) + 5;
        console.log('randID2',randID2);*/

        this.viewData();
        Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
  /****----------------------------------------------------------------------------------------------------- ****/    
      case "COMMAND_2":
        
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
              const listId=` ${row.getValueByName('ID')}`;
              console.log('listId',listId);
              this.updateListItem(listId);

          });
      }
      //  Dialog.alert(`${this.properties.sampleTextTwo}`);
  /****----------------------------------------------------------------------------------------------------- ****/    
        break;
      default:
        throw new Error("Unknown command");
    }
  }
}


