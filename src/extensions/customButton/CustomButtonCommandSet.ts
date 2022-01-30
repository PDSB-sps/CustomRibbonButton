/****spfx imports *****/
import { override } from "@microsoft/decorators";
import { Log, RandomNumberGenerator } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import { sp } from "@pnp/sp/presets/all";
import * as moment from 'moment';

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

export interface IReactPartialStateUpdateState {  
  currentDate: Date;  
  randomNumber: number;  
  ramdomText: string;  
}

const LOG_SOURCE: string = "CustomButtonCommandSet";

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
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

private getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    console.log('min',min);
    console.log('min',max);
    const randNum=Math.floor(Math.random() * (max - min + 1)) + min;
    console.log('randNum',randNum);
    return randNum;
}

/**** function to update Status and UploadID ****/
private async updateListItem(itemID: any) {
// Update list item here  
var dateObj = new Date();
var month = dateObj.getUTCMonth() + 1; //months from 1-12
var day = dateObj.getUTCDate();
var year = dateObj.getUTCFullYear();

const randID  = year+""+""+ month+""+ day +"_"+ itemID;
console.log('randID2',randID);
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Exported", //column to be updated in the list
      UploadID:  year+""+""+ month+""+ day +"_"+ itemID
    });

}
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
  /****----------------------------------------------------------------------------------------------------- ****/    
      case "COMMAND_1":
       /*  console.log("itemId", event.itemId);
        //Let filenm =""
        const listId = String(event.selectedRows[0].getValueByName("ID"));
        console.log("Its line 94");
        console.log("listId",listId);
       // foreach 
        this.updateListItem(listId); */
        var dateObj = new Date();
        var month = dateObj.getUTCMonth() + 1; //months from 1-12
        var day = dateObj.getUTCDate();
        var year = dateObj.getUTCFullYear();
        
        var randID2  = year+""+""+ month+""+ day +"_"+Math.floor(Math.random() * 16) + 5;

        console.log('randID2',randID2);

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
