/****spfx imports *****/
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { SPPermission } from "@microsoft/sp-page-context";
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

/********************************** Generating random number for UploadId **********************************/
var dateObj = new Date();
var month = dateObj.getUTCMonth() + 1; //months from 1-12
var day = dateObj.getUTCDate();
var year = dateObj.getUTCFullYear();
const newRandNum =
  year +
  "" +
  "" +
  month +
  "" +
  day +
  "_" +
  Math.floor(Math.random() * 999999) +
  5;
console.log("newRandNum", newRandNum);
//const varTemp= sp.web.lists.getByTitle("MRF").items.get()

export default class CustomButtonCommandSet extends BaseListViewCommandSet<ICustomButtonCommandSetProperties> {
  public context: any;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized CustomButtonCommandSet");
    /* 
    // code to hide button
    let newbutton: any = document.getElementsByName('New'); 
    
    newbutton.style.display = "none";  
   */
    // Dialog.alert("Its version 3");
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    console.log('Published on Dec/19/2022');
   // const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
   /* var Libraryurl = this.context.pageContext.list.title;
    console.log("Libraryurl", Libraryurl);
   

    const compareOneCommand2: Command = this.tryGetCommand("COMMAND_2");
    const compareOneCommand3: Command = this.tryGetCommand("COMMAND_3");
    const compareOneCommand4: Command = this.tryGetCommand("COMMAND_4");
    const compareOneCommand5: Command = this.tryGetCommand("COMMAND_4");

    if (
      compareOneCommand ||
      compareOneCommand2 ||
      compareOneCommand3 ||
      compareOneCommand4 ||
      compareOneCommand5
    ) {
      // This command make the button visible for the below Librayurl list only.
      compareOneCommand.visible = Libraryurl == "MRF";
      compareOneCommand2.visible = Libraryurl == "MRF";
      compareOneCommand3.visible = Libraryurl == "MRF";
      compareOneCommand4.visible = Libraryurl == "MRF";
      compareOneCommand5.visible = Libraryurl == "MRF";
    }

    */
    let isFullControl = this.checkFullControlPermission();
    console.log('isFullControl',isFullControl);
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
  
    if (isFullControl) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = isFullControl === true;
    //  this.checkFullControlPermission(compareOneCommand, SPPermission.editListItems);
      //compareOneCommand.visible;
      //alert('its visible');
    }
  }

  private checkFullControlPermission = (): boolean => {
    //Full Control group can add item to list/library and mange web.
    let permission = new SPPermission(
      this.context.pageContext.web.permissions.value
    );
    let isFullControl = permission.hasPermission(SPPermission.manageWeb);
    return isFullControl;
  };
  private deleteListItems() {
    //Get all items of list
    let varDeleteList = sp.web.lists
      .getByTitle("MRF")
      .views.getByTitle("TestView")
      .fields.removeAll();
  }

  /********************************** Function to get data from particular view of a list **********************************/
  private async viewData() {
    var url = this.context.pageContext.web.serverRelativeUrl;
    //console.log("url", url);

    const folderName = "FileUpload";
    var newURL = url + "/" + folderName;
    var varContent = "";
    const varFileName = "MileageAPFile_" + `${newRandNum}.csv`;

    const newUpload = sp.web
      .getFolderByServerRelativeUrl(newURL)
      .files.add(varFileName, File, true)
      .then(async (data) => {
        //console.log('hello',data);
        Dialog.alert("Generated and uploaded file successfully ");
        const newVar = sp.web
          .getFileByServerRelativeUrl(`${newURL}/${varFileName}`)
          .setContent(varContent);
        data.file.getItem().then((item) => {
          item.update({
            Title: "MileageTitle",
          });
        });
      });
    /****************Retrieve a list view using sharepoint framework Typescript API **********************************/
    const executeJson = (endpointUrl, payload) => {
      const opt: ISPHttpClientOptions = { method: "GET" };
      if (payload) {
        opt.method = "POST";
        opt.body = JSON.stringify(payload);
      }
      return this.context.spHttpClient.fetch(
        endpointUrl,
        SPHttpClient.configurations.v1,
        opt
      );
    };

    const getListItems = (webUrl, listTitle, queryText) => {
      var viewXml = "<View><Query>" + queryText + "</Query></View>";
      var endpointUrl =
        webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getitems";
      //console.log('hii',endpointUrl);
      var queryPayload = { query: { ViewXml: viewXml } };
      return executeJson(endpointUrl, queryPayload);
    };

    const getListViewItems = (webUrl, listTitle, viewTitle) => {
      var endpointUrl =
        webUrl +
        "/_api/web/lists/getByTitle('MRF')/Views/getbytitle('" +
        viewTitle +
        "')/ViewQuery";
      return executeJson(endpointUrl, null)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then((data) => {
          var viewQuery = data.value;
          return getListItems(webUrl, listTitle, viewQuery);
        });
    };

    /************************************** getting items and values from a view of a list**************************************/
    //const url2 = "https://pdsb1.sharepoint.com/sites/Mileage";

    getListViewItems(url, "MRF", "UploadFile")
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((response) => {
        varContent =
          varContent +
          "FISScriptV2" +
          "," +
          "UploadStatus" +
          "," +
          "EmployeeName" +
          "," +
          "ItemID" +
          "," +
          "CreatedDateOnly" +
          "," +
          "Desc1" +
          "," +
          "GroupCode" +
          "," +
          "ID" +
          "," +
          "ChargeCode" +
          "," +
          "EndDate" +
          "," +
          "StartDate" +
          "," +
          "ModifiedBy" +
          "," +
          "TotalCost" +
          "," +
          "UploadID" +
          "," +
          "Status" +
          "," +
          "\n";
        for (var item of response.value) {
          console.log("item", item);
          varContent =
            varContent +
            `"${item.FISScriptV2}"` +
            "," +
           //"Exported" +
            `"${item.UploadStatus}"` +
            "," +
            `"${item.Employee_x0020_Name}"` +
            "," +
            `"${item.ItemID}"` +
            "," +
            `"${item.CreatedDateOnly}"` +
            "," +
            `"${item.Desc1}"` +
           "," +
            `"${item.Group_x0020_Code}"` +
            "," +
            `"${item.ID}"` +
            "," +
            `"${item.ChargeCode}"` +
            "," +
            `"${item.OData_EndDate}"` +
            "," +
            `"${item.StartDate}"` +
            "," +
            `"${item.Modified}"` +
            "," +
            `"${item.Total_x0020_Cost}"` +
            "," +
           // `"${newRandNum}"` +
            `"${item.UploadID}"` +
            "," +
            `"${item.Status}"` +
            "," +
            "\n";
        }
        // console.log("Its new", varContent);
      });


  }

  /**** function to update Status and UploadID on Completed button****/
  private async updateListItem(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Completed", //column to be updated in the list
      UploadID: newRandNum,
    });
  }

  /**** function to update Status and UploadID on Pending button****/
  private async updateListItemPending(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Not Started", //column to be updated in the list
      UploadID: newRandNum,
    });
  }

  /**** function to update Status and UploadID on Pending button****/
  private async updateListItemDeferred(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Deferred", //column to be updated in the list
      UploadID: newRandNum,
    });
  }

  /**** function to update Status and UploadID on Uplaod button****/
  private async updateListItemUpload(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Exported", //column to be updated in the list
      UploadID: newRandNum,
    });
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      /********************************Generate Upload File -FIS Button---------------------------------------****************************/
      case "COMMAND_1":
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItemUpload(listId);
          });
        }
        window.setTimeout(() => {
          Dialog.alert("Status updated to Exported successfully ");
        }, 5000);
        // Dialog.alert("This is Deffered button");
        this.viewData();
        break;
        //Dialog.alert("File uploaded successfully");
      
    
      /********************************Completed Button-----------------------------------------****************************/
      case "COMMAND_2": //Completed Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItem(listId);
          });
        }
        window.setTimeout(() => {
          Dialog.alert("Status updated to completed successfully ");
        }, 5000);

        // this.viewData();
        //  Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;

      /********************************Pending Button-----------------------------------------****************************/
      case "COMMAND_3": //Pending Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItemPending(listId);
          });
        }
        window.setTimeout(() => {
          Dialog.alert("Status updated to pending successfully ");
        }, 5000);

        //this.deleteListItems();
        // Dialog.alert("This is Pending Button");
        break;

      /********************************Deffered Button-----------------------------------------****************************/
      case "COMMAND_4": //Deferred Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItemDeferred(listId);
          });
        }

        window.setTimeout(() => {
          Dialog.alert("Status updated to deferred successfully ");
        }, 5000);
        // Dialog.alert("This is Deffered button");
        break;
      /********************************Upload Button-----------------------------------------****************************/
      case "COMMAND_5": //Upload button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItemUpload(listId);
          });
        }
        Dialog.alert("File uploaded successfully");
        this.viewData();

        break;
      default:
        throw new Error("Unknown command");
    }
  }
}
