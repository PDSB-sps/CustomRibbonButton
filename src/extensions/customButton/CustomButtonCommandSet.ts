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
const varTemp= sp.web.lists.getByTitle("MRF").items.get()

export default class CustomButtonCommandSet extends BaseListViewCommandSet<ICustomButtonCommandSetProperties> {
  context: any;
  domElement: any;

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
      compareOneCommand.visible = event.selectedRows.length === 0;
    }
  }

  /********************************** Function to get data from particular view of a list **********************************/
  private async viewData() {
    var url = this.context.pageContext.web.serverRelativeUrl;
    //console.log("url", url);

    const folderName = "FileUpload";
    var newURL = url + "/" + folderName;
    var varContent = "";
    const varFileName = "MileageAPFile_" + `${newRandNum}.csv`;
    //console.log('navpreet',varFileName);

    const newUpload = sp.web
      .getFolderByServerRelativeUrl(newURL)
      .files.add(varFileName, File, true)
      .then(async (data) => {
        //console.log('hello',data);
        Dialog.alert("File uploaded sucessfully");
        const newVar = sp.web
          .getFileByServerRelativeUrl(`${newURL}/${varFileName}`)
          .setContent(varContent);
        data.file.getItem().then((item) => {
          item.update({
            Title: "MileageTitle",
          });
        });
      });
    /****************Retrieve a list view using sharepoint fremework Typecript API **********************************/
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
          //console.log("item", item);
          varContent =
            varContent +
            `"${item.FISScriptV2}"` +
            "," +
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
            `"${item.UploadID}"` +
            "," +
            `"${item.Status}"` +
            "," + 
            "\n";
        }
       // console.log("Its new", varContent);
      });
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
      /********************************Upload Button---------------------------------------****************************/
      case "COMMAND_1":
        this.viewData();
        // Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      /********************************Completed Button-----------------------------------------****************************/
      case "COMMAND_2":
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItem(listId);;
          });
        }

        window.setTimeout( () =>{
          Dialog.alert("Record updated successfully "); 
       }, 5000);
 
        // this.viewData();
        //  Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;

        /********************************Pending Button-----------------------------------------****************************/
        case "COMMAND_3":
        Dialog.alert("This is Pending Button");
          break;

        /********************************Deffered Button-----------------------------------------****************************/
          case "COMMAND_4":
          Dialog.alert("This is Deffered button");
            break;

      default:
        throw new Error("Unknown command");
    }
  }
}
