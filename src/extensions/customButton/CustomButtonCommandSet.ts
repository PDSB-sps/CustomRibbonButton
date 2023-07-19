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
import { PermissionKind, sp, Web } from "@pnp/sp/presets/all";
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

//console.log("Published on",dateObj.toDateString());
//var setNewDate = dateObj.toDateString();
//const varTemp= sp.web.lists.getByTitle("MRF").items.get()

export default class CustomButtonCommandSet extends BaseListViewCommandSet<ICustomButtonCommandSetProperties> {
  public context: any;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized CustomButtonCommandSet");
    sp.setup({
      spfxContext: this.context,
      sp: { baseUrl: this.context.pageContext.web.absoluteUrl },
    });
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
    console.log("Published on June/07/2023");
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
    // console.log("isFullControl", isFullControl);
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
    //generate random number
    var dateObj = new Date();
    var month = dateObj.getUTCMonth() + 1; //months from 1-12
    var day = dateObj.getUTCDate();
    var year = dateObj.getUTCFullYear();
    const newRandNum =
      year + "" + "" + month + "" + day + Math.floor(Math.random() * 99999) + 5;
    console.log("newRandNum", newRandNum);

    var url = this.context.pageContext.web.serverRelativeUrl;
    console.log("url", url);
    const folderName = "FileUpload";
    var newURL = url + "/" + folderName;
    var varContent = "";
    const varFileName = "MileageAPFile_" + `${newRandNum}.csv`;

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
        webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getitems?$top=200";
      //console.log("hii", endpointUrl);
      var queryPayload = { query: { ViewXml: viewXml } };
      return executeJson(endpointUrl, queryPayload);
    };

    const getListViewItems = (webUrl, listTitle, viewTitle) => {
      var endpointUrl =
        webUrl +
        "/_api/web/lists/getByTitle('MRF')/Views/getbytitle('" +
        viewTitle +
        "')/ViewQuery";
     // console.log("hii222222", endpointUrl);
      return executeJson(endpointUrl, null)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then((data) => {
          console.log("Data", data);
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
        console.log("response", response);
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
          "ChargeCode" +
          "," +
          "StartDate" +
          "," +
          "End_x0020_Date" +
          "," +
          "Approver" +
          "," +
          "TotalCost" +
          "," +
          "\n";

        for (var item of response.value) {
          console.log("item", response.value.length);

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
            `"${item.ChargeCode}"` +
            "," +
            `"${item.StartDate}"` +
            "," +
            `"${item.End_x0020_Date}"` +
            "," +
            `"${item.Approver}"` +
            "," +
            `"${item.Total_x0020_Cost}"` +
            "," +
            "\n";
           

        //  const itemIDs = item.ID;
         // console.log("itemIDs", itemIDs);
         // const list = sp.web.lists.getByTitle("MRF");
         // const i = list.items.top(200).getById(itemIDs).update({
         //   Status: "Exported",
         //   UploadID: newRandNum,
         // });
        }

        for (let j = 0; j <response.value.length; j++) {
         console.log("#############",response.value[j].ID);
          const itemIDs = response.value[j].ID //.ID;
          console.log("itemIDs", itemIDs);
          const list = sp.web.lists.getByTitle("MRF");
          const i = list.items.top(200).getById(itemIDs).update({
            Status: "Exported",
            UploadID: newRandNum,
          });
        }
        console.log("Its new", varContent);
        const newUpload = sp.web
          .getFolderByServerRelativeUrl(newURL)
          .files.add(varFileName, File, true)
          .then(async (data) => {
          //  console.log("hello0000", data);
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
      });
    ///////

  }

  /************************************************************************************************************************/
/********************************** Test Function to get data from particular view of a list ***********************************************************/
    private async viewDataTest() {
      //generate random number
      var dateObj = new Date();
      var month = dateObj.getUTCMonth() + 1; //months from 1-12
      var day = dateObj.getUTCDate();
      var year = dateObj.getUTCFullYear();
      const newRandNum =
        year + "" + "" + month + "" + day + Math.floor(Math.random() * 99999) + 5;
     // console.log("newRandNum", newRandNum);
     // console.log('Test here',sp)
      var url = "/sites/Mileage/Backups" //this.context.pageContext.web.serverRelativeUrl;
     // console.log("url", url);
     // console.log("Testurl", this.context.pageContext.web.serverRelativeUrl);
      const folderName = "FileUpload";
      var newURL = url + "/" + folderName;
      //console.log("********newURL********",newURL)
      var varContent = "";
      const varFileName = "MileageAPFile_" + `${newRandNum}.csv`;
  
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
          webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getitems?$top=200";
       // console.log("Test endpointUrl", endpointUrl);
        var queryPayload = { query: { ViewXml: viewXml } };
        return executeJson(endpointUrl, queryPayload);
      };
  
      const getListViewItems = (webUrl, listTitle, viewTitle) => {
        var endpointUrl =
          webUrl +
          "/_api/web/lists/getByTitle('MRF')/Views/getbytitle('" +
          viewTitle +
          "')/ViewQuery";
       // console.log("Test endpointUrl1", endpointUrl);
        return executeJson(endpointUrl, null)
          .then((response: SPHttpClientResponse) => {
            return response.json();
          })
          .then((data) => {
           // console.log("Test Data", data);
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
        .then(async (response) => {
         // console.log("response", response);
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
            "ChargeCode" +
            "," +
            "StartDate" +
            "," +
            "End_x0020_Date" +
            "," +
            "Approver" +
            "," +
            "TotalCost" +
            "," +
            "\n";
  
          for (var item of response.value) {
           // console.log("Test item length", response.value.length);
  
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
              `"${item.ChargeCode}"` +
              "," +
              `"${item.StartDate}"` +
              "," +
              `"${item.End_x0020_Date}"` +
              "," +
              `"${item.Approver}"` +
              "," +
              `"${item.Total_x0020_Cost}"` +
              "," +
              "\n";
  

              const itemIDs = item.ID;
           // console.log("Test itemIDs", itemIDs);
                  /////
                  const newresponse = await sp.web.lists
                  .getByTitle("MRF")
                  .items.getById(itemIDs)
                  .roleAssignments();
                 // console.log("newresponse",newresponse)
                  //const varperms = await sp.web.getUserEffectivePermissions("i:0#.f|membership|navpreet.kaur1@peelsb.com");
                  const var1perms = await sp.web.getUserEffectivePermissions("i:0#.f|membership|hilary.taylor@peelsb.com");
                  //console.log("varperms",varperms)
                  //console.log("var1perms",var1perms)

                  const varobj = await sp.web.userHasPermissions("i:0#.f|membership|hilary.taylor@peelsb.com", PermissionKind.EditListItems);
                 // console.log("varobj",varobj)
                  /////
            const list = sp.web.lists.getByTitle("MRF");
            const i = list.items.top(200).getById(itemIDs).update({
              Status: "Exported",
              UploadID: newRandNum,
            });
          }
         // console.log("Its new test", varContent);
          const newUpload = sp.web
            .getFolderByServerRelativeUrl(newURL)
            .files.add(varFileName, File, true)
            .then(async (data) => {
            //  console.log("Test hello0000", data);
              Dialog.alert("Test Generated and uploaded file successfully ");
              const newVar = sp.web
                .getFileByServerRelativeUrl(`${newURL}/${varFileName}`)
                .setContent(varContent);
              data.file.getItem().then((item) => {
                item.update({
                  Title: "MileageTitle",
                });
              });
            });
        });
      //
    }
/********************************** Test Function ends here *********************************************************************************************/

  /**** function to update Status and UploadID on Completed button****/
  private async updateListItem(itemID: any) {
    //sp.setup({spfxContext: this.context });

    // const web= Web(this.context.pageContext.web);

    //console.log("sp", sp);
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.top(200).getById(itemID).update({
      Status: "Completed", //column to be updated in the list
      // UploadID: setNewDate,
    });
  }

  /**** function to update Status and UploadID on Pending button****/
  private async updateListItemPending(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Not Started", //column to be updated in the list
      // UploadID: newRandNum,
    });
  }

  /**** function to update Status and UploadID on Pending button****/
  private async updateListItemDeferred(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Deferred", //column to be updated in the list
      // UploadID: newRandNum,
    });
  }

  /**** function to update Status and UploadID on Uplaod button****/
  private async updateListItemUpload(itemID: any) {
    let list = sp.web.lists.getByTitle("MRF");
    const i = await list.items.getById(itemID).update({
      Status: "Exported", //column to be updated in the list
      //UploadID: newRandNum,
    });
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      /********************************Generate Upload File -FIS Button---------------------------------------****************************/
      case "COMMAND_1":
        this.viewData();
        /*  setTimeout(function(){
          location.reload();
       }, 5000);*/

        break;
      //Dialog.alert("File uploaded successfully");

      /********************************Completed Button-----------------------------------------****************************/
      case "COMMAND_2": //Completed Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItem(listId).then(() => {
              //  Dialog.alert("Status updated to completed successfully ");
              location.reload();
            });
          });
        }

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
            this.updateListItemPending(listId).then(() => {
              //Dialog.alert("Status updated to pending successfully ");
              location.reload();
            });
          });
        }

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
            this.updateListItemDeferred(listId).then(() => {
              //Dialog.alert("Status updated to deferred successfully ");
              location.reload();
            });
          });
        }

        // Dialog.alert("This is Deffered button");
        break;
      /********************************Upload Button-----------------------------------------****************************/
      case "COMMAND_5": //Upload button
      /*  if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listId = ` ${row.getValueByName("ID")}`;
            //console.log("listId", listId);
            this.updateListItemUpload(listId);
          });
        }*/
        //Dialog.alert("File uploaded successfully");
        this.viewDataTest();

        break;

    }
  }
}
