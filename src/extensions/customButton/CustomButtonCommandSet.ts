/****spfx imports *****/
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
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


export default class CustomButtonCommandSet extends BaseListViewCommandSet<ICustomButtonCommandSetProperties> {
  public context: any;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized CustomButtonCommandSet");
    sp.setup({
      spfxContext: this.context,
      sp: { baseUrl: this.context.pageContext.web.absoluteUrl },
    });
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    // hide commands for other lists that are not MRF
    if (window.location.href.toLowerCase().indexOf('/mrf/') === -1){
      this.tryGetCommand("COMMAND_Upload").visible = false;
      this.tryGetCommand("COMMAND_Completed").visible = false;
      this.tryGetCommand("COMMAND_Pending").visible = false;
      this.tryGetCommand("COMMAND_Deferred").visible = false;
      //this.tryGetCommand("COMMAND_TestUpload").visible = false;
    }
  }

  /**** Get data from particular view of a list ****/
  private async viewData() {

    const newRandNum = new Date().toISOString().replace(/[^0-9]+/g, "") + Math.floor(Math.random() * 999);
    const url = this.context.pageContext.web.serverRelativeUrl;
    const folderName = "FileUpload";
    const newURL = url + "/" + folderName;
    let varContent = "";
    const varFileName = "MileageAPFile_" + `${newRandNum}.csv`;

    // Read or Write to a list
    const executeJson = async (endpointUrl, payload?) => {
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

    const getListItems = async (webUrl: string, listTitle: string, queryText: string) => {
      const viewXml = "<View><Query>" + queryText + "</Query></View>";
      const endpointUrl = webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getitems?$top=200";
      const queryPayload = { query: { ViewXml: viewXml } };
      return executeJson(endpointUrl, queryPayload);
    };

    const getListViewItems = async (webUrl: string, listTitle: string, viewTitle: string) => {
      const endpointUrl = webUrl + "/_api/web/lists/getByTitle('MRF')/Views/getbytitle('" + viewTitle +"')/ViewQuery";
      return executeJson(endpointUrl, null)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then((data) => {
          const viewQuery = data.value;
          return getListItems(webUrl, listTitle, viewQuery);
        });
    };
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
        }

        for (let j = 0; j <response.value.length; j++) {
          const itemIDs = response.value[j].ID //.ID;
          const list = sp.web.lists.getByTitle("MRF");
          const i = list.items.top(200).getById(itemIDs).update({
            Status: "Exported",
            UploadID: newRandNum,
          });
        }
        const newUpload = sp.web
          .getFolderByServerRelativeUrl(newURL)
          .files.add(varFileName, File, true)
          .then(async (data) => {
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
  }


  /**** Update Status (Completed, Not Started, Deferred, Exported) and UploadID ****/
  private async updateListItem(itemID: any, status: string, uploadId?: string) {
    const body = uploadId ? {Status: status, UploadID: uploadId} : {Status: status};
    await sp.web.lists.getByTitle("MRF").items.getById(itemID).update(body);
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_Upload": // Generate Upload File -FIS Button
        this.viewData();
        break;
      case "COMMAND_Completed": //Completed Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listItemId = ` ${row.getValueByName("ID")}`;
            this.updateListItem(listItemId, 'Completed').then(() => {
              location.reload();
            });
          });
        }
        break;
      case "COMMAND_Pending": //Pending Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listItemId = ` ${row.getValueByName("ID")}`;
            this.updateListItem(listItemId, 'Not Started').then(() => {
              location.reload();
            });
          });
        }
        break;
      case "COMMAND_Deferred": //Deferred Button
        if (event.selectedRows.length > 0) {
          // Check the selected rows
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const listItemId = ` ${row.getValueByName("ID")}`;
            this.updateListItem(listItemId, 'Deferred').then(() => {
              location.reload();
            });
          });
        }
        break;
    }
  }
}
