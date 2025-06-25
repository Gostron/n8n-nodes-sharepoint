import type { IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { executeWithErrorHandling, getSharePointConfig } from "../../libraries/sharepointUtils"
import { isGuid } from "../../libraries/guid"

export class GetListItemNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "SharePoint Get List Item",
    name: "getListItemNode",
    group: ["transform"],
    version: 1,
    description: "Get a list item from SharePoint using the PnPjs 'sp' object",
    defaults: {
      name: "SharePoint Get List Item",
    },
    inputs: [NodeConnectionType.Main],
    outputs: [NodeConnectionType.Main],
    credentials: [
      {
        name: "microsoftSharePointAppOnlyApi",
        required: true,
      },
    ],
    properties: [
      {
        displayName: "Site URL",
        name: "siteUrl",
        type: "string",
        default: "",
        placeholder: "https://contoso.sharepoint.com/sites/example",
        description: "The base URL of the SharePoint site",
        required: true,
      },
      {
        displayName: "List Name or ID",
        name: "listNameOrId",
        type: "string",
        default: "",
        placeholder: "e.g., My List or 12345678-1234-1234-1234-123456789012",
        description: "The name or ID of the SharePoint list to query",
        required: true,
      },
      {
        displayName: "Item ID",
        name: "itemId",
        type: "string",
        default: "",
        placeholder: "e.g., 1",
        description: "The ID of the SharePoint list item to retrieve",
        required: true,
      },
    ],
  }

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData()
    const returnData: INodeExecutionData[] = []

    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      const siteUrl = this.getNodeParameter("siteUrl", itemIndex, "") as string
      const listNameOrId = this.getNodeParameter("listNameOrId", itemIndex, "") as string
      const itemId = Number(this.getNodeParameter("itemId", itemIndex, "") as string)

      const { sp } = await getSharePointConfig(this, siteUrl)

      console.log(`Getting item ${itemId} from list "${listNameOrId}" at site "${siteUrl}"`)

      const result = await executeWithErrorHandling.call(
        this,
        async () => {
          // SPQueryable([sp.web, siteUrl], "/_api/web")
          const list = isGuid(listNameOrId) ? sp.web.lists.getById(listNameOrId) : sp.web.lists.getByTitle(listNameOrId)

          console.log(`List: ${list.toUrl()}`)
          return list.items.getById(itemId)()
        },
        itemIndex,
      )

      console.log(`Result: ${JSON.stringify(result, null, 2)}`)

      returnData.push({
        json: result.error ? result : result,
      })
    }

    return [returnData]
  }
}
