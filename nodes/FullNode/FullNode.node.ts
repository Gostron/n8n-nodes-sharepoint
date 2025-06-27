import type { IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { executeWithErrorHandling, getSharePointConfig } from "../../libraries/sharepointUtils"
import { isGuid } from "../../libraries/guid"
import { properties } from "./properties"

export class FullNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "SharePoint All PnP API",
    name: "fullNode",
    group: ["transform"],
    version: 1,
    description: "Access all PnP API methods for SharePoint using the PnPjs 'sp' object",
    defaults: {
      name: "SharePoint All PnP API",
    },
    inputs: [NodeConnectionType.Main],
    outputs: [NodeConnectionType.Main],
    credentials: [
      {
        name: "microsoftSharePointAppOnlyApi",
        required: true,
      },
    ],
    properties,
  }

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData()
    const returnData: INodeExecutionData[] = []

    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      const siteUrl = this.getNodeParameter("siteUrl", itemIndex, "") as string
      const listNameOrId = this.getNodeParameter("listNameOrId", itemIndex, "") as string
      const itemId = Number(this.getNodeParameter("itemId", itemIndex, "") as string)

      const { sp } = await getSharePointConfig(this, siteUrl)

      const json = await executeWithErrorHandling.call(
        this,
        async () => {
          const list = isGuid(listNameOrId) ? sp.web.lists.getById(listNameOrId) : sp.web.lists.getByTitle(listNameOrId)
          return list.items.getById(itemId)()
        },
        itemIndex,
      )

      returnData.push(json)
    }

    return [returnData]
  }
}
