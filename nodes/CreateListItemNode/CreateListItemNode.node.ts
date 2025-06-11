import type { IDataObject, IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { getSharePointConfig, executeWithErrorHandling } from "../../libraries/sharepointUtils"

export class CreateListItemNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Create List Item",
    name: "createListItemNode",
    group: ["transform"],
    version: 1,
    description: "Create a list item in SharePoint with arbitrary fields",
    defaults: {
      name: "Create List Item",
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
        displayName: "List Name or ID",
        name: "listName",
        type: "string",
        default: "",
        placeholder: "Enter the SharePoint list name or ID",
        description: "The name or ID of the SharePoint list where the item will be created",
        required: true,
      },
      {
        displayName: "Item Fields",
        name: "itemFields",
        type: "json",
        default: "{}",
        placeholder: '{ "Title": "New Item", "Field1": "Value1" }',
        description: "JSON object representing the fields and values for the new list item",
        required: true,
      },
    ],
  }

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData()
    const returnData: INodeExecutionData[] = []

    const { sp } = await getSharePointConfig(this)
    try {
      for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
        const result = await executeWithErrorHandling.call(
          this,
          async () => {
            const listName = this.getNodeParameter("listName", itemIndex, "") as string
            const itemFields = this.getNodeParameter("itemFields", itemIndex, {}) as IDataObject

            const list = sp.web.lists.getByTitle(listName)
            const addedItem = await list.items.add(itemFields)

            return addedItem.data
          },
          itemIndex,
        )

        returnData.push({
          json: result.error ? result : result,
        })
      }
    } catch (error) {
      if (this.continueOnFail()) {
        returnData.push({
          json: {
            input: this.getInputData()[0].json,
          },
          error,
        })
      } else {
        if (error.context) {
          error.context.itemIndex = 0
          throw error
        }
        throw error
      }
    }

    return [returnData]
  }
}
