import type { IDataObject, IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { getSharePointConfig, executeWithErrorHandling } from "../../libraries/sharepointUtils"

export class GetWebPropertiesNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Get Web Properties",
    name: "getWebPropertiesNode",
    group: ["transform"],
    version: 1,
    description: "Basic Get Web Properties Node",
    defaults: {
      name: "Get Web Properties Node",
    },
    inputs: [NodeConnectionType.Main],
    outputs: [NodeConnectionType.Main],
    usableAsTool: true,
    credentials: [
      {
        name: "microsoftSharePointAppOnlyApi",
        required: true,
      },
    ],
    properties: [
      // Node properties which the user gets displayed and
      // can change on the node.
      {
        displayName: "My String",
        name: "myString",
        type: "string",
        default: "",
        placeholder: "Placeholder value",
        description: "The description text",
      },
    ],
  }

  // The function below is responsible for actually doing whatever this node
  // is supposed to do. In this case, we're just appending the `myString` property
  // with whatever the user has entered.
  // You can make async calls and use `await`.
  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const returnData: INodeExecutionData[] = []

    try {
      const { sp } = await getSharePointConfig(this)
      const web = await executeWithErrorHandling.call(this, () => sp.web(), 0)

      if (web.error) {
        return [[web]]
      }

      return [
        [
          {
            json: web as unknown as IDataObject,
          },
        ],
      ]
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
