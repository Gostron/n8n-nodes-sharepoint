import type { IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { executeWithErrorHandling, getSharePointConfig } from "../../libraries/sharepointUtils"
import { properties } from "./properties"

export class FullNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "<node-display-name>",
    name: "<node-name>",
    group: ["<node-group>"],
    version: 1,
    description: "<node-description>",
    defaults: {
      name: "<node-default-name>",
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
      const { sp } = await getSharePointConfig(this, siteUrl)

			<get-parameters-snippet>

			const output = await executeWithErrorHandling.call(
				this,
				async () => {
					<get-output-snippet>
				},
				itemIndex,
			)

      returnData.push(json)
    }

    return [returnData]
  }
}
