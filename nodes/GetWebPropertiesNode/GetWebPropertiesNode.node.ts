import type { IDataObject, IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType, NodeOperationError } from "n8n-workflow"
import { getPnp, IPnpConfig } from "../../libraries/pnp/pnp"

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
    const items = this.getInputData()

    // let item: INodeExecutionData
    // let myString: string

    // // Iterates over all input items and add the key "myString" with the
    // // value the parameter "myString" resolves to.
    // // (This could be a different value for each item in case it contains an expression)
    // for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
    //   try {
    //     myString = this.getNodeParameter("myString", itemIndex, "") as string
    //     item = items[itemIndex]

    //     item.json.myString = myString
    //   } catch (error) {
    //     // This node should never fail but we want to showcase how
    //     // to handle errors.
    //     if (this.continueOnFail()) {
    //       items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex })
    //     } else {
    //       // Adding `itemIndex` allows other workflows to handle this error
    //       if (error.context) {
    //         // If the error thrown already contains the context property,
    //         // only append the itemIndex
    //         error.context.itemIndex = itemIndex
    //         throw error
    //       }
    //       throw new NodeOperationError(this.getNode(), error, {
    //         itemIndex,
    //       })
    //     }
    //   }
    // }
    const credentials = await this.getCredentials("microsoftSharePointAppOnlyApi")
    const config: IPnpConfig = {
      clientCertificatePrivateKey: (() => {
        let value = (credentials.clientCertificatePrivateKey as string).replace(/\s+/g, "")
        value = value.replace("-----BEGINRSAPRIVATEKEY-----", "").replace("-----ENDRSAPRIVATEKEY-----", "")
        // Split value in lines of 64 characters
        value = value.match(/.{1,64}/g)?.join("\n") || ""
        // Add the header and footer
        value = `-----BEGIN RSA PRIVATE KEY-----\n${value}\n-----END RSA PRIVATE KEY-----`
        return value
      })(),
      clientCertificateThumbprint: credentials.clientCertificateThumbprint as string,
      clientId: credentials.clientId as string,
      tenantId: credentials.tenantId as string,
      siteUrl: credentials.siteUrl as string,
    }

    try {
      const { sp } = await getPnp(config)
      const web = await sp.web()

      return [
        [
          {
            json: web as unknown as IDataObject,
          },
        ],
      ]
    } catch (error) {
      // This node should never fail but we want to showcase how
      // to handle errors.
      if (this.continueOnFail()) {
        items.push({
          json: {
            input: this.getInputData()[0].json,
            config,
          },
          error,
        })
      } else {
        // Adding `itemIndex` allows other workflows to handle this error
        if (error.context) {
          // If the error thrown already contains the context property,
          // only append the itemIndex
          // error.context.config = config
          error.context.itemIndex = 1
          throw error
        }
        throw new NodeOperationError(this.getNode(), error, {
          itemIndex: 1,
          // description: JSON.stringify(config),
        })
      }
    }
    return [items]
  }
}
