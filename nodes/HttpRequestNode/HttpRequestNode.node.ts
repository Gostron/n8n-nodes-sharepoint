import type { IDataObject, IExecuteFunctions, INodeExecutionData, INodeType, INodeTypeDescription } from "n8n-workflow"
import { NodeConnectionType } from "n8n-workflow"
import { getSharePointConfig, executeWithErrorHandling } from "../../libraries/sharepointUtils"

export class HttpRequestNode implements INodeType {
  description: INodeTypeDescription = {
    displayName: "SharePoint HTTP Request",
    name: "httpRequestNode",
    group: ["transform"],
    version: 1,
    description: "Perform an HTTP request to SharePoint using the PnPjs 'sp' object",
    defaults: {
      name: "SharePoint HTTP Request",
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
        displayName: "HTTP Method",
        name: "method",
        type: "options",
        options: [
          { name: "DELETE", value: "DELETE" },
          { name: "GET", value: "GET" },
          { name: "PATCH", value: "PATCH" },
          { name: "POST", value: "POST" },
        ],
        default: "GET",
        description: "HTTP method to use for the request",
        required: true,
      },
      {
        displayName: "Site Relative Endpoint URL",
        name: "relativeUrl",
        type: "string",
        default: "",
        placeholder: "/_api/web/lists",
        description: "The endpoint path relative to the site URL",
        required: true,
      },
      {
        displayName: "Query Parameters",
        name: "query",
        type: "json",
        default: "",
        placeholder: '{ "$top": 5 }',
        description: "Query parameters as a JSON object (optional)",
      },
      {
        displayName: "Headers",
        name: "headers",
        type: "json",
        default: "",
        placeholder: '{ "Accept": "application/json" }',
        description: "Request headers as a JSON object (optional)",
      },
      {
        displayName: "Body",
        name: "body",
        type: "json",
        default: "",
        placeholder: '{ "Title": "Sample" }',
        description: "Request body as a JSON object (optional, for POST/PATCH/PUT)",
      },
    ],
  }

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData()
    const returnData: INodeExecutionData[] = []

    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      const tryParse = (parameterName: string): IDataObject | undefined => {
        try {
          const value = this.getNodeParameter(parameterName, itemIndex, "{}") as string
          return JSON.parse(value)
        } catch (error) {
          this.logger?.error(`Failed to parse JSON for ${parameterName}`, { error })
          return undefined
        }
      }
      const siteUrl = this.getNodeParameter("siteUrl", itemIndex, "") as string
      const method = this.getNodeParameter("method", itemIndex, "GET") as string
      const relativeUrl = this.getNodeParameter("relativeUrl", itemIndex, "") as string
      const query = tryParse("query") || {}
      const headers = tryParse("headers") || {}
      const body = tryParse("body")

      const queryString =
        query && Object.keys(query).length > 0 ?
          "?" +
          Object.keys(query)
            .map((key) => encodeURIComponent(key) + "=" + encodeURIComponent(String(query[key])))
            .join("&")
        : ""

      const { spOdata: sp, spMethods, spQueryable } = await getSharePointConfig(this, siteUrl)

      const fullUrl = `${siteUrl}${relativeUrl}${queryString}`
      const fetchOptions: RequestInit = {
        headers: headers as Record<string, string>,
        body:
          body && ["POST", "PATCH", "PUT"].includes(method) ?
            typeof body === "string" ?
              body
            : JSON.stringify(body)
          : undefined,
      }

      const result = await executeWithErrorHandling.call(
        this,
        async () => {
          // SPQueryable([sp.web, siteUrl], "/_api/web")
          const methodLower = method.toLowerCase() as "get" | "post" | "delete" | "patch"
          const queryable = spQueryable([sp.web, fullUrl])
          console.log(`[SharePoint HTTP Node][${methodLower}] ${queryable.toRequestUrl()}`, fetchOptions)
          return spMethods[methodLower](queryable, fetchOptions)
        },
        itemIndex,
      )

      returnData.push({
        json: result.error ? result : result,
      })
    }

    return [returnData]
  }
}
