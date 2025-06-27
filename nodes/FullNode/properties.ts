import type { INodeProperties, INodePropertyOptions, NodeParameterValueType, NodePropertyTypes } from "n8n-workflow"
import { camelCase, deburr, mapValues } from "lodash"

const displayToName = (name: string) => camelCase(deburr(name))
const p = (
  displayName: string,
  type: NodePropertyTypes,
  defaultValue: NodeParameterValueType,
  options?: string[] | undefined,
  show?: Record<string, string | string[]> | undefined,
): INodeProperties => ({
  displayName,
  name: displayToName(displayName),
  type,
  default: defaultValue,
  options: options ? options?.map<INodePropertyOptions>((option) => ({ name: option, value: option })) : undefined,
  displayOptions: show ? { show: mapValues(show, (v) => (Array.isArray(v) ? v : [v])) } : undefined,
})

export const properties: INodeProperties[] = [
  {
    displayName: "Site URL",
    name: "siteUrl",
    type: "string",
    default: "",
    placeholder: "https://contoso.sharepoint.com/sites/example",
    description: "The base URL of the SharePoint site",
    required: true,
  },
  p("Level 1", "options", "", ["web", "site"]),
  p(
    "Level 2",
    "options",
    "",
    [
      "webs",
      "allProperties",
      "update",
      "applyTheme",
      "applyWebTempalte",
      "getChanges",
      "siteGroups",
      "associatedOwnerGroup",
      "associatedMemberGroup",
      "associatedVisitorGroup",
      "contentTypes",
      "fields",
      "getFileByServerRelativePath",
      "getFolderByServerRelativePath",
      "getFileById",
      "getFolderById",
      "siteUsers",
      "getUserById",
      "ensureUser",
    ],
    { level1: "web" },
  ),
]
