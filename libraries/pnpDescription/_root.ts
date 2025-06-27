import { spqueryable } from "./_common"
import { web } from "./web"

const pnpDescription: PnpDescriptionNode[] = [
  web,
  {
    label: "Site",
    key: "site",
    ...spqueryable,
    children: [
      {
        label: "Get Document Libraries",
        key: "getDocumentLibraries",
        description: "Gets the document libraries on a site. Static method. (SharePoint Online only)",
        parameters: [
          {
            label: "Absolute web URL",
            key: "absoluteWebUrl",
            description: "The absolute url of the web whose document libraries should be returned",
            type: "string",
            required: true,
          },
        ],
      },
    ],
  },
]

export default pnpDescription
