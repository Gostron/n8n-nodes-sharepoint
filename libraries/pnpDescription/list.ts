import { items } from "./items"

export const list: PnpDescriptionNode[] = [
  {
    label: "Get Items by CAML Query",
    key: "getItemsByCAMLQuery",
    description: "Returns the collection of items in the list based on the provided CAML query",
    parameters: [
      {
        label: "CAML Query",
        key: "query",
        description: "A query that is performed against the list",
        type: "object",
        required: true,
      },
      {
        label: "Expands",
        key: "expands",
        description: "An expanded array of items that contains fields to expand in the CAML query",
        type: "string",
        spread: true,
      },
    ],
  },
  {
    label: "Update",
    key: "update",
    description: "Updates this list instance with the supplied properties",
    parameters: [
      {
        label: "Properties",
        key: "properties",
        description: "A plain object hash of values to update for the list",
        type: "object",
        required: true,
      },
      {
        label: "ETag",
        key: "eTag",
        description: "Value used in the IF-Match header, by default '*'",
        type: "string",
      },
    ],
  },
  items,
]
