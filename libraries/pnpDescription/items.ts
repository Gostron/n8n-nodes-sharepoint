import { spcollection, spqueryable } from "./_common"
import { item } from "./item"

export const items: PnpDescriptionNode = {
  label: "Items",
  key: "items",
  description: "Gets the collection of all items that are contained in the list",
  ...spcollection,
  children: [
    {
      label: "Get by ID",
      key: "getById",
      description: "Gets an item from the collection by id",
      callable: true,
      parameters: [
        ...(spqueryable.parameters || []),
        {
          label: "ID",
          key: "id",
          description: "The Id of the item (number)",
          type: "number",
        },
      ],
      children: item,
    },
    {
      label: "Add",
      key: "add",
      description: "Adds a new item to the collection",
      parameters: [
        {
          label: "Properties",
          key: "properties",
          description: "The new item's properties",
          type: "object",
          required: true,
        },
      ],
    },
  ],
}
