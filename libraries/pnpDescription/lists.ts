import { spcollection, spqueryable } from "./_common"
import { list } from "./list"

export const lists: PnpDescriptionNode = {
  label: "Lists",
  key: "lists",
  description: "Gets the collection of all lists that are contained in the Web site",
  ...spcollection,
  children: [
    {
      label: "Get by ID",
      key: "getById",
      description: "Gets a list from the collection by guid id",
      callable: true,
      parameters: [
        ...(spqueryable.parameters || []),
        {
          label: "ID",
          key: "id",
          description: "The Id of the list (GUID)",
          type: "string",
        },
      ],
      children: list,
    },
    {
      label: "Get by Title",
      key: "getByTitle",
      description: "Gets a list from the collection by title",
      callable: true,
      parameters: [
        ...(spqueryable.parameters || []),
        {
          label: "Title",
          key: "title",
          description: "The title of the list",
          type: "string",
        },
      ],
      children: list,
    },
    {
      label: "Add",
      key: "add",
      description: "Adds a new list to the collection",
      parameters: [
        {
          label: "Title",
          key: "title",
          description: "The new list's title",
          type: "string",
          required: true,
        },
        {
          label: "Description",
          key: "description",
          description: "The new list's description",
          type: "string",
        },
        {
          label: "Template",
          key: "template",
          description: "The list template value",
          type: "number",
        },
        {
          label: "Enable content types",
          key: "enableContentTypes",
          description:
            "If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled",
          type: "boolean",
        },
        {
          label: "Additional settings",
          key: "additionalSettings",
          description: "Will be passed as part of the list creation body or used to update an existing list",
          type: "object",
        },
      ],
    },
  ],
}
