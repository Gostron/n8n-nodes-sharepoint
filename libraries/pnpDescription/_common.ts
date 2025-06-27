export const spqueryable: Pick<PnpDescriptionNode, "parameters" | "callable"> = {
  callable: true,
  parameters: [
    {
      label: "Select",
      key: "select",
      description: "Choose which fields to return",
      type: "string",
      spread: true,
      chain: true,
    },
    {
      label: "Expand",
      key: "expand",
      description: "Choose which fields to expand",
      type: "string",
      spread: true,
      chain: true,
    },
  ],
}

export const spcollection: Pick<PnpDescriptionNode, "parameters" | "callable"> = {
  callable: true,
  parameters: [
    ...(spqueryable.parameters || []),
    {
      label: "Filter",
      key: "filter",
      description:
        "Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)",
      type: "string",
      chain: true,
    },
    {
      label: "OrderBy",
      key: "orderBy",
      description: "Order based on the supplied fields add ` asc` or ` desc` to the field name",
      type: "string",
      chain: true,
    },
    {
      label: "Skip",
      key: "skip",
      description: "Skip the specified number of items",
      type: "number",
      chain: true,
    },
    {
      label: "Top",
      key: "top",
      description: "Limits the query to only return the specified number of items",
      type: "number",
      chain: true,
    },
  ],
}
