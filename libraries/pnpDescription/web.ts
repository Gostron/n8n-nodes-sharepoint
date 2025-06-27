import { spqueryable } from "./_common"
import { lists } from "./lists"

export const web: PnpDescriptionNode = {
  label: "Web",
  key: "web",
  description: "Access to the current web instance",
  ...spqueryable,
  children: [lists],
}
