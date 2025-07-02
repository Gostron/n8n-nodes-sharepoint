import { mkdir, readFile, writeFile } from "fs/promises"
import pnpDescription from "../libraries/pnpDescription/_root"
import { camelCase } from "lodash"

const writeProperties = async (node: PnpDescriptionNode, filePath: string) => {
  const output = (await readFile("./provisioning/_templateProperties.ts", "utf8")).replace(
    /<properties>/g,
    node.parameters ? JSON.stringify(node.parameters, null, 2) : "[]",
  )

  await writeFile(filePath, output, "utf8")
}

const writeNode = async (node: PnpDescriptionNode, filePath: string) => {
  const output = (await readFile("./provisioning/_templateNode.ts", "utf8"))
    .replace(/<node-display-name>/g, node.label)
    .replace(/<node-name>/g, camelCase(node.key))
    .replace(/<node-group>/g, "SharePoint")
    .replace(/<node-description>/g, node.description || "-")
    .replace(/<node-default-name>/g, camelCase(node.key))
    .replace(
      /<get-parameters-snippet>/g,
      node.parameters ? `const parameters = this.getNodeParameter("parameters", itemIndex, []) as any[];` : "",
    )
    .replace(/<get-output-snippet>/g, `const json = await sp.${camelCase(node.key)}(parameters)\nreturn json`)
  await writeFile(filePath, output, "utf8")
}

const processNode = async (node: PnpDescriptionNode) => {
  if (node.callable) {
    const nodeName = camelCase(`Get ${node.label} node`)
    console.log(`Processing node: ${nodeName}`)
    await mkdir(`./nodes/${nodeName}`).catch(() => {})
    const nodePath = `./nodes/${nodeName}/${nodeName}.node.ts`
    const propertiesPath = `./nodes/${nodeName}/properties.ts`
    await writeNode(node, nodePath)
    await writeProperties(node, propertiesPath)
  }
}

export const main = async () => {
  console.log(pnpDescription)
  await writeFile("./libraries/pnpDescription/gen_full.json", JSON.stringify(pnpDescription, null, 2), "utf8")
  for (const node of pnpDescription) {
    await processNode(node)
  }
}

main()
  .then(() => {
    console.log("Execution completed successfully.")
  })
  .catch((error) => {
    console.error("Execution failed:", error)
    process.exit(1)
  })
