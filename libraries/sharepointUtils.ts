import type { IExecuteFunctions } from "n8n-workflow"
import { NodeOperationError } from "n8n-workflow"
import { getPnp, IPnpConfig } from "./pnp/pnp"

type Unpromisify<T> = T extends Promise<infer U> ? U : T

type PnpResult = Unpromisify<ReturnType<typeof getPnp>>

export async function getSharePointConfig(
  context: IExecuteFunctions,
  alternateSiteUrl?: string,
): Promise<
  PnpResult & {
    config: IPnpConfig
  }
> {
  const credentials = await context.getCredentials("microsoftSharePointAppOnlyApi")
  const config: IPnpConfig = {
    clientCertificatePrivateKey: (() => {
      let value = (credentials.clientCertificatePrivateKey as string).replace(/\s+/g, "")
      value = value.replace("-----BEGINRSAPRIVATEKEY-----", "").replace("-----ENDRSAPRIVATEKEY-----", "")
      value = value.match(/.{1,64}/g)?.join("\n") || ""
      value = `-----BEGIN RSA PRIVATE KEY-----\n${value}\n-----END RSA PRIVATE KEY-----`
      return value
    })(),
    clientCertificateThumbprint: credentials.clientCertificateThumbprint as string,
    clientId: credentials.clientId as string,
    tenantId: credentials.tenantId as string,
    siteUrl: alternateSiteUrl || `https://${credentials.tenantName as string}-admin.sharepoint.com`,
  }
  return { config, ...(await getPnp(config)) }
}

export async function executeWithErrorHandling(this: IExecuteFunctions, fn: () => Promise<any>, itemIndex: number) {
  try {
    return await fn()
  } catch (error) {
    if (this.continueOnFail()) {
      return {
        error,
        json: this.getInputData(itemIndex)[0]?.json ?? {},
      }
    } else {
      if (error.context) {
        error.context.itemIndex = itemIndex
        throw error
      }
      throw new NodeOperationError(this.getNode(), error, { itemIndex })
    }
  }
}
