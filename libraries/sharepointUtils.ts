import type { IExecuteFunctions } from "n8n-workflow"
import { NodeOperationError } from "n8n-workflow"
import { getPnp, IPnpConfig } from "./pnp/pnp"

type Unpromisify<T> = T extends Promise<infer U> ? U : T

type PnpResult = Unpromisify<ReturnType<typeof getPnp>>

export async function getSharePointConfig(
  context: IExecuteFunctions,
  siteUrl: string,
): Promise<
  PnpResult & {
    config: IPnpConfig
  }
> {
  const credentials = await context.getCredentials("microsoftSharePointAppOnlyApi")

  if (!credentials.clientCertificatePrivateKey) {
    throw new NodeOperationError(context.getNode(), "Client certificate private key is required", {
      description: "Please provide the client certificate private key in the credentials.",
    })
  }
  if (!credentials.clientCertificateThumbprint) {
    throw new NodeOperationError(context.getNode(), "Client certificate thumbprint is required", {
      description: "Please provide the client certificate thumbprint in the credentials.",
    })
  }
  if (!credentials.clientId) {
    throw new NodeOperationError(context.getNode(), "Client ID is required", {
      description: "Please provide the client ID in the credentials.",
    })
  }
  if (!credentials.tenantName) {
    throw new NodeOperationError(context.getNode(), "Tenant name is required", {
      description: "Please provide the tenant name in the credentials.",
    })
  }
  if (!credentials.tenantId) {
    throw new NodeOperationError(context.getNode(), "Tenant ID is required", {
      description: "Please provide the tenant ID in the credentials.",
    })
  }

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
    tenantName: credentials.tenantName as string,
  }
  // context.logger.info(JSON.stringify(config, null, 2))
  return { config, ...(await getPnp.call(context, config, siteUrl)) }
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
