import { DefaultHeaders, DefaultInit } from "@pnp/sp/presets/all.js"
import { NodeFetchWithRetry, SPDefault } from "@pnp/nodejs/index.js"
import { TimelinePipe } from "@pnp/core/index.js"
import { DefaultParse, Queryable } from "@pnp/queryable/index.js"
import { OnlineAddinOnly } from "./OnlineAddinOnly.js"

const getAuth = (url: string, options: { clientId: string; clientSecret: string }) => {
  return new OnlineAddinOnly(url, options).getAuth()
}

/** Authenticates SharePoint using Client ID / Client Secret for App only. This method is only supported and is the only method to use SP Add In rights. */
export const SPNodeAuth: (clientId: string, clientSecret: string) => TimelinePipe<Queryable> = (
  clientId,
  clientSecret,
) => {
  return (instance) => {
    instance.on.auth(async (url, init) => {
      const result = await getAuth(url.toString(), { clientId, clientSecret })
      init.headers = {
        ...init.headers,
        ...result.headers,
      }

      return [url, init]
    })

    return instance
  }
}

/** Full Timeline comporting the custom legacy auth and the default timeline. */
export const SPAddIn: (clientId: string, clientSecret: string) => TimelinePipe<Queryable> = (
  clientId,
  clientSecret,
) => {
  return (instance) =>
    instance.using(
      SPNodeAuth(clientId, clientSecret),
      DefaultHeaders(),
      DefaultInit(),
      NodeFetchWithRetry(),
      DefaultParse(),
    )
}

export const SPAddInCertificate: (
  clientId: string,
  certificate: string,
  thumbprint: string,
  tenantId: string,
  tenantName: string,
) => TimelinePipe<Queryable> = (clientId, certificate, thumbprint, tenantId, tenantName) => {
  return (instance) =>
    instance.using(
      SPDefault({
        msal: {
          config: {
            auth: {
              authority: `https://login.microsoftonline.com/${tenantId}`,
              clientId: clientId,
              clientCertificate: {
                privateKey: certificate,
                thumbprint: thumbprint,
              },
            },
          },
          scopes: [`${tenantName}/.default`],
        },
      }),
      DefaultHeaders(),
      DefaultInit(),
      NodeFetchWithRetry(),
      DefaultParse(),
    )
}
