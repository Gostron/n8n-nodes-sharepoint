import { type SPFI } from "@pnp/sp/presets/all.js"
import { type GraphFI } from "@pnp/graph/presets/all.js"

const getESMPnP = () => {
  return Promise.all([
    eval(`import("@pnp/sp/presets/all.js")`),
    eval(`import("@pnp/graph/presets/all.js")`),
    eval(`import("@pnp/nodejs/index.js")`),
    eval(`import("@pnp/queryable")`),
  ]).then(([sp, graph, nodejs, queryable]) => {
    return {
      spfi: sp.spfi,
      graphfi: graph.graphfi,
      SPDefault: nodejs.SPDefault,
      GraphDefault: nodejs.GraphDefault,
      InjectHeaders: queryable.InjectHeaders,
    }
  })
}

export type IPnpConfigCommon = {
  siteUrl: string
  clientId: string
  tenantId: string
}

export type IPnpConfigCertificate = IPnpConfigCommon & {
  clientCertificatePrivateKey: string
  clientCertificateThumbprint: string
}
export type IPnpConfigSecret = IPnpConfigCommon & {
  clientSecret: string
}

export type IPnpConfig = IPnpConfigCertificate | IPnpConfigSecret
type IPnpConfigInternal = IPnpConfigCommon & Partial<IPnpConfigSecret> & Partial<IPnpConfigCertificate>

export const getPnp = async (config: IPnpConfig) => {
  const { spfi, graphfi, SPDefault, GraphDefault, InjectHeaders } = await getESMPnP()

  const NoOdata = (instance: any) => InjectHeaders({ Accept: "application/json;odata=nometadata" })(instance)
  const adaptiveAuth = (config: IPnpConfig, mode: "sp" | "graph") => {
    const { clientId, siteUrl, tenantId, clientCertificatePrivateKey, clientCertificateThumbprint, clientSecret } =
      config as IPnpConfigInternal
    const authCertificate = {
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientId: clientId,
      clientSecret: clientSecret,
      clientCertificate:
        clientCertificatePrivateKey && clientCertificateThumbprint ?
          {
            privateKey: clientCertificatePrivateKey,
            thumbprint: clientCertificateThumbprint,
          }
        : undefined,
    }
    if (mode === "sp") {
      const tenantName = siteUrl.split("/").slice(0, 3).join("/")
      if (clientSecret)
        throw new Error(
          "Client Secret is not supported for SharePoint Add In authentication. Use client certificate instead.",
        )
      else
        return SPDefault({
          msal: {
            config: { auth: authCertificate },
            scopes: [`${tenantName}/.default`],
          },
        })
    } else if (mode === "graph") {
      return GraphDefault({
        msal: {
          config: { auth: authCertificate },
          scopes: ["https://graph.microsoft.com/.default"],
        },
      })
    }
    throw new Error("Invalid authentication mode specified. Use 'sp' for SharePoint or 'graph' for Microsoft Graph.")
  }

  return {
    /** Returns a global SharePoint queryable object by authenticating using client ID / client Secret */
    sp: spfi(config.siteUrl).using(adaptiveAuth(config, "sp"), NoOdata) as SPFI,
    spOdata: spfi(config.siteUrl).using(adaptiveAuth(config, "sp")) as SPFI,

    /** Returns a global Graph queryable object by authenticating using client ID / client Secret */
    graph: graphfi().using(adaptiveAuth(config, "graph")) as GraphFI,
  }
}
