import { IExecuteFunctions, IHttpRequestOptions } from "n8n-workflow"
import type { ISPInvokableFactory, ISPQueryable, SPFI, SPInit } from "@pnp/sp/presets/all.js"
import type { GraphFI, GraphInit } from "@pnp/graph/presets/all.js"
import type { IGraphDefaultProps, ISPDefaultProps } from "@pnp/nodejs"
import type { TimelinePipe } from "@pnp/core"
import type { Queryable } from "@pnp/queryable"
import { ConfidentialClientApplication, Configuration } from "@azure/msal-node"

const getESMPnP = () => {
  return Promise.all([
    eval(`import("@pnp/sp/presets/all.js")`),
    eval(`import("@pnp/graph/presets/all.js")`),
    eval(`import("@pnp/nodejs/index.js")`),
    eval(`import("@pnp/queryable")`),
  ]).then(([sp, graph, nodejs, queryable]) => {
    return {
      spfi: sp.spfi as (root?: SPInit | SPFI) => SPFI,
      graphfi: graph.graphfi as (root?: GraphInit | GraphFI) => GraphFI,
      BearerToken: queryable.BearerToken as (token: string) => TimelinePipe<Queryable>,
      SPDefault: nodejs.SPDefault as (props: ISPDefaultProps) => TimelinePipe<Queryable>,
      SPDefaultHeaders: sp.DefaultHeaders as () => TimelinePipe,
      SPDefaultInit: sp.DefaultInit as () => TimelinePipe,
      GraphDefaultHeaders: graph.DefaultHeaders as () => TimelinePipe,
      GraphDefaultInit: graph.DefaultInit as () => TimelinePipe,
      NodeFetchWithRetry: nodejs.NodeFetchWithRetry as () => TimelinePipe,
      DefaultParse: queryable.DefaultParse as () => TimelinePipe,
      GraphDefault: nodejs.GraphDefault as (props: IGraphDefaultProps) => TimelinePipe<Queryable>,
      InjectHeaders: queryable.InjectHeaders as (headers: Record<string, string>, prepend?: boolean) => TimelinePipe,
      spQueryable: sp.SPQueryable as ISPInvokableFactory<ISPQueryable<any>>,
      spMethods: {
        get: sp.spGet as <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>,
        post: sp.spPost as <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>,
        delete: sp.spDelete as <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>,
        patch: sp.spPatch as <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>,
      },
    }
  })
}

export type IPnpConfig = {
  clientId: string
  clientSecret?: string
  tenantId?: string
  tenantName: string
  clientCertificatePrivateKey: string
  clientCertificateThumbprint: string
}

const getTenantId = async (
  tenantName: string,
  fetcher: (requestOptions: IHttpRequestOptions) => Promise<any> = (options) =>
    fetch(options.url, options as any).then((res) => res.json()),
) => {
  if (!/\.onmicrosoft\.com$/.test(tenantName)) tenantName += ".onmicrosoft.com"
  const response = await fetcher({
    url: `https://login.microsoftonline.com/${tenantName}/.well-known/openid-configuration`,
  })
  // console.log(`Tenant ID for ${tenantName} is ${response.issuer}`)
  // console.log(JSON.stringify(response, null, 2))
  return response.issuer.split("/")[3]
}

export async function getPnp(this: IExecuteFunctions, pnpConfig: IPnpConfig, siteUrl: string) {
  const {
    spfi,
    graphfi,
    InjectHeaders,
    spMethods,
    spQueryable,
    SPDefaultHeaders,
    SPDefaultInit,
    GraphDefaultHeaders,
    GraphDefaultInit,
    NodeFetchWithRetry,
    DefaultParse,
  } = await getESMPnP()
  const NoOdata = (instance: any) => InjectHeaders({ Accept: "application/json;odata=nometadata" })(instance)
  const tenantId = pnpConfig.tenantId || (await getTenantId(pnpConfig.tenantName, this?.helpers.httpRequest))

  const config: Configuration = {
    auth: {
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientId: pnpConfig.clientId,
      clientSecret: pnpConfig.clientSecret,
      clientCertificate: {
        privateKey: pnpConfig.clientCertificatePrivateKey,
        thumbprint: pnpConfig.clientCertificateThumbprint,
      },
    },
  }

  const CACHE: Record<
    string,
    Record<"sp" | "graph", { token?: string; expiresOn?: number }>
  > = (this?.getWorkflowStaticData("global") as Record<string, any>) || {}
  const CACHE_KEY = `n8n-custom-node-sharepoint-app-only-bearer-token-${siteUrl}`

  // console.log(Object.entries(CACHE))

  if (!CACHE[CACHE_KEY]) CACHE[CACHE_KEY] = { sp: {}, graph: {} }

  const BearerTokenFactory = (mode: "sp" | "graph", scopes: string[]) => {
    return (instance: Queryable<any>) => {
      instance.on.auth(async (url: URL, init: RequestInit) => {
        const model = CACHE[CACHE_KEY][mode]
        if (!model || !model.token || !model.expiresOn || model.expiresOn <= Date.now()) {
          console.log(`Acquiring new token for ${url.hostname}...`)
          const client = new ConfidentialClientApplication(config)
          const result = await client.acquireTokenByClientCredential({ scopes })
          if (!result) {
            throw new Error("Failed to acquire token")
          }
          model.token = result.accessToken
          model.expiresOn = result.expiresOn!.getTime()

          CACHE[CACHE_KEY][mode] = {
            token: model.token,
            expiresOn: model.expiresOn,
          }
        }

        init.headers = {
          ...init.headers,
          Authorization: `Bearer ${model.token}`,
        }

        return [url, init]
      })

      return instance
    }
  }

  const SPBearerToken = BearerTokenFactory("sp", [
    `https://${pnpConfig.tenantName.replace(".onmicrosoft.com", "")}.sharepoint.com/.default`,
  ])
  const GraphBearerToken = BearerTokenFactory("graph", ["https://graph.microsoft.com/.default"])

  const spcommonPipes = [SPDefaultHeaders(), SPDefaultInit(), NodeFetchWithRetry(), DefaultParse()]
  const graphcommonPipes = [NodeFetchWithRetry(), DefaultParse(), GraphDefaultHeaders(), GraphDefaultInit()]

  return {
    /** Returns a global SharePoint queryable object by authenticating using client ID / client Secret */
    sp: spfi(siteUrl).using(SPBearerToken, ...spcommonPipes, NoOdata) as SPFI,
    spOdata: spfi(siteUrl).using(SPBearerToken, ...spcommonPipes) as SPFI,
    spMethods: spMethods as {
      get: <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>
      post: <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>
      delete: <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>
      patch: <T = any>(o: ISPQueryable<any>, init?: RequestInit) => Promise<T>
    },
    spQueryable: spQueryable as ISPInvokableFactory<ISPQueryable<any>>,

    /** Returns a global Graph queryable object by authenticating using client ID / client Secret */
    graph: graphfi().using(GraphBearerToken, ...graphcommonPipes) as GraphFI,
  }
}
