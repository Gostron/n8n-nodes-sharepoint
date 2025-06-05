import { parseURL } from "whatwg-url"
import { Cache } from "./Cache.js"
import fetch from "node-fetch"
import { HostingEnvironment, UrlHelper } from "./UrlHelper.js"

export const SharePointServicePrincipal = "00000003-0000-0ff1-ce00-000000000000"
export const HighTrustTokenLifeTime: number = 12 * 60 * 60
export const FbaAuthEndpoint = "_vti_bin/authentication.asmx"
export const TmgAuthEndpoint = "CookieAuth.dll?Logon"
export const FormsPath = "_forms/default.aspx?wa=wsignin1.0"
export const RtFa = "rtFa"
export const FedAuth = "FedAuth"
export const AdfsOnlineRealm = "urn:federation:MicrosoftOnline"

export interface IAuthResponse {
  headers: { [key: string]: any }
  options?: { [key: string]: any }
}

export class OnlineAddinOnly {
  private hostingEnvironment: HostingEnvironment
  private endpointsMappings: Map<HostingEnvironment, string>

  private static TokenCache: Cache = new Cache()

  constructor(private _siteUrl: string, private _authOptions: { clientId: string; clientSecret: string }) {
    this.endpointsMappings = new Map()
    this.hostingEnvironment = UrlHelper.ResolveHostingEnvironment(this._siteUrl)
    this.InitEndpointsMappings()
  }

  public getAuth(): Promise<IAuthResponse> {
    const sharepointhostname = parseURL(this._siteUrl)?.host || ""
    const cacheKey = `${sharepointhostname}@${this._authOptions.clientSecret}@${this._authOptions.clientId}`

    const cachedToken: string = OnlineAddinOnly.TokenCache.get<string>(cacheKey)!

    if (cachedToken) {
      return Promise.resolve({
        headers: {
          Authorization: `Bearer ${cachedToken}`,
        },
      })
    }

    return this.getRealm(this._siteUrl)
      .then((realm) => {
        return Promise.all([realm, this.getAuthUrl(realm)])
      })
      .then((data) => {
        const realm: string = data[0]
        const authUrl: string = data[1]

        const resource = `${SharePointServicePrincipal}/${sharepointhostname}@${realm}`
        const fullClientId = `${this._authOptions.clientId}@${realm}`

        const body = new FormData()
        body.set("grant_type", "client_credentials")
        body.set("client_id", fullClientId)
        body.set("client_secret", this._authOptions.clientSecret)
        body.set("resource", resource)

        // console.log({
        //   grant_type: "client_credentials",
        //   client_id: fullClientId,
        //   client_secret: this._authOptions.clientSecret,
        //   resource: resource,
        //   authUrl: authUrl,
        // })

        return fetch(authUrl, {
          method: "POST",
          body: body,
        }).then((r) => r.json() as Promise<{ expires_in: string; access_token: string }>)
      })
      .then((data) => {
        const expiration: number = parseInt(data.expires_in, 10)
        OnlineAddinOnly.TokenCache.set(cacheKey, data.access_token, expiration - 60)

        return {
          headers: {
            Authorization: `Bearer ${data.access_token}`,
          },
        }
      })
  }

  protected InitEndpointsMappings(): void {
    this.endpointsMappings.set(HostingEnvironment.Production, "accounts.accesscontrol.windows.net")
    this.endpointsMappings.set(HostingEnvironment.China, "accounts.accesscontrol.chinacloudapi.cn")
    this.endpointsMappings.set(HostingEnvironment.German, "login.microsoftonline.de")
    this.endpointsMappings.set(HostingEnvironment.USDefence, "accounts.accesscontrol.windows.net")
    this.endpointsMappings.set(HostingEnvironment.USGovernment, "accounts.accesscontrol.windows.net")
  }

  private getAuthUrl(realm: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const url = this.AcsRealmUrl + realm

      fetch(url)
        .then((r) => r.json() as Promise<any>)
        .then((data: { endpoints: { protocol: string; location: string }[] }) => {
          for (let i = 0; i < data.endpoints.length; i++) {
            if (data.endpoints[i].protocol === "OAuth2") {
              resolve(data.endpoints[i].location)
              return undefined
            }
          }
        })
        .catch(reject)
    })
  }

  private get AcsRealmUrl(): string {
    return `https://${this.endpointsMappings.get(this.hostingEnvironment)}/metadata/json/1?realm=`
  }

  private getRealm(siteUrl: string): Promise<string> {
    // if (this._authOptions.realm) {
    //   return Promise.resolve(this._authOptions.realm)
    // }

    return fetch(`${UrlHelper.removeTrailingSlash(siteUrl)}/_vti_bin/client.svc`, {
      method: "POST",
      headers: { Authorization: "Bearer " },
    }).then((data) => {
      const header: string = data.headers.get("www-authenticate")
      const index: number = header.indexOf('Bearer realm="')
      return header.substring(index + 14, index + 50)
    })
  }
}
