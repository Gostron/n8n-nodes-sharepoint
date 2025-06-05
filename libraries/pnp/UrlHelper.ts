import { parseURL } from "whatwg-url"

export enum HostingEnvironment {
  Production = 0,
  German = 1,
  China = 2,
  USGovernment = 3,
  USDefence = 4,
}

export class UrlHelper {
  public static removeTrailingSlash(url: string): string {
    return url.replace(/(\/$)|(\\$)/, "")
  }

  public static removeLeadingSlash(url: string): string {
    return url.replace(/(^\/)|(^\\)/, "")
  }

  public static trimSlashes(url: string): string {
    return url.replace(/(^\/)|(^\\)|(\/$)|(\\$)/g, "")
  }

  public static ResolveHostingEnvironment(siteUrl: string): HostingEnvironment {
    const host = parseURL(siteUrl)?.host! as string

    if (host.indexOf(".sharepoint.com") !== -1) {
      return HostingEnvironment.Production
    } else if (host.indexOf(".sharepoint.cn") !== -1) {
      return HostingEnvironment.China
    } else if (host.indexOf(".sharepoint.de") !== -1) {
      return HostingEnvironment.German
    } else if (host.indexOf(".sharepoint-mil.us") !== -1) {
      return HostingEnvironment.USDefence
    } else if (host.indexOf(".sharepoint.us") !== -1) {
      return HostingEnvironment.USGovernment
    }

    return HostingEnvironment.Production // As default, for O365 Dedicated, #ToInvestigate
    // throw new Error('Unable to resolve hosting environment. Site url: ' + siteUrl)
  }
}
