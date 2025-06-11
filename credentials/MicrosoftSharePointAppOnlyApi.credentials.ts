import { ICredentialType, INodeProperties } from "n8n-workflow"

export class MicrosoftSharePointAppOnlyApi implements ICredentialType {
  name = "microsoftSharePointAppOnlyApi"
  displayName = "Microsoft SharePoint App Only API"

  documentationUrl = "https://learn.microsoft.com/en-us/sharepoint/auth/auth-concepts"

  properties: INodeProperties[] = [
    {
      displayName: "Client ID",
      name: "clientId",
      type: "string",
      default: "",
      placeholder: "e.g., 12345678-1234-1234-1234-123456789012",
      description:
        "The App ID of the Entra ID Application (Client) used for authentication. This is a GUID that uniquely identifies your application in Entra ID.",
    },
    {
      displayName: "Client Certificate Private Key",
      name: "clientCertificatePrivateKey",
      type: "string",
      typeOptions: {
        password: true,
      },
      default: "",
      placeholder: "e.g., -----BEGIN RSA PRIVATE KEY-----\nMIIEpAIBAAKCAQEA...\n-----END RSA PRIVATE KEY-----",
      description:
        "The private key of the client certificate used for authentication. It should be in PEM format. You can remove header, footer and line returns if needed.",
    },
    {
      displayName: "Site URL",
      name: "siteUrl",
      type: "string",
      default: "",
    },
    {
      displayName: "Client Certificate Thumbprint",
      name: "clientCertificateThumbprint",
      type: "string",
      default: "",
      placeholder: "e.g., 1234567890ABCDEF1234567890ABCDEF12345678",
    },
    {
      displayName: "Tenant ID",
      name: "tenantId",
      type: "string",
      default: "",
      placeholder: "e.g., 12345678-1234-1234-1234-123456789012",
      description:
        "The ID of the Microsoft 365 tenant. This is a GUID that uniquely identifies your tenant. You can find it on https://whatismytenantid.com or in the Azure portal.",
    },
    {
      displayName: "Tenant name",
      name: "tenantName",
      type: "string",
      default: "",
      placeholder: "contoso",
      description:
        "The name of the tenant, used to construct the site URL (e.g., for a SharePoint URL https://contoso.sharepoint.com, this would be 'contoso').",
    },
  ]
}
