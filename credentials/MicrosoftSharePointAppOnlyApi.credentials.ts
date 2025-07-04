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
      required: true,
      description:
        "The App ID of the Entra ID Application (Client) used for authentication. This is a GUID that uniquely identifies your application in Entra ID.",
    },
    {
      displayName: "Client Secret",
      name: "clientSecret",
      type: "string",
      typeOptions: {
        password: true,
      },
      default: "",
      placeholder: "e.g., 12345678-1234-1234-1234-123456789012",
      required: false,
      description:
        "The secret key of the Entra ID Application (Client) used for authentication. It is used in conjunction with the Client ID to authenticate your application.",
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
        "The private key of the client certificate used for authentication. It should be in PEM format. Header, footer and line returns are ignored.",
    },
    {
      displayName: "Client Certificate Thumbprint",
      name: "clientCertificateThumbprint",
      type: "string",
      default: "",
      placeholder: "e.g., 1234567890ABCDEF1234567890ABCDEF12345678",
    },
    {
      displayName: "Tenant Name",
      name: "tenantName",
      type: "string",
      default: "",
      required: true,
      placeholder: "e.g., contoso",
      description:
        "The name of the Microsoft 365 tenant. This is typically in the format 'yourtenant.onmicrosoft.com'. It is used to identify your tenant in the Microsoft cloud.",
    },
    {
      displayName: "Tenant ID",
      name: "tenantId",
      type: "string",
      default: "",
      required: false,
      placeholder: "e.g., 12345678-1234-1234-1234-123456789012",
      description:
        "The ID of the Microsoft 365 tenant. This is a GUID that uniquely identifies your tenant. You can find it on https://whatismytenantid.com or in the Azure portal.",
    },
  ]
}
