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
    },
    {
      displayName: "Client Certificate Private Key",
      name: "clientCertificatePrivateKey",
      type: "string",
      typeOptions: {
        password: true,
      },
      default: "",
    },
    {
      displayName: "Client Certificate Thumbprint",
      name: "clientCertificateThumbprint",
      type: "string",
      default: "",
    },
    {
      displayName: "Tenant ID",
      name: "tenantId",
      type: "string",
      default: "",
    },
    {
      displayName: "Site URL",
      name: "siteUrl",
      type: "string",
      default: "",
      placeholder: "https://<your-tenant-name>.sharepoint.com/sites/<your-site-name>",
    },
  ]
}
