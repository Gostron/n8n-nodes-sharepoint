declare type PnpDescriptionNode = {
  label: string
  key: string
  description?: string
  callable?: boolean
  /** If parameters, is a callable function with parameters */
  parameters?: {
    label: string
    key: string
    description?: string
    type: "string" | "number" | "boolean" | "object"
    required?: boolean
    /** If `true`, allows multiple values for this parameter */
    spread?: boolean
    /** If `true`, will be injected as a chain method and not as a parameter */
    chain?: boolean
  }[]
  /** If not children, is a leaf node */
  children?: PnpDescriptionNode[]
}
