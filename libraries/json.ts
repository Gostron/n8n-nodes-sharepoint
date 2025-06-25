import { IDataObject, IExecuteFunctions } from "n8n-workflow"

export function tryParse<T extends IDataObject>(
  this: IExecuteFunctions,
  parameterName: string,
  index: number,
  defaultValue: T,
): T | undefined {
  try {
    const value = this.getNodeParameter(parameterName, index, "{}") as string
    if (!value) return defaultValue
    return JSON.parse(value)
  } catch (error) {
    this.logger?.error(`Failed to parse JSON for ${parameterName}`, { error })
    return defaultValue
  }
}
