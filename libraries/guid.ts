export const isGuid = (value: string): boolean => {
  const regex = /^[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}$/
  return regex.test(value)
}
