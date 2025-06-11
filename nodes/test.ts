import { getPnp } from "../libraries/pnp/pnp.js"

export const main = async () => {
  const { sp } = await getPnp({})
  console.log(await sp.web())
}

main()
  .then(() => {
    console.log("Execution completed successfully.")
  })
  .catch((error) => {
    console.error("Execution failed:", error)
    process.exit(1)
  })
