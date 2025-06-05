import { getPnp } from "../libraries/pnp/pnp.js"

export const main = async () => {
  const { sp } = await getPnp({
    clientCertificatePrivateKey:
      "-----BEGIN RSA PRIVATE KEY-----\nMIIEpAIBAAKCAQEAmijpdPdoFkiyR7dMTHJnw6Sn/+JRUIEOiu39TngebCbmPNGG\nXwhQVpVK4G1d7FqhW/7orUy27NFQw8OsjAEc7L8HJpN4iVvb3JNXCMrh5wR3H/2z\nrNEUTG7FhFijKys2Wzb/vOb4VXbFEro3KYGQOiSuTErodHXmC7Q3iZZaCp2CPtJ4\n+IpZcxolPYRbPVHPrQRLAAnRfNSRsplxFn01BJ8KmB/r20zx+XcfDkEbc0oDQcu7\nslsKF0aIpAm0RKYa9jJ43OuEr8K2KAmR5sg9iQMkJmR2N7JbTlIV5GgKYHr6Zn27\n24HEi+Z/i25sqweu8Pi1dc6YFl0EN7tlt5cJJQIDAQABAoIBAGQVc5vQk+8Du1r8\nEbj//YXy/G8QS6JsZzijUfUD3xxwFMIfdZF1IFNWEYqq9nNgU6oaUI1SZOBS9JBk\nFT2/6zt4ufe9fmAPFyqZLcQzk34cVAqc922XQApvUCSgNy7rmxqVFmtZuJgjx8Zk\nxzNXPn/BGLfHWith77xhgDz/M+33WnSXHn3i4tE7LyfDHNSMhFzut3G/KcYR+ocb\n2aKZgWthJktgbJoNBmyfmKeUdhQA6ctPZv9BZ+3grEg643aFf4xUj54qhFOgtbTA\nE6ZOB68FGj/b4pTml7FRFylYvIIs4APg9stFp3XJfuQ30Q0y5GO+1i50Jk7zzLz8\njgmzunECgYEAwc8WW4r+BfBhTYxClvmOZUQtENSdZNbjCdd6jkyVd6vum/BQyRa0\naeIHvt4kB3JNsE0WKeAK5pokNXoMxakwMuMSXuQlZs/PIhJQMJJ5NBSBBshIjLYM\nRhcY3hFy9VWRT/DZjHNu/bEm0/ZIBRQH+avUefite0fDsmKPz6mHWa8CgYEAy6C/\n7E0fJ4P8TMKkxkOosiT0FMp3WoreonGAKPQA47QtHOvIKBJk+jUATJQOjkj//tcy\nLKDOD85ttqzkoFVGje7bGhNhC2032SwYfYV9z6PX24nsDWjaTfdSMligL9wYID0U\nAhZKjCiZeNCs4g90R11VXY9PVyy8b2f+ruDgg2sCgYATaU5v9MfkiGL2hWnV/UDJ\n2743xVPOWcd7oN1hi0IdLldDvxoYSfHf+QeVkmJBbK1jTxU9NHdjCWU/Be5pjbyK\nHDwmzOsCFSZF31ewxbrmAHe72iuKDGHGU2HmPBEriVp4i0L+0kD3n9qnuC8Wcx8p\nXpB1dvbJNjLflweYYP5xeQKBgQC24piddiLODdfTZVoyi9/+p/vklHegBWuyADi4\nD3ahDFkcSZQKkYLJykKLhMqR9nSdgM+aj46jWabmU+A/NHfa0DVO9SrK5XwsfFM+\noV9+10vu7K/q10qCjefwOfMdKRMuGU1YFoc73NcCUIGFw5WO5v/duPHsfMx9TqzK\nikR7IwKBgQCR8rrhcPkvNb5DA8scAmBOaW8glRu8xT7EBvKzX9sGHJ0LEs2JRQhw\n1toIMYw7vWTnaNvsSh7fHs6UH8SMYFu0zy/0Tcci/W2qv2pzgRu0u+kQskZKDfOY\neS3qtZN9nt64mjKGBJk9NfpmmNAIhSKJQ1+W/iO1xHBJV2no71PxIA==\n-----END RSA PRIVATE KEY-----",
    clientCertificateThumbprint: "A35B34D08CAB776AA5ADB675CCC970F559AAE76B",
    clientId: "c87cd543-a5e5-49de-84a4-45beb9819952",
    tenantId: "fd1fd5b5-b277-4005-86ab-70d2315c6cd2",
    siteUrl: "https://asigroupe.sharepoint.com/sites/MOCA",
  })
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
