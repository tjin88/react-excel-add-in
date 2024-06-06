# Please place the server certificate in this folder ***
**Server certificate is required for the server to run**
**You must also create a .env file containing an OpenAI api key**

# Implementing API Calls on the Excel Add-in

### Generating self-signed cert
openssl genrsa -out server.key 2048
openssl req -new -key server.key -out server.csr
openssl x509 -req -in server.csr -signkey server.key -out server.crt

### Trusting the cert
Mac:
- Applications > Utilities > Keychain Access
- Drag/Drop server.crt file into **System**
- Expand *Trust* section
- Set to *Always Trust*

Windows:
- Open *Certificate Manager*
- Press *Win + R* to open the Run dialog
- Type "mmc" and press Enter --> This will open the Microsoft Management Console
- In the MMC, go to File > Add/Remove Snap-in
- Choose Computer account and click Next. Select Local computer and click Finish
- In the MMC, expand Certificates (Local Computer).
- Right-click on the Trusted Root Certification Authorities store.
- Select All Tasks > Import.
- Follow the wizard to import the server.crt file.
- Make sure to place the certificate in the Trusted Root Certification Authorities store.
    - If you want to verify the cert is trusted, you should be able to find it in *Trusted Root Certification Authorities > Certificates*
