import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

export default async function handler(req, res) {
  if (req.method === "GET" || req.method === "POST") {
    const report = req.method === "POST" ? req.body : req.query;
    try {
      const credential = new ClientSecretCredential(
        process.env.TENANT_ID,
        process.env.CLIENT_ID,
        process.env.CLIENT_SECRET
      );
      const graphClient = Client.initWithMiddleware({
        authProvider: { getAccessToken: async () => {
          const token = await credential.getToken("https://graph.microsoft.com/.default");
          return token.token;
        }}
      });
      await graphClient.api("/me/drive/root:/FraudReports/fraud-report.json:/content")
        .put(JSON.stringify(report, null, 2));
      await graphClient.api("/me/sendMail").post({
        message: {
          subject: "Fraud Alert Detected",
          body: { contentType: "Text", content: `Fraud detected at ${report.l || report.url}` },
          toRecipients: [{ emailAddress: { address: "your-email@example.com" } }]
        }
      });
      res.status(200).json({ status: "report saved and email sent" });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Graph API failed" });
    }
  } else {
    res.status(405).end();
  }
}
