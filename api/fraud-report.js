export default function handler(req, res) {
  if (req.method === "POST") {
    const report = req.body;
    console.log("Fraud report received:", report);

    // TODO: Here we will later add Microsoft Graph code
    // to push the file into OneDrive and send email.

    res.status(200).json({ status: "received" });
  } else {
    res.status(405).end();
  }
}
