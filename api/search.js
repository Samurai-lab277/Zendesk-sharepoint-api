export default async function handler(req, res) {
  const { query } = req.body;

  // Step 1: Get Microsoft token
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        client_id: "YOUR_CLIENT_ID",
        client_secret: "YOUR_CLIENT_SECRET",
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default"
      })
    }
  );

  const tokenData = await tokenRes.json();

  // Step 2: Search SharePoint
  const searchRes = await fetch(
    "https://graph.microsoft.com/v1.0/search/query",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${tokenData.access_token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        requests: [
          {
            entityTypes: ["driveItem"],
            query: {
              queryString: query
            }
          }
        ]
      })
    }
  );

  const data = await searchRes.json();

  // Step 3: Clean results
  const results = data.value[0].hitsContainers[0].hits.map(hit => ({
    name: hit.resource.name,
    url: hit.resource.webUrl
  }));

  res.status(200).json({ results });
}
