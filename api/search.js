export default async function handler(req, res) {
  const { query } = req.body;

  // Step 1: Get Microsoft token
  const tokenRes = await fetch(
    `https://login.microsoftonline.com/201f1044-2f66-4605-9418-05871a6d9d05/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        client_id: "9c36a2ff-4005-4f54-a620-7e58386de032",
        client_secret: "fVc8Q~JuXON7DtOKeUjHMkt5WMZLsPZvI2BFza~2",
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default"
      })
    }
  );

const tokenData = await tokenRes.json();

// 🔍 Debug token issues
if (!tokenData.access_token) {
  return res.status(500).json({
    error: "Token failed",
    details: tokenData
  });
}

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

// 🔍 Debug SharePoint response issues
if (!data.value || !data.value[0]?.hitsContainers?.[0]?.hits) {
  return res.status(200).json({
    results: [],
    debug: data
  });
}

// Step 3: Clean results
const results = data.value[0].hitsContainers[0].hits.map(hit => ({
  name: hit.resource.name,
  url: hit.resource.webUrl
}));

res.status(200).json({ results });
``
