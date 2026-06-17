export default async function handler(req, res) {
  const secret = req.headers["x-api-key"];
  if (secret !== process.env.API_SECRET) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const query = req.body?.query || "manual";

const tokenRes = await fetch(
  `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
  {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: "client_credentials",
      scope: "https://graph.microsoft.com/.default"
    })
  }
);

    const tokenData = await tokenRes.json();

    if (!tokenData.access_token) {
      return res.status(500).json({ error: "Token failed", details: tokenData });
    }

    const searchRes = await fetch(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tokenData.access_token}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          requests: [{
            entityTypes: ["driveItem"],
            query: { queryString: query },
            contentSources: [`/sites/${process.env.SP_SITE_NAME}`]
          }]
        })
      }
    );

    const data = await searchRes.json();

    const hits = data.value?.[0]?.hitsContainers?.[0]?.hits;
    if (!hits) {
      return res.status(200).json({ results: [] });
    }

    const results = hits.map((hit) => ({
      name: hit.resource.name,
      url: hit.resource.webUrl
    }));

    return res.status(200).json({ results });

  } catch (err) {
    return res.status(500).json({ error: "Crash", message: err.message });
  }
}
