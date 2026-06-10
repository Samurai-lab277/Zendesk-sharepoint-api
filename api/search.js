export default async function handler(req, res) {
  try {
    const query = req.body?.query || "manual";

    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded"
        },
        body: new URLSearchParams({
          client_id: process.env.CLIENT_ID,
          client_secret: process.env.CLIENT_SECRET,
          grant_type: "client_credentials",
          scope: "https://graph.microsoft.com/.default"
        })
      }
    );

    const tokenText = await tokenRes.text();

    let tokenData;
    try {
      tokenData = JSON.parse(tokenText);
    } catch {
      return res.status(500).json({
        error: "Token parse failed",
        raw: tokenText
      });
    }

    if (!tokenData.access_token) {
      return res.status(500).json({
        error: "Token failed",
        details: tokenData
      });
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

    const searchText = await searchRes.text();

    let data;
    try {
      data = JSON.parse(searchText);
    } catch {
      return res.status(500).json({
        error: "Search parse failed",
        raw: searchText
      });
    }

    if (!data.value || !data.value[0]?.hitsContainers?.[0]?.hits) {
      return res.status(200).json({
        results: [],
        debug: data
      });
    }

    const results = data.value[0].hitsContainers[0].hits.map((hit) => ({
      name: hit.resource.name,
      url: hit.resource.webUrl
    }));

    return res.status(200).json({ results });

  } catch (err) {
    return res.status(500).json({
      error: "Crash",
      message: err.message,
      stack: err.stack
    });
  }
}
