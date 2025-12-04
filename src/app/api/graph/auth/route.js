export async function POST(request) {
  try {
    const { tenantId, clientId, clientSecret } = await request.json();

    if (!tenantId || !clientId || !clientSecret) {
      return Response.json(
        { error: "Missing required fields" },
        { status: 400 },
      );
    }

    // Real Microsoft Graph authentication
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const params = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: params.toString(),
    });

    const data = await response.json();

    if (!response.ok) {
      // Return actual Microsoft error
      return Response.json(
        {
          error:
            data.error_description || data.error || "Authentication failed",
          errorCode: data.error,
        },
        { status: response.status },
      );
    }

    return Response.json({
      success: true,
      accessToken: data.access_token,
      expiresIn: data.expires_in,
    });
  } catch (error) {
    console.error("Auth error:", error);
    return Response.json(
      { error: error.message || "Authentication failed" },
      { status: 500 },
    );
  }
}
