export async function POST(request) {
  try {
    const { accessToken, mailbox, messageId } = await request.json();

    if (!accessToken || !mailbox || !messageId) {
      return Response.json({ error: "Invalid request" }, { status: 400 });
    }

    // Get message in MIME format
    const mimeUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${messageId}/$value`;

    const response = await fetch(mimeUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      return Response.json(
        { error: "Failed to fetch message" },
        { status: response.status },
      );
    }

    const mimeContent = await response.text();

    return new Response(mimeContent, {
      headers: {
        "Content-Type": "message/rfc822",
        "Content-Disposition": `attachment; filename="message-${messageId}.eml"`,
      },
    });
  } catch (error) {
    console.error("Download EML error:", error);
    return Response.json(
      { error: error.message || "Download failed" },
      { status: 500 },
    );
  }
}
