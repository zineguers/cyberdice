export async function POST(request) {
  try {
    const { accessToken, mailbox, messageId } = await request.json();

    if (!accessToken || !mailbox || !messageId) {
      return Response.json({ error: "Invalid request" }, { status: 400 });
    }

    // Get attachments
    const attachmentsUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${messageId}/attachments`;

    const response = await fetch(attachmentsUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      return Response.json(
        { error: "Failed to fetch attachments" },
        { status: response.status },
      );
    }

    const data = await response.json();

    return Response.json({
      success: true,
      attachments: data.value || [],
    });
  } catch (error) {
    console.error("Attachments error:", error);
    return Response.json(
      { error: error.message || "Failed to fetch attachments" },
      { status: 500 },
    );
  }
}
